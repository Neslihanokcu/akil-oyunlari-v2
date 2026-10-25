# -*- coding: utf-8 -*-
"""
Akıl dan Zeka Oyunları Yönetim ve Takip Paneli (Orijinal Kararlı Sürüm)
"""

import hashlib
import hmac
import html
import io
import math
import os
import re
import secrets
import shutil
import sqlite3
import time
from contextlib import contextmanager
from datetime import date, datetime

import pandas as pd
import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "akil_oyunlari.db")
DEFAULT_PIN = "6991"
PAIR_CATEGORY = "İkili Oyunlar"
INDIV_CATEGORY = "Bireysel Oyunlar"
GROUP_CATEGORY = "Toplu Oyunlar"

GAME_LIBRARY = {
    "İkili Oyunlar": ["Mangala", "Goon", "9 Taş", "Küre", "Pentego", "Abluka", "Surakarta", "Renk Kat"],
    "Bireysel Oyunlar": ["Şakuli", "Motif", "Equilibrio", "9/24"],
    "Toplu Oyunlar": ["The Road Game", "Renkli Bardaklar", "Baybak", "Dik Go"],
}
CATEGORY_ICONS = {"İkili Oyunlar": "⚔️", "Bireysel Oyunlar": "🧩", "Toplu Oyunlar": "👥"}

PG_HOME = "Genel Bakış"
PG_BRANCH = "Şube Paneli"
PG_MATCH = "Maç Girişi"
PG_BOARD = "Sıralama ve Geçmiş"
PG_INDIV = "Bireysel Oyunlar"
PG_GROUP = "Grup ve Turnuva"
PG_LIB = "Oyun Kütüphanesi"
PG_ADMIN = "Yönetici Paneli"

PAGE_ICONS = {
    PG_HOME: "home", PG_BRANCH: "school", PG_MATCH: "swords", PG_BOARD: "trophy",
    PG_INDIV: "puzzle", PG_GROUP: "users", PG_LIB: "dice", PG_ADMIN: "lock",
}

esc = html.escape
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NO_AUTH_MSG = "Bu işlem için yönetici girişi gerekir."
MAX_PIN_FAILS = 5      
LOCK_SECONDS = 300     
PAGE_SIZE = 20         
PROG_KAYNAK = {"manuel": "Manuel", "gecis": "Geçiş Verisi"}

def tr_upper(s):
    return str(s).replace("i", "İ").replace("ı", "I").upper()

_TR_ALPHA = "ABCÇDEFGHIİJKLMNOÖPRSŞTUÜVYZ"

def tr_key(s):
    return tuple(_TR_ALPHA.find(c) if c in _TR_ALPHA else 100 + ord(c) for c in tr_upper(s))

def _nz(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v

def _to_int(v):
    v = _nz(v)
    if v is None:
        return None
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None

def prev_school_year():
    t = date.today()
    start = t.year if t.month >= 9 else t.year - 1
    return f"{start - 1}-{start}"

def _norm(s):
    s = str(s).strip()
    for a, b in (("İ", "i"), ("I", "i"), ("ı", "i")):
        s = s.replace(a, b)
    s = s.lower()
    for a, b in zip("şğüöç", "sguoc"):
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9]", "", s)

@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def _migrate(c):
    cols = {r["name"] for r in c.execute("PRAGMA table_info(students)")}
    if "arsiv" not in cols:
        c.execute("ALTER TABLE students ADD COLUMN arsiv INTEGER NOT NULL DEFAULT 0")
    if "arsiv_tarih" not in cols:
        c.execute("ALTER TABLE students ADD COLUMN arsiv_tarih TEXT")
    if "arsiv_notu" not in cols:
        c.execute("ALTER TABLE students ADD COLUMN arsiv_notu TEXT")
    c.execute("CREATE INDEX IF NOT EXISTS idx_students_arsiv ON students(arsiv)")

def init_db():
    with db() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT, ogrenci_no TEXT UNIQUE,
                ad_soyad TEXT NOT NULL, sinif INTEGER NOT NULL, sube TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS games (
                id INTEGER PRIMARY KEY AUTOINCREMENT, ad TEXT NOT NULL UNIQUE,
                kategori TEXT NOT NULL, varsayilan INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS matches (
                id INTEGER PRIMARY KEY AUTOINCREMENT, tarih TEXT NOT NULL,
                game_id INTEGER NOT NULL REFERENCES games(id) ON DELETE CASCADE,
                p1_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                p2_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                sonuc TEXT NOT NULL CHECK (sonuc IN ('1','2','D')),
                CHECK (p1_id <> p2_id)
            );
            CREATE TABLE IF NOT EXISTS progress (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                game_id INTEGER NOT NULL REFERENCES games(id) ON DELETE CASCADE,
                kart INTEGER, seviye TEXT, tarih TEXT NOT NULL, kaynak TEXT NOT NULL DEFAULT 'manuel', donem TEXT
            );
            CREATE TABLE IF NOT EXISTS game_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT, ad TEXT NOT NULL,
                game_id INTEGER NOT NULL REFERENCES games(id) ON DELETE CASCADE,
                turnuva TEXT NOT NULL DEFAULT '', tarih TEXT NOT NULL, UNIQUE (game_id, turnuva, ad)
            );
            CREATE TABLE IF NOT EXISTS group_members (
                group_id INTEGER NOT NULL REFERENCES game_groups(id) ON DELETE CASCADE,
                student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                PRIMARY KEY (group_id, student_id)
            );
            CREATE TABLE IF NOT EXISTS group_scores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL REFERENCES game_groups(id) ON DELETE CASCADE,
                student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                tarih TEXT NOT NULL, tur TEXT NOT NULL DEFAULT '', kart INTEGER NOT NULL DEFAULT 0, puan INTEGER NOT NULL DEFAULT 0
            );
            """
        )
        _migrate(c)
        for kategori, oyunlar in GAME_LIBRARY.items():
            for ad in oyunlar:
                c.execute("INSERT OR IGNORE INTO games (ad, kategori, varsayilan) VALUES (?,?,1)", (ad, kategori))
    if _get_setting("admin_pin") is None:
        _store_pin(DEFAULT_PIN)

def _get_setting(key):
    with db() as c:
        row = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else None

def _set_setting(key, value):
    with db() as c:
        c.execute("INSERT INTO settings (key, value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))

def is_admin():
    return bool(st.session_state.get("admin"))

def require_admin():
    if not is_admin():
        raise PermissionError(NO_AUTH_MSG)

def admin_logout():
    st.session_state["admin"] = False
    st.session_state.pop("del_pending", None)

def _hash_pin(pin, salt_hex):
    return hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), bytes.fromhex(salt_hex), 120_000).hex()

def _store_pin(pin):
    salt = secrets.token_hex(16)
    _set_setting("admin_pin", f"{salt}${_hash_pin(pin, salt)}")

def set_pin(pin):
    require_admin()
    _store_pin(pin)

def verify_pin(pin):
    stored = _get_setting("admin_pin")
    if not stored or "$" not in stored:
        return False
    salt, h = stored.split("$", 1)
    return hmac.compare_digest(_hash_pin(pin, salt), h)

def _int_setting(key):
    try:
        return int(float(_get_setting(key) or 0))
    except (TypeError, ValueError):
        return 0

def lock_remaining():
    return max(0, _int_setting("pin_lock_until") - int(time.time()))

def register_pin_failure():
    n = _int_setting("pin_fail_count") + 1
    if n >= MAX_PIN_FAILS:
        _set_setting("pin_lock_until", str(int(time.time()) + LOCK_SECONDS))
        n = 0
    _set_setting("pin_fail_count", str(n))
    return n

def reset_pin_failures():
    _set_setting("pin_fail_count", "0")
    _set_setting("pin_lock_until", "0")

def _sort_students(df):
    if df.empty:
        return df
    order = sorted(range(len(df)), key=lambda i: (int(df.sinif.iloc[i]), df.sube.iloc[i], tr_key(df.ad_soyad.iloc[i])))
    return df.iloc[order].reset_index(drop=True)

def load_students(durum="aktif"):
    where = {"aktif": "WHERE arsiv=0", "arsiv": "WHERE arsiv=1"}.get(durum, "")
    with db() as c:
        df = pd.read_sql_query(f"SELECT id, ogrenci_no, ad_soyad, sinif, sube, arsiv, arsiv_tarih, arsiv_notu FROM students {where}", c)
    return _sort_students(df)

def add_student(ogrenci_no, ad_soyad, sinif, sube):
    require_admin()
    ad_soyad = re.sub(r"\s+", " ", ad_soyad).strip()
    sube = sube.strip().upper()
    try:
        with db() as c:
            cur = c.execute("INSERT INTO students (ogrenci_no, ad_soyad, sinif, sube) VALUES (?,?,?,?)", (ogrenci_no.strip() or None, ad_soyad, int(sinif), sube))
            if not ogrenci_no.strip():
                c.execute("UPDATE students SET ogrenci_no=? WHERE id=?", (f"OG{cur.lastrowid:04d}", cur.lastrowid))
        return True, f"{ad_soyad} eklendi."
    except sqlite3.IntegrityError:
        return False, "Bu öğrenci numarası zaten kayıtlı."

def update_student(sid, ogrenci_no, ad_soyad, sinif, sube):
    require_admin()
    try:
        with db() as c:
            c.execute("UPDATE students SET ogrenci_no=?, ad_soyad=?, sinif=?, sube=? WHERE id=?", (ogrenci_no.strip(), ad_soyad, int(sinif), sube, int(sid)))
        return True, "Güncellendi."
    except sqlite3.IntegrityError:
        return False, "Bu numara başka bir öğrenciye ait."

def delete_student(sid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM students WHERE id=?", (int(sid),))

_ALIAS_ID = {"ogrenciid", "ogrencino", "ogrno", "okulno", "no", "id", "numara"}
_ALIAS_NAME = {"adsoyad", "adisoyadi", "adsoyadi", "isim", "ogrenci"}
_ALIAS_SINIF = {"sinif", "sinifi", "seviye"}
_ALIAS_SUBE = {"sube", "subesi"}

def _read_uploaded_table(uploaded):
    name = uploaded.name.lower()
    data = uploaded.getvalue()
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "cp1254", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding=enc, dtype=str)
            except Exception:
                continue
    return pd.read_excel(io.BytesIO(data), dtype=str)

def parse_student_file(uploaded):
    raw = _read_uploaded_table(uploaded)
    cols = {}
    for col in raw.columns:
        n = _norm(col)
        for key, alias in (("id", _ALIAS_ID), ("name", _ALIAS_NAME), ("sinif", _ALIAS_SINIF), ("sube", _ALIAS_SUBE)):
            if n in alias and key not in cols:
                cols[key] = col
    out = pd.DataFrame()
    out["ad_soyad"] = raw[cols["name"]].fillna("").astype(str).str.strip()
    out["sinif"] = pd.to_numeric(raw[cols["sinif"]].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    out["sube"] = raw[cols["sube"]].fillna("").astype(str).str.strip().map(tr_upper)
    out["ogrenci_no"] = raw[cols["id"]].fillna("").astype(str).str.strip() if "id" in cols else ""
    out = out[out.ad_soyad != ""].copy()
    out["sinif"] = out.sinif.astype(int)
    return out, []

def import_students(df, replace=False):
    require_admin()
    added = updated = skipped = arch = 0
    with db() as c:
        if replace:
            c.execute("DELETE FROM students")
        for r in df.itertuples(index=False):
            no = (r.ogrenci_no or "").strip()
            if no:
                row = c.execute("SELECT id FROM students WHERE ogrenci_no=?", (no,)).fetchone()
                if row:
                    c.execute("UPDATE students SET ad_soyad=?, sinif=?, sube=? WHERE id=?", (r.ad_soyad, int(r.sinif), r.sube, row["id"]))
                    updated += 1
                    continue
                c.execute("INSERT INTO students (ogrenci_no, ad_soyad, sinif, sube) VALUES (?,?,?,?)", (no, r.ad_soyad, int(r.sinif), r.sube))
                added += 1
            else:
                cur = c.execute("INSERT INTO students (ad_soyad, sinif, sube) VALUES (?,?,?)", (r.ad_soyad, int(r.sinif), r.sube))
                c.execute("UPDATE students SET ogrenci_no=? WHERE id=?", (f"OG{cur.lastrowid:04d}", cur.lastrowid))
                added += 1
    return added, updated, skipped, arch

def load_games():
    with db() as c:
        return pd.read_sql_query("SELECT id, ad, kategori, varsayilan FROM games ORDER BY kategori, id", c)

def add_game(ad, kategori):
    require_admin()
    try:
        with db() as c:
            c.execute("INSERT INTO games (ad, kategori, varsayilan) VALUES (?,?,0)", (ad.strip(), kategori))
        return True, "Oyun eklendi."
    except sqlite3.IntegrityError:
        return False, "Bu oyun zaten var."

def delete_game(gid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM games WHERE id=? AND varsayilan=0", (int(gid),))

def add_match(game_id, p1, p2, sonuc, tarih=None):
    require_admin()
    if tarih is None:
        tarih = date.today()
    with db() as c:
        c.execute("INSERT INTO matches (tarih, game_id, p1_id, p2_id, sonuc) VALUES (?,?,?,?,?)", (tarih.strftime("%Y-%m-%d"), int(game_id), int(p1), int(p2), sonuc))
    return True, "Maç kaydedildi."

def fmt_tarih(v):
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        return str(v)

def delete_match(mid):
    require_admin()
    with db() as c:
        cur = c.execute("DELETE FROM matches WHERE id=?", (int(mid),))
        return cur.rowcount > 0

def delete_all_matches():
    require_admin()
    with db() as c:
        c.execute("DELETE FROM matches")

def load_matches(game_id=None, active_only=False):
    sql = "SELECT id, tarih, game_id, p1_id, p2_id, sonuc FROM matches"
    conds, params = [], []
    if game_id is not None:
        conds.append("game_id=?")
        params.append(int(game_id))
    if active_only:
        conds.append("(p1_id IN (SELECT id FROM students WHERE arsiv=0) OR p2_id IN (SELECT id FROM students WHERE arsiv=0))")
    if conds:
        sql += " WHERE " + " AND ".join(conds)
    with db() as c:
        return pd.read_sql_query(sql + " ORDER BY id DESC", c, params=params)

def match_history(game_id=None, newest_first=True, include_archived=False):
    sql = (
        "SELECT m.id AS id, m.tarih AS tarih, g.ad AS oyun, "
        "s1.ad_soyad || CASE WHEN s1.arsiv=1 THEN ' (arşiv)' ELSE '' END AS oyuncu1, "
        "(s1.sinif || '-' || s1.sube) AS sube1, "
        "s2.ad_soyad || CASE WHEN s2.arsiv=1 THEN ' (arşiv)' ELSE '' END AS oyuncu2, "
        "(s2.sinif || '-' || s2.sube) AS sube2, m.sonuc AS sonuc "
        "FROM matches m JOIN games g ON g.id=m.game_id "
        "JOIN students s1 ON s1.id=m.p1_id JOIN students s2 ON s2.id=m.p2_id"
    )
    conds, params = [], []
    if game_id is not None:
        conds.append("m.game_id=?")
        params.append(int(game_id))
    if not include_archived:
        conds.append("(s1.arsiv=0 OR s2.arsiv=0)")
    if conds:
        sql += " WHERE " + " AND ".join(conds)
    direction = "DESC" if newest_first else "ASC"
    with db() as c:
        return pd.read_sql_query(sql + f" ORDER BY substr(m.tarih, 1, 10) {direction}, m.id {direction}", c, params=params)

def played_pairs(game_id):
    m = load_matches(game_id)
    return {frozenset((int(a), int(b))) for a, b in zip(m.p1_id, m.p2_id)}

def archive_students(ids, note=""):
    require_admin()
    stamp = date.today().strftime("%Y-%m-%d")
    with db() as c:
        cur = c.executemany("UPDATE students SET arsiv=1, arsiv_tarih=?, arsiv_notu=? WHERE id=? AND arsiv=0", [(stamp, note or None, i) for i in ids])
        return cur.rowcount

def restore_students(ids):
    require_admin()
    with db() as c:
        cur = c.executemany("UPDATE students SET arsiv=0, arsiv_tarih=NULL, arsiv_notu=NULL WHERE id=? AND arsiv=1", [(i,) for i in ids])
        return cur.rowcount

def count_archived():
    with db() as c:
        return int(c.execute("SELECT COUNT(*) FROM students WHERE arsiv=1").fetchone()[0])

def archived_data_summary():
    with db() as c:
        return pd.read_sql_query(
            "SELECT s.id AS id, "
            "(SELECT COUNT(*) FROM matches m WHERE m.p1_id=s.id OR m.p2_id=s.id) AS mac, "
            "(SELECT COUNT(*) FROM progress p WHERE p.student_id=s.id) AS ilerleme, "
            "(SELECT COUNT(*) FROM group_scores g WHERE g.student_id=s.id) AS grup_skoru "
            "FROM students s WHERE s.arsiv=1", c,
        )

def save_progress(student_id, game_id, kart, seviye, tarih=None, kaynak="manuel", donem=None):
    require_admin()
    if tarih is None:
        tarih = date.today()
    with db() as c:
        c.execute("INSERT INTO progress (student_id, game_id, kart, seviye, tarih, kaynak, donem) VALUES (?,?,?,?,?,?,?)", (int(student_id), int(game_id), _to_int(kart), _nz(seviye), tarih.strftime("%Y-%m-%d"), kaynak, donem))
    return True, "Kaydedildi."

def load_progress(game_id=None, student_id=None):
    sql = "SELECT p.id AS id, p.student_id AS student_id, p.game_id AS game_id, g.ad AS oyun, g.kategori AS kategori, p.kart AS kart, p.seviye AS seviye, p.tarih AS tarih, p.kaynak AS kaynak, p.donem AS donem FROM progress p JOIN games g ON g.id=p.game_id"
    conds, params = [], []
    if game_id is not None:
        conds.append("p.game_id=?")
        params.append(int(game_id))
    if student_id is not None:
        conds.append("p.student_id=?")
        params.append(int(student_id))
    if conds:
        sql += " WHERE " + " AND ".join(conds)
    with db() as c:
        return pd.read_sql_query(sql, c, params=params)

def latest_progress(df):
    if df.empty:
        return df
    return df.sort_values(["tarih", "id"], ascending=[False, False]).drop_duplicates(["student_id", "game_id"], keep="first").reset_index(drop=True)

def import_progress(records, tarih, donem):
    require_admin()
    stamp = tarih.strftime("%Y-%m-%d")
    n = 0
    with db() as c:
        for r in records:
            c.execute("DELETE FROM progress WHERE student_id=? AND game_id=? AND kaynak='gecis' AND COALESCE(donem,'')=?", (r["student_id"], r["game_id"], donem or ""))
            c.execute("INSERT INTO progress (student_id, game_id, kart, seviye, tarih, kaynak, donem) VALUES (?,?,?,?,?, 'gecis', ?)", (r["student_id"], r["game_id"], r["kart"], r["seviye"], stamp, donem or ""))
            n += 1
    return n

def read_progress_file(uploaded):
    return _read_uploaded_table(uploaded), {}

def auto_map_columns(raw, cols, games):
    return {}

def build_progress_records(raw, cols, mapping):
    return [], {"matched": 0, "unmatched": [], "dup": 0, "blank": 0, "archived": 0, "info": {}, "ts_max": None}

def progress_template_xlsx():
    return None

def create_group(ad, game_id, turnuva, member_ids, tarih=None):
    require_admin()
    if tarih is None:
        tarih = date.today()
    with db() as c:
        cur = c.execute("INSERT INTO game_groups (ad, game_id, turnuva, tarih) VALUES (?,?,?,?)", (ad.strip(), int(game_id), turnuva.strip(), tarih.strftime("%Y-%m-%d")))
        c.executemany("INSERT OR IGNORE INTO group_members (group_id, student_id) VALUES (?,?)", [(cur.lastrowid, i) for i in member_ids])
    return True, "Grup oluşturuldu."

def update_group(gid, ad, turnuva, member_ids):
    require_admin()
    with db() as c:
        c.execute("UPDATE game_groups SET ad=?, turnuva=? WHERE id=?", (ad.strip(), turnuva.strip(), int(gid)))
        c.execute("DELETE FROM group_members WHERE group_id=?", (int(gid),))
        c.executemany("INSERT OR IGNORE INTO group_members (group_id, student_id) VALUES (?,?)", [(int(gid), i) for i in member_ids])
    return True, "Güncellendi."

def delete_group(gid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM game_groups WHERE id=?", (int(gid),))

def load_groups():
    with db() as c:
        return pd.read_sql_query("SELECT gg.id AS id, gg.ad AS ad, gg.turnuva AS turnuva, gg.tarih AS tarih, gg.game_id AS game_id, g.ad AS oyun, (SELECT COUNT(*) FROM group_members gm JOIN students s ON s.id=gm.student_id WHERE gm.group_id=gg.id AND s.arsiv=0) AS uye FROM game_groups gg JOIN games g ON g.id=gg.game_id ORDER BY gg.tarih DESC, gg.id DESC", c)

def group_label(r):
    return f"{r.oyun} · " + (f"{r.turnuva} · " if r.turnuva else "") + f"{r.ad}"

def group_members_df(gid):
    with db() as c:
        df = pd.read_sql_query("SELECT s.id AS id, s.ogrenci_no AS ogrenci_no, s.ad_soyad AS ad_soyad, s.sinif AS sinif, s.sube AS sube FROM group_members gm JOIN students s ON s.id=gm.student_id WHERE gm.group_id=? AND s.arsiv=0", c, params=[int(gid)])
    return _sort_students(df)

def save_group_scores(gid, rows, tarih=None, tur=""):
    require_admin()
    if tarih is None:
        tarih = date.today()
    with db() as c:
        c.executemany("INSERT INTO group_scores (group_id, student_id, tarih, tur, kart, puan) VALUES (?,?,?,?,?,?)", [(int(gid), s, tarih.strftime("%Y-%m-%d"), tur, _to_int(k) or 0, _to_int(p) or 0) for s, k, p in rows])
    return True, "Skorlar kaydedildi."

def delete_group_score(sid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM group_scores WHERE id=?", (int(sid),))

def group_scores_df():
    with db() as c:
        return pd.read_sql_query("SELECT sc.id AS id, sc.tarih AS tarih, sc.tur AS tur, sc.kart AS kart, sc.puan AS puan, sc.student_id AS student_id, sc.group_id AS group_id, gg.ad AS grup, gg.turnuva AS turnuva, gg.game_id AS game_id, g.ad AS oyun, s.ad_soyad AS ad_soyad, s.sinif AS sinif, s.sube AS sube, s.arsiv AS arsiv FROM group_scores sc JOIN game_groups gg ON gg.id=sc.group_id JOIN games g ON g.id=gg.game_id JOIN students s ON s.id=sc.student_id ORDER BY sc.tarih DESC, sc.id DESC", c)

def player_standings(sc, key="puan"):
    cols = ["Sıra", "Ad Soyad", "Sınıf", "Şube", "Kayıt", "Kart", "Puan"]
    if sc.empty:
        return pd.DataFrame(columns=cols)
    g = sc.groupby(["student_id", "ad_soyad", "sinif", "sube"], as_index=False).agg(kayit=("id", "count"), kart=("kart", "sum"), puan=("puan", "sum"))
    other = "kart" if key == "puan" else "puan"
    g = g.sort_values([key, other, "ad_soyad"], ascending=[False, False, True]).reset_index(drop=True)
    g.insert(0, "Sıra", range(1, len(g) + 1))
    return g.rename(columns={"ad_soyad": "Ad Soyad", "sinif": "Sınıf", "sube": "Şube", "kayit": "Kayıt", "kart": "Kart", "puan": "Puan"})[cols]

def group_totals(sc, key="puan"):
    cols = ["Sıra", "Oyun", "Turnuva", "Grup", "Oyuncu", "Kart", "Puan"]
    if sc.empty:
        return pd.DataFrame(columns=cols)
    g = sc.groupby(["group_id", "oyun", "turnuva", "grup"], as_index=False).agg(oyuncu=("student_id", "nunique"), kart=("kart", "sum"), puan=("puan", "sum"))
    other = "kart" if key == "puan" else "puan"
    g = g.sort_values([key, other], ascending=False).reset_index(drop=True)
    g.insert(0, "Sıra", range(1, len(g) + 1))
    g["turnuva"] = g.turnuva.replace("", "—")
    return g.rename(columns={"oyun": "Oyun", "turnuva": "Turnuva", "grup": "Grup", "oyuncu": "Oyuncu", "kart": "Kart", "puan": "Puan"})[cols]

def compute_stats(students, matches):
    df = students.copy()
    cols = ["galibiyet", "beraberlik", "maglubiyet"]
    if matches.empty:
        for c in cols:
            df[c] = 0
    else:
        a = pd.DataFrame({"id": matches.p1_id, "galibiyet": (matches.sonuc == "1").astype(int), "beraberlik": (matches.sonuc == "D").astype(int), "maglubiyet": (matches.sonuc == "2").astype(int)})
        b = pd.DataFrame({"id": matches.p2_id, "galibiyet": (matches.sonuc == "2").astype(int), "beraberlik": (matches.sonuc == "D").astype(int), "maglubiyet": (matches.sonuc == "1").astype(int)})
        agg = pd.concat([a, b]).groupby("id", as_index=False).sum()
        df = df.merge(agg, on="id", how="left")
        df[cols] = df[cols].fillna(0).astype(int)
    df["mac"] = df[cols].sum(axis=1)
    df["puan"] = df.galibiyet * 3 + df.beraberlik
    df["basari"] = ((df.galibiyet + 0.5 * df.beraberlik) / df.mac.where(df.mac > 0) * 100).fillna(0).round(1)
    df["seviye"] = ((df.galibiyet + 0.5 * df.beraberlik + 1) / (df.mac + 2)).round(3)
    return df

def branch_summary(stats):
    if stats.empty:
        return pd.DataFrame(columns=["sinif", "sube", "mevcut", "aktif", "galibiyet", "beraberlik", "mac", "basari"])
    g = stats.groupby(["sinif", "sube"]).agg(mevcut=("id", "count"), aktif=("mac", lambda s: int((s > 0).sum())), galibiyet=("galibiyet", "sum"), beraberlik=("beraberlik", "sum"), mac=("mac", "sum")).reset_index()
    g["basari"] = ((g.galibiyet + 0.5 * g.beraberlik) / g.mac.where(g.mac > 0) * 100).fillna(0).round(1)
    return g.sort_values(["sinif", "sube"]).reset_index(drop=True)

def auto_pair(ids, rating, already):
    ordered = sorted(ids, key=lambda i: (-rating[i], i))
    pairs, remaining = [], list(ordered)
    while len(remaining) >= 2:
        a = remaining.pop(0)
        best = min(range(len(remaining)), key=lambda k: (frozenset((a, remaining[k])) in already, abs(rating[a] - rating[remaining[k]]), k))
        b = remaining.pop(best)
        pairs.append((a, b))
    return pairs, remaining

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=Nunito:wght@400;600;700;800&display=swap');
:root{
  --bg:#F8F5F0; --card:#FFFDFA; --line:#E6DCD0; --milk:#F1EADF; --latte:#E2D3BF;
  --side:#45372F; --side-hi:#54443A;
  --taba:#A9794F; --taba-d:#8A5E3E; --mocha:#6F5443; --espresso:#3A2E27; --muted:#85766A;
  --shadow:0 4px 12px rgba(0,0,0,.03); --shadow-h:0 10px 24px rgba(69,55,47,.10);
}
html, body, .stApp{font-family:'Nunito',sans-serif;}
.stApp{background:var(--bg);color:var(--espresso);}
header[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:2rem;max-width:1200px;}
h1,h2,h3,h4{font-family:'DM Serif Display',Georgia,serif !important;color:var(--espresso) !important;}
section[data-testid="stSidebar"]{background:var(--side) !important;}
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] label{color:#EFE6DA;}
.brand{display:flex;align-items:center;gap:.8rem;padding:.2rem .3rem 1.2rem .3rem;margin-bottom:.9rem;border-bottom:1px solid rgba(255,255,255,.10);}
.brand-logo{width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,#B98B5F,#8A5E3E);display:flex;align-items:center;justify-content:center;font-size:1.45rem;}
.brand-t{font-family:'DM Serif Display',Georgia,serif;font-size:1.1rem;line-height:1.15;color:#FBF4EA;}
.brand-s{font-size:.7rem;color:#BBA896;letter-spacing:.1em;text-transform:uppercase;}
.stButton > button{border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--espresso);font-weight:700;}
.stButton > button[kind="primary"]{background:linear-gradient(135deg,#B5855A,#8F6342) !important;color:#FFF9F0 !important;border:none !important;}
</style>
"""

def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)

def hero(title, subtitle, icon_name=""):
    st.markdown(f'<div style="padding:1rem 0;"><h2>{esc(title)}</h2><p style="color:#85766A;">{esc(subtitle)}</p></div>', unsafe_allow_html=True)

def stat_cards(items):
    cols = st.columns(len(items))
    for col, item in zip(cols, items):
        col.markdown(f'<div style="background:#FFFDFA;border:1px solid #E6DCD0;border-radius:12px;padding:1rem;"><div style="font-size:0.75rem;color:#85766A;font-weight:700;text-transform:uppercase;">{esc(item[1])}</div><div style="font-size:2rem;font-weight:800;color:#3A2E27;margin-top:0.3rem;">{esc(str(item[2]))}</div></div>', unsafe_allow_html=True)

def flash(kind, msg):
    st.session_state["flash"] = (kind, msg)

def show_flash():
    f = st.session_state.pop("flash", None)
    if f:
        kind, msg = f
        {"success": st.success, "error": st.error, "warning": st.warning}.get(kind, st.info)(msg)

def student_labels(df):
    return {int(r.id): f"{r.ad_soyad}  ·  {r.sinif}-{r.sube}" for r in df.itertuples()}

def class_branch_filter(df, key):
    c1, c2 = st.columns(2)
    siniflar = ["Tümü"] + sorted(int(x) for x in df.sinif.unique())
    sel_s = c1.selectbox("Sınıf", siniflar, key=f"{key}_sinif", format_func=lambda x: x if x == "Tümü" else f"{x}. Sınıf")
    base = df if sel_s == "Tümü" else df[df.sinif == sel_s]
    subeler = ["Tümü"] + sorted(base.sube.unique().tolist(), key=tr_key)
    sel_b = c2.selectbox("Şube", subeler, key=f"{key}_sube", format_func=lambda x: x if x == "Tümü" else f"{x} Şubesi")
    return base if sel_b == "Tümü" else base[base.sube == sel_b]

def to_csv_bytes(df):
    return df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")

def to_xlsx_bytes(sheets):
    try:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            for name, df in sheets.items():
                df.to_excel(xw, sheet_name=name[:31], index=False)
        return buf.getvalue()
    except Exception:
        return None

def export_buttons(df, base_name, key, sheet="Veri"):
    c1, _ = st.columns([1, 4])
    xlsx = to_xlsx_bytes({sheet: df})
    if xlsx is not None:
        c1.download_button("Excel İndir (.xlsx)", xlsx, file_name=f"{base_name}.xlsx", mime=XLSX_MIME, key=f"{key}_xlsx", use_container_width=True)

def page_dashboard():
    students = load_students()
    matches = load_matches(active_only=True)
    if students.empty:
        st.info("Henüz öğrenci listesi yüklenmemiş. Yönetici Paneli'nden yükleme yapabilirsiniz.")
        return
    stats = compute_stats(students, matches)
    summary = branch_summary(stats)
    stat_cards([
        ("👥", "Toplam Öğrenci", len(students)),
        ("⚔️", "Toplam Maç", len(matches)),
        ("🏫", "Şube Sayısı", len(summary)),
    ])

def page_branch():
    hero("Şube Kontrol Paneli", "Sınıf ve şube bazlı başarı durumları")
    students = load_students()
    if students.empty:
        st.info("Öğrenci bulunamadı.")
        return
    stats = compute_stats(students, load_matches(active_only=True))
    f = class_branch_filter(stats, "br")
    show = f[["ogrenci_no", "ad_soyad", "sinif", "sube", "mac", "galibiyet", "beraberlik", "maglubiyet", "puan", "basari"]].rename(columns={
        "ogrenci_no": "No", "ad_soyad": "Ad Soyad", "sinif": "Sınıf", "sube": "Şube", "mac": "Maç", "galibiyet": "G", "beraberlik": "B", "maglubiyet": "M", "puan": "Puan", "basari": "Başarı %"
    })
    st.dataframe(show, hide_index=True, use_container_width=True)
    export_buttons(show, "ogrenci_listesi", "br")

def page_match():
    hero("Maç Girişi", "Yönetici şifresiyle maç kaydedin")
    students = load_students()
    if students.empty:
        st.info("Öğrenci yok.")
        return
    if not is_admin():
        st.warning("🔒 Salt Okunur Mod: Maç kaydetmek için Yönetici Paneli'nden giriş yapmalısınız.")
    games = load_games()
    pair_games = games[games.kategori == PAIR_CATEGORY]
    gname = dict(zip(pair_games.id, pair_games.ad))
    g_col, d_col = st.columns([3, 2])
    gid = g_col.selectbox("Oyun", list(gname.keys()), format_func=lambda i: gname[i])
    tarih = d_col.date_input("Maç Tarihi", value=date.today(), max_value=date.today(), format="DD.MM.YYYY")
    pool = class_branch_filter(students, "man")
    ids = [int(i) for i in pool.id]
    labels = student_labels(students)
    p1 = st.selectbox("1. Oyuncu", ids, index=None, key="man_p1", format_func=lambda i: labels[i])
    p2 = st.selectbox("2. Oyuncu", [i for i in ids if i != p1], index=None, key="man_p2", format_func=lambda i: labels[i])
    ready = p1 is not None and p2 is not None and is_admin()
    c1, c2, c3 = st.columns(3)
    if c1.button("1. Kazandı", disabled=not ready, use_container_width=True, type="primary" if ready else "secondary"):
        add_match(gid, p1, p2, "1", tarih)
        flash("success", "Maç kaydedildi.")
        st.rerun()
    if c2.button("Berabere", disabled=not ready, use_container_width=True):
        add_match(gid, p1, p2, "D", tarih)
        flash("success", "Maç kaydedildi.")
        st.rerun()
    if c3.button("2. Kazandı", disabled=not ready, use_container_width=True):
        add_match(gid, p1, p2, "2", tarih)
        flash("success", "Maç kaydedildi.")
        st.rerun()

def page_leaderboard():
    hero("Sıralama ve Geçmiş", "Genel liderlik tablosu")
    students = load_students()
    if students.empty:
        st.info("Öğrenci yok.")
        return
    stats = compute_stats(students, load_matches(active_only=True))
    f = class_branch_filter(stats, "lb")
    show = f.sort_values(["puan", "basari"], ascending=False)[["ad_soyad", "sinif", "sube", "mac", "galibiyet", "beraberlik", "maglubiyet", "puan", "basari"]].rename(columns={
        "ad_soyad": "Ad Soyad", "sinif": "Sınıf", "sube": "Şube", "mac": "Maç", "galibiyet": "G", "beraberlik": "B", "maglubiyet": "M", "puan": "Puan", "basari": "Başarı %"
    })
    st.dataframe(show, hide_index=True, use_container_width=True)
    export_buttons(show, "siralama", "lb")

def page_individual():
    hero("Bireysel Oyunlar", "Kart ve seviye takibi")
    students = load_students()
    games = load_games()
    if students.empty:
        st.info("Öğrenci yok.")
        return
    labels = student_labels(students)
    sid = st.selectbox("Öğrenci Seçin", list(labels.keys()), index=None, format_func=lambda i: labels[i])
    if sid:
        prog = load_progress(student_id=sid)
        cur = latest_progress(prog)
        if cur.empty:
            st.info("Bu öğrenciye ait kayıt yok.")
        else:
            st.dataframe(cur[["oyun", "kart", "seviye", "tarih"]], hide_index=True, use_container_width=True)

def page_groups():
    hero("Grup ve Turnuva", "Toplu oyun grupları")
    groups = load_groups()
    if groups.empty:
        st.info("Grup yok.")
    else:
        st.dataframe(groups[["oyun", "turnuva", "ad", "uye"]].rename(columns={"oyun": "Oyun", "turnuva": "Turnuva", "ad": "Grup", "uye": "Üye"}), hide_index=True, use_container_width=True)

def page_library():
    hero("Oyun Kütüphanesi", "Kategoriler")
    games = load_games()
    for cat in GAME_LIBRARY.keys():
        sub = games[games.kategori == cat]
        st.markdown(f"#### {cat}")
        st.write(", ".join(sub.ad.tolist()))

def render_import_ui(key_prefix):
    if not is_admin():
        st.warning("Yönetici girişi gerekli.")
        return
    st.markdown("#### Öğrenci Listesi Yükle (Excel/CSV)")
    up = st.file_uploader("Dosya Seç", type=["xlsx", "csv"], key=f"{key_prefix}_up")
    if up is not None:
        try:
            df, _ = parse_student_file(up)
            st.success(f"{len(df)} öğrenci okundu.")
            st.dataframe(df.head(5), hide_index=True)
            if st.button("Listeyi Kaydet", type="primary", key=f"{key_prefix}_btn"):
                import_students(df, replace=False)
                flash("success", "Öğrenciler başarıyla yüklendi.")
                st.rerun()
        except Exception as e:
            st.error(f"Hata: {e}")

def page_admin():
    hero("Yönetici Paneli", "Güvenlik ve ayarlar")
    if not is_admin():
        with st.form("login"):
            pin = st.text_input("PIN Şifre", type="password")
            if st.form_submit_button("Giriş Yap", type="primary"):
                if verify_pin(pin):
                    st.session_state["admin"] = True
                    st.rerun()
                else:
                    st.error("Hatalı PIN (Varsayılan: 6991)")
        return
    c1, c2 = st.columns([5, 1])
    c1.success("Yönetici oturumu açık.")
    c2.button("Çıkış", on_click=admin_logout, use_container_width=True)
    t1, t2 = st.tabs(["Veri Yükleme", "PIN Değiştir"])
    with t1:
        render_import_ui("adm")
    with t2:
        with st.form("pin"):
            new_p = st.text_input("Yeni PIN", type="password")
            if st.form_submit_button("Güncelle", type="primary"):
                if new_p.isdigit() and len(new_p) >= 4:
                    set_pin(new_p)
                    st.success("PIN güncellendi.")
                else:
                    st.error("En az 4 haneli rakam girin.")

PAGES = {
    PG_HOME: page_dashboard,
    PG_BRANCH: page_branch,
    PG_MATCH: page_match,
    PG_BOARD: page_leaderboard,
    PG_INDIV: page_individual,
    PG_GROUP: page_groups,
    PG_LIB: page_library,
    PG_ADMIN: page_admin,
}

def main():
    st.set_page_config(page_title="Akıl ve Zeka Oyunları", page_icon="🧠", layout="wide")
    init_db()
    inject_css()
    with st.sidebar:
        st.markdown(
            '<div class="brand"><div class="brand-logo">🧠</div><div><div class="brand-t">Akıl ve Zeka</div><div class="brand-s">Yönetim & Takip</div></div></div>',
            unsafe_allow_html=True,
        )
        page = st.radio("Menü", list(PAGES.keys()), label_visibility="collapsed", key="nav")
        if is_admin():
            st.markdown("<br><small style='color:#BBA896'>Yönetici Modu: AÇIK</small>", unsafe_allow_html=True)
            if st.button("Çıkış Yap", use_container_width=True):
                admin_logout()
                st.rerun()
        else:
            st.markdown("<br><small style='color:#BBA896'>Salt Okunur Mod</small>", unsafe_allow_html=True)
    show_flash()
    PAGES[page]()

if __name__ == "__main__":
    main()

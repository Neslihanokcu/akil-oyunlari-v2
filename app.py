# -*- coding: utf-8 -*-
"""
Akıl ve Zeka Oyunları Yönetim ve Takip Paneli
=============================================
Çalıştırma:
    pip install streamlit pandas openpyxl
    streamlit run app.py

- Veriler aynı klasördeki `akil_oyunlari.db` (SQLite) dosyasında saklanır.
- Varsayılan yönetici PIN'i: 6991.
- Tasarım: Modern Minimalist / Kurumsal SaaS Landing Page & Panel Entegrasyonu.
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

# ----------------------------------------------------------------------------
# SABİTLER
# ----------------------------------------------------------------------------
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

THEME_TOML = """[theme]
base = "light"
primaryColor = "#A9794F"
backgroundColor = "#F8F5F0"
secondaryBackgroundColor = "#F1EADF"
textColor = "#3A2E27"
font = "sans serif"
"""

# Menü sayfa anahtarları
PG_LANDING = "Karşılama Ekranı"
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
MAX_PIN_FAILS = 5      # üst üste hatalı PIN sınırı
LOCK_SECONDS = 300     # kilit süresi (sn)
PAGE_SIZE = 20         # sayfa başına kayıt


# ----------------------------------------------------------------------------
# YARDIMCI FONKSİYONLAR
# ----------------------------------------------------------------------------
def ensure_theme_config():
    try:
        cfg_dir = os.path.join(BASE_DIR, ".streamlit")
        cfg = os.path.join(cfg_dir, "config.toml")
        if not os.path.exists(cfg):
            os.makedirs(cfg_dir, exist_ok=True)
            with open(cfg, "w", encoding="utf-8") as f:
                f.write(THEME_TOML)
    except OSError:
        pass


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


# ----------------------------------------------------------------------------
# VERİTABANI
# ----------------------------------------------------------------------------
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


def _backup_before_migration():
    if not os.path.exists(DB_PATH):
        return
    try:
        conn = sqlite3.connect(DB_PATH)
        try:
            cols = [r[1] for r in conn.execute("PRAGMA table_info(students)")]
        finally:
            conn.close()
        if cols and "arsiv" not in cols:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(DB_PATH, os.path.join(BASE_DIR, f"akil_oyunlari_yedek_{stamp}.db"))
    except (OSError, sqlite3.Error):
        pass


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
    _backup_before_migration()
    with db() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY, value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ogrenci_no TEXT UNIQUE,
                ad_soyad TEXT NOT NULL,
                sinif INTEGER NOT NULL,
                sube TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS games (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ad TEXT NOT NULL UNIQUE,
                kategori TEXT NOT NULL,
                varsayilan INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS matches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tarih TEXT NOT NULL,
                game_id INTEGER NOT NULL REFERENCES games(id) ON DELETE CASCADE,
                p1_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                p2_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                sonuc TEXT NOT NULL CHECK (sonuc IN ('1','2','D')),
                CHECK (p1_id <> p2_id)
            );
            CREATE INDEX IF NOT EXISTS idx_matches_game ON matches(game_id);
            CREATE INDEX IF NOT EXISTS idx_students_sb ON students(sinif, sube);

            CREATE TABLE IF NOT EXISTS progress (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
                game_id INTEGER NOT NULL REFERENCES games(id) ON DELETE CASCADE,
                kart INTEGER,
                seviye TEXT,
                tarih TEXT NOT NULL,
                kaynak TEXT NOT NULL DEFAULT 'manuel',
                donem TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_progress_sg ON progress(student_id, game_id);

            CREATE TABLE IF NOT EXISTS game_groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ad TEXT NOT NULL,
                game_id INTEGER NOT NULL REFERENCES games(id) ON DELETE CASCADE,
                turnuva TEXT NOT NULL DEFAULT '',
                tarih TEXT NOT NULL,
                UNIQUE (game_id, turnuva, ad)
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
                tarih TEXT NOT NULL,
                tur TEXT NOT NULL DEFAULT '',
                kart INTEGER NOT NULL DEFAULT 0,
                puan INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS idx_gscores_group ON group_scores(group_id);
            CREATE INDEX IF NOT EXISTS idx_gscores_student ON group_scores(student_id);

            CREATE TABLE IF NOT EXISTS waitlist (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ad_soyad TEXT NOT NULL,
                eposta TEXT NOT NULL,
                kurum TEXT,
                tarih TEXT NOT NULL
            );
            """
        )
        _migrate(c)
        for kategori, oyunlar in GAME_LIBRARY.items():
            for ad in oyunlar:
                c.execute(
                    "INSERT OR IGNORE INTO games (ad, kategori, varsayilan) VALUES (?,?,1)",
                    (ad, kategori),
                )
    if _get_setting("admin_pin") is None:
        _store_pin(DEFAULT_PIN)


def _get_setting(key):
    with db() as c:
        row = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else None


def _set_setting(key, value):
    with db() as c:
        c.execute(
            "INSERT INTO settings (key, value) VALUES (?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )


def save_to_waitlist(ad_soyad, eposta, kurum):
    with db() as c:
        c.execute(
            "INSERT INTO waitlist (ad_soyad, eposta, kurum, tarih) VALUES (?,?,?,?)",
            (ad_soyad, eposta, kurum, date.today().strftime("%Y-%m-%d"))
        )


# ---- Yetkilendirme ---------------------------------------------------------
def is_admin():
    return bool(st.session_state.get("admin"))


def require_admin():
    if not is_admin():
        raise PermissionError(NO_AUTH_MSG)


def admin_logout():
    st.session_state["admin"] = False
    st.session_state.pop("del_pending", None)


# ---- PIN -------------------------------------------------------------------
def _hash_pin(pin, salt_hex):
    return hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), bytes.fromhex(salt_hex), 120_000).hex()


def _store_pin(pin):
    salt = secrets.token_hex(16)
    _set_setting("admin_pin", f"{salt}\${_hash_pin(pin, salt)}")


def set_pin(pin):
    require_admin()
    _store_pin(pin)


def verify_pin(pin):
    stored = _get_setting("admin_pin")
    if not stored or "\$" not in stored:
        return False
    salt, h = stored.split("\$", 1)
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


# ---- Öğrenciler ------------------------------------------------------------
def _sort_students(df):
    if df.empty:
        return df
    order = sorted(range(len(df)), key=lambda i: (int(df.sinif.iloc[i]), df.sube.iloc[i], tr_key(df.ad_soyad.iloc[i])))
    return df.iloc[order].reset_index(drop=True)


def load_students(durum="aktif"):
    where = {"aktif": "WHERE arsiv=0", "arsiv": "WHERE arsiv=1"}.get(durum, "")
    with db() as c:
        df = pd.read_sql_query(
            "SELECT id, ogrenci_no, ad_soyad, sinif, sube, arsiv, arsiv_tarih, arsiv_notu "
            f"FROM students {where}", c
        )
    return _sort_students(df)


def _dup_no_message(no):
    with db() as c:
        row = c.execute("SELECT arsiv FROM students WHERE ogrenci_no=?", (no,)).fetchone()
    if row and row["arsiv"]:
        return "Bu numara arşivdeki bir öğrenciye ait. Yönetici Paneli > Arşiv bölümünden geri alabilirsiniz."
    return None


def add_student(ogrenci_no, ad_soyad, sinif, sube):
    if not is_admin():
        return False, NO_AUTH_MSG
    ad_soyad = re.sub(r"\s+", " ", ad_soyad).strip()
    sube = sube.strip().upper()
    if not ad_soyad or not sube:
        return False, "Ad Soyad ve Şube boş bırakılamaz."
    try:
        with db() as c:
            cur = c.execute(
                "INSERT INTO students (ogrenci_no, ad_soyad, sinif, sube) VALUES (?,?,?,?)",
                (ogrenci_no.strip() or None, ad_soyad, int(sinif), sube),
            )
            if not ogrenci_no.strip():
                c.execute("UPDATE students SET ogrenci_no=? WHERE id=?", (f"OG{cur.lastrowid:04d}", cur.lastrowid))
        return True, f"{ad_soyad} eklendi."
    except sqlite3.IntegrityError:
        return False, _dup_no_message(ogrenci_no.strip()) or "Bu öğrenci numarası zaten kayıtlı."


def update_student(sid, ogrenci_no, ad_soyad, sinif, sube):
    if not is_admin():
        return False, NO_AUTH_MSG
    ad_soyad = re.sub(r"\s+", " ", ad_soyad).strip()
    sube = sube.strip().upper()
    if not ad_soyad or not sube or not ogrenci_no.strip():
        return False, "Tüm alanlar doldurulmalıdır."
    try:
        with db() as c:
            c.execute(
                "UPDATE students SET ogrenci_no=?, ad_soyad=?, sinif=?, sube=? WHERE id=?",
                (ogrenci_no.strip(), ad_soyad, int(sinif), sube, int(sid)),
            )
        return True, "Öğrenci bilgileri güncellendi."
    except sqlite3.IntegrityError:
        return False, _dup_no_message(ogrenci_no.strip()) or "Bu öğrenci numarası başka bir öğrenciye ait."


def delete_student(sid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM students WHERE id=?", (int(sid),))


# ---- Excel / CSV içe aktarma ------------------------------------------------
_ALIAS_ID = {"ogrenciid", "ogrencino", "ogrno", "okulno", "no", "id", "numara", "ogrencinumarasi"}
_ALIAS_NAME = {"adsoyad", "adisoyadi", "adsoyadi", "adisoyad", "isim", "ogrenci", "ogrenciadi", "adivesoyadi",
               "ogrenciadisoyadi", "ogrenciadsoyad", "ogrenciadsoyadi"}
_ALIAS_SINIF = {"sinif", "sinifi", "sinifduzeyi", "seviye"}
_ALIAS_SUBE = {"sube", "subesi"}


def _read_uploaded_table(uploaded):
    name = uploaded.name.lower()
    data = uploaded.getvalue()
    raw = None
    try:
        if name.endswith(".csv"):
            for enc in ("utf-8-sig", "cp1254", "latin-1"):
                try:
                    raw = pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding=enc, dtype=str)
                    break
                except (UnicodeDecodeError, pd.errors.ParserError):
                    continue
        else:
            raw = pd.read_excel(io.BytesIO(data), dtype=str)
    except Exception as e:
        raise ValueError(f"Dosya okunamadı: {e}")
    if raw is None or raw.empty:
        raise ValueError("Dosya boş veya okunamadı.")
    return raw


def parse_student_file(uploaded):
    raw = _read_uploaded_table(uploaded)
    cols = {}
    for col in raw.columns:
        n = _norm(col)
        for key, alias in (("id", _ALIAS_ID), ("name", _ALIAS_NAME), ("sinif", _ALIAS_SINIF), ("sube", _ALIAS_SUBE)):
            if n in alias and key not in cols:
                cols[key] = col
    missing = [lbl for k, lbl in (("name", "Ad_Soyad"), ("sinif", "Sınıf"), ("sube", "Şube")) if k not in cols]
    if missing:
        raise ValueError(
            "Gerekli sütun(lar) bulunamadı: " + ", ".join(missing)
            + f". Dosyadaki sütunlar: {', '.join(map(str, raw.columns))}"
        )

    out = pd.DataFrame()
    out["ad_soyad"] = raw[cols["name"]].fillna("").astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    out["sinif"] = pd.to_numeric(raw[cols["sinif"]].astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    out["sube"] = raw[cols["sube"]].fillna("").astype(str).str.strip().map(tr_upper)
    if "id" in cols:
        out["ogrenci_no"] = raw[cols["id"]].fillna("").astype(str).str.strip().replace("nan", "")
    else:
        out["ogrenci_no"] = ""

    notes = []
    total = len(out)
    valid = (out.ad_soyad != "") & out.sinif.notna() & (out.sube != "")
    if (~valid).sum():
        notes.append(f"{int((~valid).sum())} satır eksik bilgi nedeniyle atlandı.")
    out = out[valid].copy()
    out["sinif"] = out.sinif.astype(int)
    out = out[["ogrenci_no", "ad_soyad", "sinif", "sube"]]

    before = len(out)
    has_id = out.ogrenci_no != ""
    dup_id = out[has_id].duplicated("ogrenci_no", keep="first")
    out = out.drop(index=dup_id[dup_id].index)
    out = out.drop_duplicates(subset=["ad_soyad", "sinif", "sube"], keep="first")
    if before - len(out):
        notes.append(f"{before - len(out)} tekrar eden satır birleştirildi.")
    if "id" not in cols:
        notes.append("Dosyada öğrenci numarası sütunu yok; numaralar otomatik üretilecek (OG0001, OG0002...).")
    if total and out.empty:
        raise ValueError("Geçerli öğrenci satırı bulunamadı.")
    return out.reset_index(drop=True), notes


def import_students(df, replace=False):
    require_admin()
    added = updated = skipped = arch = 0
    with db() as c:
        if replace:
            c.execute("DELETE FROM students")
        for r in df.itertuples(index=False):
            no = (r.ogrenci_no or "").strip()
            if no:
                row = c.execute("SELECT id, arsiv FROM students WHERE ogrenci_no=?", (no,)).fetchone()
                if row:
                    arch += int(row["arsiv"] or 0)
                    c.execute(
                        "UPDATE students SET ad_soyad=?, sinif=?, sube=? WHERE id=?",
                        (r.ad_soyad, int(r.sinif), r.sube, row["id"]),
                    )
                    updated += 1
                    continue
                c.execute(
                    "INSERT INTO students (ogrenci_no, ad_soyad, sinif, sube) VALUES (?,?,?,?)",
                    (no, r.ad_soyad, int(r.sinif), r.sube),
                )
                added += 1
            else:
                row = c.execute(
                    "SELECT id, arsiv FROM students WHERE ad_soyad=? AND sinif=? AND sube=?",
                    (r.ad_soyad, int(r.sinif), r.sube),
                ).fetchone()
                if row:
                    arch += int(row["arsiv"] or 0)
                    skipped += 1
                    continue
                cur = c.execute(
                    "INSERT INTO students (ad_soyad, sinif, sube) VALUES (?,?,?)",
                    (r.ad_soyad, int(r.sinif), r.sube),
                )
                c.execute("UPDATE students SET ogrenci_no=? WHERE id=?", (f"OG{cur.lastrowid:04d}", cur.lastrowid))
                added += 1
    return added, updated, skipped, arch


# ---- Oyunlar ---------------------------------------------------------------
def load_games():
    with db() as c:
        return pd.read_sql_query("SELECT id, ad, kategori, varsayilan FROM games ORDER BY kategori, id", c)


def add_game(ad, kategori):
    if not is_admin():
        return False, NO_AUTH_MSG
    ad = ad.strip()
    if not ad:
        return False, "Oyun adı boş olamaz."
    try:
        with db() as c:
            c.execute("INSERT INTO games (ad, kategori, varsayilan) VALUES (?,?,0)", (ad, kategori))
        return True, f"'{ad}' kütüphaneye eklendi."
    except sqlite3.IntegrityError:
        return False, "Bu isimde bir oyun zaten var."


def delete_game(gid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM games WHERE id=? AND varsayilan=0", (int(gid),))


# ---- Maçlar ----------------------------------------------------------------
def add_match(game_id, p1, p2, sonuc, tarih=None):
    if game_id is None:
        return False, "Lütfen bir oyun seçin."
    if p1 is None or p2 is None:
        return False, "Lütfen iki öğrenciyi de seçin."
    if int(p1) == int(p2):
        return False, "Bir öğrenci kendisiyle eşleştirilemez."
    if sonuc not in ("1", "2", "D"):
        return False, "Geçersiz sonuç."
    if tarih is None:
        tarih = date.today()
    if isinstance(tarih, datetime):
        tarih = tarih.date()
    if not isinstance(tarih, date):
        return False, "Lütfen geçerli bir maç tarihi seçin."
    if tarih > date.today():
        return False, "Maç tarihi bugünden sonra olamaz."
    with db() as c:
        c.execute(
            "INSERT INTO matches (tarih, game_id, p1_id, p2_id, sonuc) VALUES (?,?,?,?,?)",
            (tarih.strftime("%Y-%m-%d"), int(game_id), int(p1), int(p2), sonuc),
        )
    return True, "Maç kaydedildi."


def fmt_tarih(v):
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").strftime("%d.%m.%Y")
    except (ValueError, TypeError):
        return str(v) if v else "—"


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
        return pd.read_sql_query(
            sql + f" ORDER BY substr(m.tarih, 1, 10) {direction}, m.id {direction}", c, params=params
        )


def played_pairs(game_id):
    m = load_matches(game_id)
    return {frozenset((int(a), int(b))) for a, b in zip(m.p1_id, m.p2_id)}


# ---- Arşiv -----------------------------------------------------------------
def archive_students(ids, note=""):
    require_admin()
    ids = sorted({int(i) for i in ids})
    if not ids:
        return 0
    stamp = date.today().strftime("%Y-%m-%d")
    note = (note or "").strip() or None
    with db() as c:
        cur = c.executemany(
            "UPDATE students SET arsiv=1, arsiv_tarih=?, arsiv_notu=? WHERE id=? AND arsiv=0",
            [(stamp, note, i) for i in ids],
        )
        return cur.rowcount


def restore_students(ids):
    require_admin()
    ids = sorted({int(i) for i in ids})
    if not ids:
        return 0
    with db() as c:
        cur = c.executemany(
            "UPDATE students SET arsiv=0, arsiv_tarih=NULL, arsiv_notu=NULL WHERE id=? AND arsiv=1",
            [(i,) for i in ids],
        )
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


# ---- Bireysel oyun ilerlemesi -----------------------------------
def save_progress(student_id, game_id, kart, seviye, tarih=None, kaynak="manuel", donem=None):
    kart, seviye = _to_int(kart), _nz(seviye)
    if kart is not None and kart < 0:
        return False, "Kart sayısı negatif olamaz."
    if tarih is None:
        tarih = date.today()
    if isinstance(tarih, datetime):
        tarih = tarih.date()
    if not isinstance(tarih, date):
        return False, "Geçerli bir tarih seçin."
    if tarih > date.today():
        return False, "Tarih bugünden sonra olamaz."
    with db() as c:
        c.execute(
            "INSERT INTO progress (student_id, game_id, kart, seviye, tarih, kaynak, donem) VALUES (?,?,?,?,?,?,?)",
            (int(student_id), int(game_id), kart, str(seviye) if seviye is not None else None,
             tarih.strftime("%Y-%m-%d"), kaynak, donem),
        )
    return True, "Kaydedildi."


def load_progress(game_id=None, student_id=None):
    sql = (
        "SELECT p.id AS id, p.student_id AS student_id, p.game_id AS game_id, g.ad AS oyun, g.kategori AS kategori, "
        "p.kart AS kart, p.seviye AS seviye, p.tarih AS tarih, p.kaynak AS kaynak, p.donem AS donem "
        "FROM progress p JOIN games g ON g.id=p.game_id"
    )
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
    d = df.sort_values(["tarih", "id"], ascending=[False, False])
    return d.drop_duplicates(["student_id", "game_id"], keep="first").reset_index(drop=True)


def import_progress(records, tarih, donem):
    require_admin()
    if isinstance(tarih, datetime):
        tarih = tarih.date()
    stamp = tarih.strftime("%Y-%m-%d")
    n = 0
    with db() as c:
        for r in records:
            c.execute(
                "DELETE FROM progress WHERE student_id=? AND game_id=? AND kaynak='gecis' AND COALESCE(donem,'')=?",
                (r["student_id"], r["game_id"], donem or ""),
            )
            c.execute(
                "INSERT INTO progress (student_id, game_id, kart, seviye, tarih, kaynak, donem) "
                "VALUES (?,?,?,?,?, 'gecis', ?)",
                (r["student_id"], r["game_id"], r["kart"], r["seviye"], stamp, donem or ""),
            )
            n += 1
    return n


_ALIAS_TIME = {"zamandamgasi", "tarih", "zaman", "timestamp"}


def read_progress_file(uploaded):
    raw = _read_uploaded_table(uploaded)
    cols = {}
    for col in raw.columns:
        n = _norm(col)
        for key, alias in (("id", _ALIAS_ID), ("name", _ALIAS_NAME), ("time", _ALIAS_TIME)):
            if n in alias and key not in cols:
                cols[key] = col
    if "id" not in cols and "name" not in cols:
        raise ValueError(
            "Öğrenciyi belirleyecek sütun bulunamadı: dosyada Ogrenci_ID (veya Ad_Soyad) sütunu olmalıdır."
        )
    return raw, cols


def auto_map_columns(raw, cols, games):
    keyed = set(cols.values())
    gkeys = [(int(r.id), _norm(r.ad)) for r in games.itertuples()]
    out = {}
    for col in raw.columns:
        if col in keyed:
            continue
        n = _norm(col)
        out[col] = [gid for gid, gk in gkeys if gk and gk in n]
    return out


def _norm_no(v):
    s = "" if v is None else str(v).strip()
    if s.lower() in ("nan", "none"):
        return ""
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return s.upper()


def _norm_name(s):
    return re.sub(r"\s+", " ", tr_upper(str(s))).strip()


def build_progress_records(raw, cols, mapping):
    all_st = load_students("tumu")
    no_map, name_map, info, arch = {}, {}, {}, set()
    for r in all_st.itertuples():
        sid = int(r.id)
        if r.ogrenci_no:
            no_map[_norm_no(r.ogrenci_no)] = sid
        name_map.setdefault(_norm_name(r.ad_soyad), []).append(sid)
        info[sid] = (r.ad_soyad, f"{r.sinif}-{r.sube}" + (" (arşiv)" if int(r.arsiv) else ""))
        if int(r.arsiv):
            arch.add(sid)
    id_col, name_col, time_col = cols.get("id"), cols.get("name"), cols.get("time")
    ts = pd.to_datetime(raw[time_col], errors="coerce") if time_col else None

    row_sid, unmatched, blank = {}, [], 0
    for idx in raw.index:
        raw_id = _norm_no(raw.at[idx, id_col]) if id_col else ""
        raw_name = _nz(raw.at[idx, name_col]) if name_col else None
        if not raw_id and not raw_name:
            blank += 1
            continue
        sid = no_map.get(raw_id) if raw_id else None
        if sid is None and raw_name:
            cand = name_map.get(_norm_name(raw_name), [])
            if len(cand) == 1:
                sid = cand[0]
        if sid is None:
            unmatched.append(f"{raw_id or '—'} · {raw_name or '—'}")
        else:
            row_sid[idx] = sid

    order = list(row_sid.keys())
    if ts is not None:
        order.sort(key=lambda i: ts.at[i] if pd.notna(ts.at[i]) else pd.Timestamp.min)
    final, dup = {}, 0
    for i in order:
        if row_sid[i] in final:
            dup += 1
        final[row_sid[i]] = i

    recs = {}
    for sid, i in final.items():
        for col, gids in mapping.items():
            v = _nz(raw.at[i, col])
            if v is None or str(v).strip().lower() in ("nan", "none", "-"):
                continue
            s = str(v).strip()
            kart = seviye = None
            try:
                f = float(s.replace(",", "."))
                if f >= 0 and f == int(f):
                    kart = int(f)
                else:
                    seviye = s
            except (ValueError, OverflowError):
                seviye = s
            for gid in gids:
                d = recs.setdefault((sid, int(gid)), {"kart": None, "seviye": None})
                if kart is not None and d["kart"] is None:
                    d["kart"] = kart
                if seviye is not None and d["seviye"] is None:
                    d["seviye"] = seviye
    records = [{"student_id": s, "game_id": g, **d} for (s, g), d in recs.items()]
    return records, {
        "matched": len(final), "unmatched": unmatched, "dup": dup, "blank": blank,
        "archived": len([s for s in final if s in arch]), "info": info,
        "ts_max": (ts.max().date() if ts is not None and ts.notna().any() else None),
    }


def progress_template_xlsx():
    games = load_games()
    indiv = games[games.kategori == INDIV_CATEGORY].ad.tolist()
    example = pd.DataFrame(
        [["OG0001", "Örnek Öğrenci 1"] + [20 + 5 * i for i in range(len(indiv))],
         ["OG0002", "Örnek Öğrenci 2"] + [10 + 3 * i for i in range(len(indiv))]],
        columns=["Ogrenci_ID", "Ad_Soyad"] + indiv,
    )
    help_df = pd.DataFrame({"Açıklama": [
        "Ogrenci_ID sütunu zorunludur. Ad_Soyad isteğe bağlıdır.",
        "Her oyun ayrı bir sütundur.",
    ]})
    return to_xlsx_bytes({"Geçiş Verisi": example, "Açıklama": help_df})


# ---- Toplu oyunlar: grup / turnuva / skor ----------------------------------------
def _clean(s):
    return re.sub(r"\s+", " ", s or "").strip()


def _as_date(tarih):
    if tarih is None:
        return date.today()
    if isinstance(tarih, datetime):
        return tarih.date()
    return tarih if isinstance(tarih, date) else None


def create_group(ad, game_id, turnuva, member_ids, tarih=None):
    ad, turnuva = _clean(ad), _clean(turnuva)
    if game_id is None:
        return False, "Lütfen bir toplu oyun seçin."
    if not ad:
        return False, "Grup adı boş bırakılamaz."
    member_ids = sorted({int(i) for i in (member_ids or [])})
    if not member_ids:
        return False, "Gruba en az bir oyuncu ekleyin."
    tarih = _as_date(tarih)
    if tarih is None or tarih > date.today():
        return False, "Geçerli bir tarih seçin."
    try:
        with db() as c:
            kat = c.execute("SELECT kategori FROM games WHERE id=?", (int(game_id),)).fetchone()
            if not kat or kat["kategori"] != GROUP_CATEGORY:
                return False, "Gruplar yalnızca toplu oyunlar için oluşturulabilir."
            cur = c.execute(
                "INSERT INTO game_groups (ad, game_id, turnuva, tarih) VALUES (?,?,?,?)",
                (ad, int(game_id), turnuva, tarih.strftime("%Y-%m-%d")),
            )
            c.executemany(
                "INSERT OR IGNORE INTO group_members (group_id, student_id) VALUES (?,?)",
                [(cur.lastrowid, i) for i in member_ids],
            )
        return True, f"'{ad}' grubu {len(member_ids)} oyuncu ile oluşturuldu."
    except sqlite3.IntegrityError:
        return False, "Bu oyunda, aynı turnuva altında bu isimde bir grup zaten var."


def update_group(gid, ad, turnuva, member_ids):
    ad, turnuva = _clean(ad), _clean(turnuva)
    if not ad:
        return False, "Grup adı boş bırakılamaz."
    want = {int(i) for i in (member_ids or [])}
    try:
        with db() as c:
            c.execute("UPDATE game_groups SET ad=?, turnuva=? WHERE id=?", (ad, turnuva, int(gid)))
            current = {r[0] for r in c.execute("SELECT student_id FROM group_members WHERE group_id=?", (int(gid),))}
            active = {r[0] for r in c.execute(
                "SELECT s.id FROM group_members gm JOIN students s ON s.id=gm.student_id "
                "WHERE gm.group_id=? AND s.arsiv=0", (int(gid),))}
            c.executemany("DELETE FROM group_members WHERE group_id=? AND student_id=?",
                          [(int(gid), i) for i in active - want])
            c.executemany("INSERT OR IGNORE INTO group_members (group_id, student_id) VALUES (?,?)",
                          [(int(gid), i) for i in want - current])
        return True, "Grup güncellendi."
    except sqlite3.IntegrityError:
        return False, "Bu oyunda, aynı turnuva altında bu isimde başka bir grup zaten var."


def delete_group(gid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM game_groups WHERE id=?", (int(gid),))


def load_groups():
    with db() as c:
        return pd.read_sql_query(
            "SELECT gg.id AS id, gg.ad AS ad, gg.turnuva AS turnuva, gg.tarih AS tarih, gg.game_id AS game_id, "
            "g.ad AS oyun, "
            "(SELECT COUNT(*) FROM group_members gm JOIN students s ON s.id=gm.student_id "
            " WHERE gm.group_id=gg.id AND s.arsiv=0) AS uye "
            "FROM game_groups gg JOIN games g ON g.id=gg.game_id ORDER BY gg.tarih DESC, gg.id DESC", c)


def group_label(r):
    return f"{r.oyun} · " + (f"{r.turnuva} · " if r.turnuva else "") + f"{r.ad}"


def group_members_df(gid):
    with db() as c:
        df = pd.read_sql_query(
            "SELECT s.id AS id, s.ogrenci_no AS ogrenci_no, s.ad_soyad AS ad_soyad, s.sinif AS sinif, s.sube AS sube "
            "FROM group_members gm JOIN students s ON s.id=gm.student_id "
            "WHERE gm.group_id=? AND s.arsiv=0", c, params=[int(gid)])
    return _sort_students(df)


def save_group_scores(gid, rows, tarih=None, tur=""):
    tarih = _as_date(tarih)
    if tarih is None or tarih > date.today():
        return False, "Geçerli bir tarih seçin."
    clean = []
    for sid, kart, puan in rows:
        k, p = _to_int(kart), _to_int(puan)
        if k is None and p is None:
            continue
        if (k or 0) < 0 or (p or 0) < 0:
            return False, "Kart ve puan negatif olamaz."
        clean.append((int(sid), k or 0, p or 0))
    if not clean:
        return False, "Kaydedilecek skor girilmedi."
    with db() as c:
        members = {r[0] for r in c.execute("SELECT student_id FROM group_members WHERE group_id=?", (int(gid),))}
        if any(s not in members for s, _, _ in clean):
            return False, "Skor girilen oyunculardan biri bu grubun üyesi değil."
        c.executemany(
            "INSERT INTO group_scores (group_id, student_id, tarih, tur, kart, puan) VALUES (?,?,?,?,?,?)",
            [(int(gid), s, tarih.strftime("%Y-%m-%d"), _clean(tur), k, p) for s, k, p in clean],
        )
    return True, f"{len(clean)} oyuncunun skoru kaydedildi."


def delete_group_score(sid):
    require_admin()
    with db() as c:
        c.execute("DELETE FROM group_scores WHERE id=?", (int(sid),))


def group_scores_df():
    with db() as c:
        return pd.read_sql_query(
            "SELECT sc.id AS id, sc.tarih AS tarih, sc.tur AS tur, sc.kart AS kart, sc.puan AS puan, "
            "sc.student_id AS student_id, sc.group_id AS group_id, gg.ad AS grup, gg.turnuva AS turnuva, "
            "gg.game_id AS game_id, g.ad AS oyun, s.ad_soyad AS ad_soyad, s.sinif AS sinif, s.sube AS sube, "
            "s.arsiv AS arsiv "
            "FROM group_scores sc JOIN game_groups gg ON gg.id=sc.group_id JOIN games g ON g.id=gg.game_id "
            "JOIN students s ON s.id=sc.student_id ORDER BY sc.tarih DESC, sc.id DESC", c)


def player_standings(sc, key="puan"):
    cols = ["Sıra", "Ad Soyad", "Sınıf", "Şube", "Kayıt", "Kart", "Puan"]
    if sc.empty:
        return pd.DataFrame(columns=cols)
    g = sc.groupby(["student_id", "ad_soyad", "sinif", "sube"], as_index=False).agg(
        kayit=("id", "count"), kart=("kart", "sum"), puan=("puan", "sum"))
    other = "kart" if key == "puan" else "puan"
    g = g.sort_values([key, other, "ad_soyad"], ascending=[False, False, True]).reset_index(drop=True)
    g.insert(0, "Sıra", range(1, len(g) + 1))
    return g.rename(columns={"ad_soyad": "Ad Soyad", "sinif": "Sınıf", "sube": "Şube", "kayit": "Kayıt",
                             "kart": "Kart", "puan": "Puan"})[cols]


def group_totals(sc, key="puan"):
    cols = ["Sıra", "Oyun", "Turnuva", "Grup", "Oyuncu", "Kart", "Puan"]
    if sc.empty:
        return pd.DataFrame(columns=cols)
    g = sc.groupby(["group_id", "oyun", "turnuva", "grup"], as_index=False).agg(
        oyuncu=("student_id", "nunique"), kart=("kart", "sum"), puan=("puan", "sum"))
    other = "kart" if key == "puan" else "puan"
    g = g.sort_values([key, other], ascending=False).reset_index(drop=True)
    g.insert(0, "Sıra", range(1, len(g) + 1))
    g["turnuva"] = g.turnuva.replace("", "—")
    return g.rename(columns={"oyun": "Oyun", "turnuva": "Turnuva", "grup": "Grup", "oyuncu": "Oyuncu",
                             "kart": "Kart", "puan": "Puan"})[cols]


# ---- İstatistikler ---------------------------------------------------------
def compute_stats(students, matches):
    df = students.copy()
    cols = ["galibiyet", "beraberlik", "maglubiyet"]
    if matches.empty:
        for c in cols:
            df[c] = 0
    else:
        a = pd.DataFrame({
            "id": matches.p1_id,
            "galibiyet": (matches.sonuc == "1").astype(int),
            "beraberlik": (matches.sonuc == "D").astype(int),
            "maglubiyet": (matches.sonuc == "2").astype(int),
        })
        b = pd.DataFrame({
            "id": matches.p2_id,
            "galibiyet": (matches.sonuc == "2").astype(int),
            "beraberlik": (matches.sonuc == "D").astype(int),
            "maglubiyet": (matches.sonuc == "1").astype(int),
        })
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
    g = (
        stats.groupby(["sinif", "sube"])
        .agg(
            mevcut=("id", "count"),
            aktif=("mac", lambda s: int((s > 0).sum())),
            galibiyet=("galibiyet", "sum"),
            beraberlik=("beraberlik", "sum"),
            mac=("mac", "sum"),
        )
        .reset_index()
    )
    g["basari"] = ((g.galibiyet + 0.5 * g.beraberlik) / g.mac.where(g.mac > 0) * 100).fillna(0).round(1)
    return g.sort_values(["sinif", "sube"]).reset_index(drop=True)


def auto_pair(ids, rating, already):
    ordered = sorted(ids, key=lambda i: (-rating[i], i))
    pairs, remaining = [], list(ordered)
    while len(remaining) >= 2:
        a = remaining.pop(0)
        best = min(
            range(len(remaining)),
            key=lambda k: (frozenset((a, remaining[k])) in already, abs(rating[a] - rating[remaining[k]]), k),
        )
        b = remaining.pop(best)
        pairs.append((a, b))
    return pairs, remaining


# ----------------------------------------------------------------------------
# İKONLAR
# ----------------------------------------------------------------------------
_ICON_PATHS = {
    "home": '<path d="M3 11l9-8 9 8"/><path d="M5 10v10h14V10"/>',
    "school": '<path d="M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 2 9 2 12 0v-5"/>',
    "swords": '<path d="M14 4l6-1-1 6-9 9-4-4z"/><path d="M5 15l-2 2 4 4 2-2"/>',
    "trophy": '<path d="M8 21h8M12 17v4M7 4h10v5a5 5 0 0 1-10 0z"/><path d="M7 6H4v2a3 3 0 0 0 3 3M17 6h3v2a3 3 0 0 1-3 3"/>',
    "dice": '<rect x="4" y="4" width="16" height="16" rx="3"/><circle cx="9" cy="9" r="1"/><circle cx="15" cy="15" r="1"/><circle cx="15" cy="9" r="1"/><circle cx="9" cy="15" r="1"/>',
    "lock": '<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>',
    "users": '<circle cx="9" cy="8" r="3"/><path d="M3 20c0-3 3-5 6-5s6 2 6 5"/><path d="M16 5a3 3 0 0 1 0 6M18 15c2 .5 3 2.2 3 4"/>',
    "chart": '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    "scale": '<path d="M12 3v18M5 21h14M5 7h14"/><path d="M5 7l-3 7a3 3 0 0 0 6 0zM19 7l-3 7a3 3 0 0 0 6 0z"/>',
    "target": '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="4"/>',
    "puzzle": '<path d="M10 4a2 2 0 1 1 4 0v2h4v4h-2a2 2 0 1 0 0 4h2v6H4v-6h2a2 2 0 1 0 0-4H4V6h6z"/>',
    "brain": '<path d="M9 4a3 3 0 0 0-3 3 3 3 0 0 0-2 5 3 3 0 0 0 2 5 3 3 0 0 0 6 1V4a3 3 0 0 0-3 0zM15 4a3 3 0 0 1 3 3 3 3 0 0 1 2 5 3 3 0 0 1-2 5 3 3 0 0 1-6 1"/>',
    "archive": '<rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11h14V8M10 12h4"/>',
    "calendar": '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
}

_EMOJI_TO_ICON = {
    "⚔️": "swords", "⚖️": "scale", "👩‍🎓": "users", "📈": "chart", "🏫": "school", "🏆": "trophy",
    "🎲": "dice", "🔐": "lock", "👥": "users", "🎯": "target", "📊": "chart", "🧠": "brain", "🧩": "puzzle",
}


def icon(name, size=22):
    return (
        f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
        f'stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
        f'{_ICON_PATHS[name]}</svg>'
    )


def _ico(x, size=22):
    name = _EMOJI_TO_ICON.get(x, x)
    return icon(name, size) if name in _ICON_PATHS else x


# ----------------------------------------------------------------------------
# TASARIM (Clean Light & Modern Dark Hero Mix)
# ----------------------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=Nunito:wght@400;600;700;800&display=swap');
:root{
  --bg:#F9F9F8; --card:#FFFFFF; --line:#EAEAEA; --milk:#F1F3F2; --latte:#E2E5E4;
  --side:#1E232A; --side-hi:#282E37;
  --taba:#10B981; --taba-d:#059669; --mocha:#374151; --espresso:#111827; --muted:#6B7280;
  --shadow:0 2px 8px rgba(0,0,0,.04); --shadow-h:0 8px 20px rgba(0,0,0,.08);
}
html, body, .stApp{font-family:'Nunito',system-ui,-apple-system,'Segoe UI',sans-serif;}
.stApp{background:var(--bg);color:var(--espresso);}
.stApp p, .stApp label, .stApp li, .stApp span, .stApp div[data-testid="stMarkdownContainer"]{color:var(--espresso);}
header[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:1.5rem;max-width:1200px;}
h1,h2,h3,h4{font-family:'DM Serif Display',Georgia,serif !important;color:var(--espresso) !important;letter-spacing:.2px;}
hr{border-color:var(--line);}

/* ============ LANDING PAGE HERO (Koyu / Modern SaaS Stili) ============ */
.hero-landing {
  background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
  border-radius: 20px;
  padding: 3.5rem 2.5rem;
  color: #FFFFFF;
  text-align: center;
  position: relative;
  overflow: hidden;
  box-shadow: 0 20px 40px rgba(0,0,0,0.15);
  margin-bottom: 2rem;
}
.hero-landing h1 {
  color: #FFFFFF !important;
  font-size: 3.2rem !important;
  line-height: 1.15;
  margin-bottom: 1rem;
}
.hero-landing p {
  color: #94A3B8 !important;
  font-size: 1.25rem;
  max-width: 700px;
  margin: 0 auto 2rem auto;
}
.badge-saas {
  display: inline-block;
  background: rgba(16, 185, 129, 0.15);
  color: #10B981;
  padding: 0.4rem 1rem;
  border-radius: 99px;
  font-size: 0.85rem;
  font-weight: 700;
  letter-spacing: 0.05em;
  text-transform: uppercase;
  margin-bottom: 1.2rem;
  border: 1px solid rgba(16, 185, 129, 0.3);
}

/* ============ SIDEBAR ============ */
section[data-testid="stSidebar"], section[data-testid="stSidebar"] > div{background:var(--side) !important;}
section[data-testid="stSidebar"]{border-right:none;}
section[data-testid="stSidebar"] .block-container{padding-top:1.5rem;}
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] div{color:#E5E7EB;}
[data-testid="stSidebarCollapseButton"] button, [data-testid="stSidebarCollapseButton"] span,
[data-testid="stSidebarCollapseButton"] svg{color:#E5E7EB !important;fill:#E5E7EB !important;}
.brand{display:flex;align-items:center;gap:.8rem;padding:.2rem .3rem 1.2rem .3rem;margin-bottom:.9rem;border-bottom:1px solid rgba(255,255,255,.10);}
.brand-logo{width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,#10B981,#059669);display:flex;align-items:center;justify-content:center;font-size:1.45rem;box-shadow:0 6px 14px rgba(0,0,0,.28);}
section[data-testid="stSidebar"] .brand-t{font-family:'DM Serif Display',Georgia,serif;font-size:1.1rem;line-height:1.15;color:#F9FAFB;}
section[data-testid="stSidebar"] .brand-s{font-size:.7rem;color:#9CA3AF;letter-spacing:.1em;text-transform:uppercase;margin-top:3px;}
section[data-testid="stSidebar"] .side-label{font-size:.68rem;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#9CA3AF;margin:0 0 .5rem .4rem;}
section[data-testid="stSidebar"] div[role="radiogroup"]{display:flex;flex-direction:column;gap:.3rem;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label{background:transparent;border:1px solid transparent;border-radius:10px;padding:.65rem .9rem;width:100%;cursor:pointer;transition:background .2s ease, transform .2s ease;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover{background:rgba(255,255,255,.07);transform:translateX(3px);}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked){background:linear-gradient(135deg,#10B981 0%,#059669 100%);box-shadow:0 6px 14px rgba(0,0,0,.25);}
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child{display:none;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label p{font-weight:700;font-size:.96rem;margin:0;color:#E5E7EB;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p,
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) div{color:#FFFFFF !important;}
section[data-testid="stSidebar"] .side-note{margin-top:1.4rem;padding:.75rem .95rem;border-radius:10px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.10);font-size:.8rem;color:#D1D5DB;}

/* ============ BANNER & METRİKLER ============ */
.welcome{position:relative;overflow:hidden;display:flex;align-items:center;justify-content:space-between;gap:1.2rem;padding:2rem 2.2rem;margin-bottom:1.3rem;border-radius:16px;background:#FFFFFF;border:1px solid var(--line);box-shadow:var(--shadow);}
.welcome-t{font-family:'DM Serif Display',Georgia,serif;font-size:2.15rem;line-height:1.15;color:var(--espresso);}
.welcome-s{

# -*- coding: utf-8 -*-
"""
Akıl ve Zeka Oyunları Yönetim ve Takip Paneli (Sadeleştirilmiş Kurumsal Sürüm)
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

# ----------------------------------------------------------------------------
# YARDIMCI FONKSİYONLAR
# ----------------------------------------------------------------------------
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

def init_db():
    with db() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT, ogrenci_no TEXT UNIQUE,
                ad_soyad TEXT NOT NULL, sinif INTEGER NOT NULL, sube TEXT NOT NULL,
                arsiv INTEGER NOT NULL DEFAULT 0, arsiv_tarih TEXT, arsiv_notu TEXT
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
                sonuc TEXT NOT NULL CHECK (sonuc IN ('1','2','D'))
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

def load_games():
    with db() as c:
        return pd.read_sql_query("SELECT id, ad, kategori, varsayilan FROM games ORDER BY kategori, id", c)

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

def save_progress(student_id, game_id, kart, seviye, tarih=None, kaynak="manuel", donem=None):
    require_admin()
    kart, seviye = _to_int(kart), _nz(seviye)
    if tarih is None:
        tarih = date.today()
    with db() as c:
        c.execute(
            "INSERT INTO progress (student_id, game_id, kart, seviye, tarih, kaynak, donem) VALUES (?,?,?,?,?,?,?)",
            (int(student_id), int(game_id), kart, str(seviye) if seviye is not None else None, tarih.strftime("%Y-%m-%d"), kaynak, donem),
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

def load_groups():
    with db() as c:
        return pd.read_sql_query(
            "SELECT gg.id AS id, gg.ad AS ad, gg.turnuva AS turnuva, gg.tarih AS tarih, gg.game_id AS game_id, "
            "g.ad AS oyun, (SELECT COUNT(*) FROM group_members gm JOIN students s ON s.id=gm.student_id WHERE gm.group_id=gg.id AND s.arsiv=0) AS uye "
            "FROM game_groups gg JOIN games g ON g.id=gg.game_id ORDER BY gg.tarih DESC, gg.id DESC", c)

def group_members_df(gid):
    with db() as c:
        df = pd.read_sql_query(
            "SELECT s.id AS id, s.ogrenci_no AS ogrenci_no, s.ad_soyad AS ad_soyad, s.sinif AS sinif, s.sube AS sube "
            "FROM group_members gm JOIN students s ON s.id=gm.student_id WHERE gm.group_id=? AND s.arsiv=0", c, params=[int(gid)])
    return _sort_students(df)

def save_group_scores(gid, rows, tarih=None, tur=""):
    require_admin()
    if tarih is None:
        tarih = date.today()
    clean = []
    for sid, kart, puan in rows:
        k, p = _to_int(kart), _to_int(puan)
        if k is None and p is None:
            continue
        clean.append((int(sid), k or 0, p or 0))
    if not clean:
        return False, "Skor girilmedi."
    with db() as c:
        c.executemany(
            "INSERT INTO group_scores (group_id, student_id, tarih, tur, kart, puan) VALUES (?,?,?,?,?,?)",
            [(int(gid), s, tarih.strftime("%Y-%m-%d"), _clean(tur), k, p) for s, k, p in clean],
        )
    return True, "Skorlar kaydedildi."

def group_scores_df():
    with db() as c:
        return pd.read_sql_query(
            "SELECT sc.id AS id, sc.tarih AS tarih, sc.tur AS tur, sc.kart AS kart, sc.puan AS puan, "
            "sc.student_id AS student_id, sc.group_id AS group_id, gg.ad AS grup, gg.turnuva AS turnuva, "
            "gg.game_id AS game_id, g.ad AS oyun, s.ad_soyad AS ad_soyad, s.sinif AS sinif, s.sube AS sube, s.arsiv AS arsiv "
            "FROM group_scores sc JOIN game_groups gg ON gg.id=sc.group_id JOIN games g ON g.id=gg.game_id "
            "JOIN students s ON s.id=sc.student_id ORDER BY sc.tarih DESC, sc.id DESC", c)

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
    g = stats.groupby(["sinif", "sube"]).agg(
        mevcut=("id", "count"), aktif=("mac", lambda s: int((s > 0).sum())),
        galibiyet=("galibiyet", "sum"), beraberlik=("beraberlik", "sum"), mac=("mac", "sum")
    ).reset_index()
    g["basari"] = ((g.galibiyet + 0.5 * g.beraberlik) / g.mac.where(g.mac > 0) * 100).fillna(0).round(1)
    return g.sort_values(["sinif", "sube"]).reset_index(drop=True)

# ----------------------------------------------------------------------------
# KURUMSAL MİNİMALİST STİL (Çiğ Renkler ve Emojiler Arındırıldı)
# ----------------------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root{
  --bg:#F9F9F8; --card:#FFFFFF; --line:#E5E5E0; --side:#332D29;
  --primary:#8C6D53; --primary-dark:#6E533F; --text:#2C2C2A; --muted:#767670;
}
html, body, .stApp{font-family:'Inter',sans-serif; background:var(--bg); color:var(--text);}
header[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:2rem; max-width:1200px;}
h1,h2,h3,h4{font-family:'Inter',sans-serif !important; color:var(--text) !important; font-weight:700;}

section[data-testid="stSidebar"]{background:var(--side) !important;}
section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] label{color:#EAE6E1;}
.brand{padding:1rem 0 1.5rem 0; border-bottom:1px solid rgba(255,255,255,0.1); margin-bottom:1rem;}
.brand-t{font-size:1.1rem; font-weight:700; color:#FFF;}
.brand-s{font-size:0.7rem; color:#A89F96; text-transform:uppercase; letter-spacing:0.08em; margin-top:2px;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label{padding:0.5rem 0.7rem; border-radius:6px; transition:background 0.2s;}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover{background:rgba(255,255,255,0.05);}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked){background:var(--primary);}
section[data-testid="stSidebar"] div[role="radiogroup"] > label p{font-weight:600; font-size:0.9rem;}

.stat{background:var(--card); border:1px solid var(--line); border-radius:8px; padding:1.2rem; box-shadow:0 1px 3px rgba(0,0,0,0.02);}
.stat-label{font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; letter-spacing:0.05em;}
.stat-val{font-size:2rem; font-weight:700; color:var(--text); margin-top:0.4rem;}

.stButton > button{border-radius:6px; border:1px solid var(--line); background:var(--card); color:var(--text); font-weight:600;}
.stButton > button:hover{border-color:var(--primary); color:var(--primary-dark);}
.stButton > button[kind="primary"]{background:var(--primary) !important; color:#FFF !important; border:none !important;}

[data-testid="stTextInput"] input, [data-testid="stSelectbox"] div{border-radius:6px !important;}
</style>
"""

def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)

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
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            for name, df in sheets.items():
                df.to_excel(xw, sheet_name=name[:31], index=False)
    except Exception:
        return None
    return buf.getvalue()

def export_buttons(df, base_name, key, sheet="Veri"):
    stamp = date.today().strftime("%Y%m%d")
    c1, c2, _ = st.columns([1, 1, 3])
    xlsx = to_xlsx_bytes({sheet: df})
    if xlsx is not None:
        c1.download_button("Excel İndir", xlsx, file_name=f"{base_name}_{stamp}.xlsx", mime=XLSX_MIME, key=f"{key}_xlsx", use_container_width=True)
    c2.download_button("CSV İndir", to_csv_bytes(df), file_name=f"{base_name}_{stamp}.csv", mime="text/csv", key=f"{key}_csv", use_container_width=True)

# ----------------------------------------------------------------------------
# SAYFALAR
# ----------------------------------------------------------------------------
def page_dashboard():
    st.markdown("### Genel Bakış")
    students = load_students()
    matches = load_matches(active_only=True)
    if students.empty:
        st.info("Kayıtlı öğrenci bulunmuyor. Yönetici Paneli'nden öğrenci listesi yükleyebilirsiniz.")
        return
    stats = compute_stats(students, matches)
    summary = branch_summary(stats)
    
    col1, col2, col3, col4 = st.columns(4)
    col1.markdown(f'<div class="stat"><div class="stat-label">Toplam Öğrenci</div><div class="stat-val">{len(students)}</div></div>', unsafe_allow_html=True)
    col2.markdown(f'<div class="stat"><div class="stat-label">Toplam Maç</div><div class="stat-val">{len(matches)}</div></div>', unsafe_allow_html=True)
    col3.markdown(f'<div class="stat"><div class="stat-label">Şube Sayısı</div><div class="stat-val">{len(summary)}</div></div>', unsafe_allow_html=True)
    col4.markdown(f'<div class="stat"><div class="stat-label">Aktif Katılım</div><div class="stat-val">{int((stats.mac > 0).sum())}</div></div>', unsafe_allow_html=True)

def page_branch():
    st.markdown("### Şube Kontrol Paneli")
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
    if not show.empty:
        export_buttons(show, "sube_ogrenci_listesi", "br")

def page_match():
    st.markdown("### Maç Girişi ve Eşleştirme")
    students = load_students()
    if students.empty:
        st.info("Öğrenci yok.")
        return
    if not is_admin():
        st.warning("Salt Okunur Mod: Maç kaydetmek için Yönetici Paneli'nden giriş yapmalısınız.")
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
    if c1.button("1. Oyuncu Kazandı", disabled=not ready, use_container_width=True, type="primary" if ready else "secondary"):
        add_match(gid, p1, p2, "1", tarih)
        flash("success", "Maç kaydedildi.")
        st.rerun()
    if c2.button("Berabere", disabled=not ready, use_container_width=True):
        add_match(gid, p1, p2, "D", tarih)
        flash("success", "Maç kaydedildi.")
        st.rerun()
    if c3.button("2. Oyuncu Kazandı", disabled=not ready, use_container_width=True):
        add_match(gid, p1, p2, "2", tarih)
        flash("success", "Maç kaydedildi.")
        st.rerun()

def page_leaderboard():
    st.markdown("### Sıralama ve Geçmiş")
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
    if not show.empty:
        export_buttons(show, "siralama", "lb")

def page_individual():
    st.markdown("### Bireysel Oyunlar ve Karneler")
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
            st.info("Bu öğrenciye ait ilerleme kaydı bulunmuyor.")
        else:
            st.dataframe(cur[["oyun", "kart", "seviye", "tarih"]], hide_index=True, use_container_width=True)

def page_groups():
    st.markdown("### Grup ve Turnuva Yönetimi")
    groups = load_groups()
    if groups.empty:
        st.info("Henüz grup oluşturulmamış.")
    else:
        st.dataframe(groups[["oyun", "turnuva", "ad", "uye"]].rename(columns={"oyun": "Oyun", "turnuva": "Turnuva", "ad": "Grup Adı", "uye": "Üye"}), hide_index=True, use_container_width=True)

def page_library():
    st.markdown("### Oyun Kütüphanesi")
    games = load_games()
    for cat in GAME_LIBRARY.keys():
        sub = games[games.kategori == cat]
        st.markdown(f"#### {cat}")
        st.write(", ".join(sub.ad.tolist()))

def render_import_ui(key_prefix):
    if not is_admin():
        st.warning("Veri yüklemek için yönetici girişi yapmalısınız.")
        return
    st.markdown("#### Öğrenci Listesi Yükle (Excel/CSV)")
    up = st.file_uploader("Dosya Seç", type=["xlsx", "csv"], key=f"{key_prefix}_up")
    if up is not None:
        try:
            df = pd.read_excel(up) if up.name.endswith(".xlsx") else pd.read_csv(up)
            st.success(f"{len(df)} satır okundu.")
            st.dataframe(df.head(5), hide_index=True)
        except Exception as e:
            st.error(f"Hata: {e}")

def page_admin():
    st.markdown("### Yönetici Paneli")
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
    st.set_page_config(
        page_title="Akıl ve Zeka Oyunları", page_icon="🧠", layout="wide", initial_sidebar_state="auto"
    )
    init_db()
    inject_css()

    with st.sidebar:
        st.markdown(
            '<div class="brand"><div class="brand-t">Akıl ve Zeka Oyunları</div><div class="brand-s">Yönetim & Takip</div></div>',
            unsafe_allow_html=True,
        )
        page = st.radio("Menü", list(PAGES.keys()), label_visibility="collapsed", key="nav")
        if is_admin():
            st.markdown("<br><small style='color:#A89F96'>Oturum Açık</small>", unsafe_allow_html=True)
            if st.button("Çıkış Yap", use_container_width=True):
                admin_logout()
                st.rerun()
        else:
            st.markdown("<br><small style='color:#A89F96'>Salt Okunur Mod</small>", unsafe_allow_html=True)

    show_flash()
    PAGES[page]()

if __name__ == "__main__":
    main()

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
MAX_PIN_FAILS = 5
LOCK_SECONDS = 300
PAGE_SIZE = 20


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


def is_admin():
    return bool(st.session_state.get("admin"))


def admin_logout():
    st.session_state["admin"] = False
    st.session_state.pop("del_pending", None)


def _hash_pin(pin, salt_hex):
    return hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), bytes.fromhex(salt_hex), 120_000).hex()


def _store_pin(pin):
    salt = secrets.token_hex(16)
    _set_setting("admin_pin", f"{salt}\${_hash_pin(pin, salt)}")


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


def import_students(df, replace=False):
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


def load_games():
    with db() as c:
        return pd.read_sql_query("SELECT id, ad, kategori, varsayilan FROM games ORDER BY kategori, id", c)


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


def played_pairs(game_id):
    m = load_matches(game_id)
    return {frozenset((int(a), int(b))) for a, b in zip(m.p1_id, m.p2_id)}


def count_archived():
    with db() as c:
        return int(c.execute("SELECT COUNT(*) FROM students WHERE arsiv=1").fetchone()[0])


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
# İKONLAR & TASARIM (CSS)
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

.welcome{position:relative;overflow:hidden;display:flex;align-items:center;justify-content:space-between;gap:1.2rem;padding:2rem 2.2rem;margin-bottom:1.3rem;border-radius:16px;background:#FFFFFF;border:1px solid var(--line);box-shadow:var(--shadow);}
.welcome-t{font-family:'DM Serif Display',Georgia,serif;font-size:2.15rem;line-height:1.15;color:var(--espresso);}
.welcome-s{margin-top:.5rem;font-size:1.02rem;color:var(--muted);max-width:560px;}

.stat{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:1.1rem 1.25rem;box-shadow:var(--shadow);transition:transform .2s ease, box-shadow .2s ease;height:100%;}
.stat:hover{transform:translateY(-2px);box-shadow:var(--shadow-h);}
.stat-top{display:flex;align-items:center;justify-content:space-between;gap:.5rem;}
.stat-label{font-size:.74rem;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;font-weight:700;}
.stat-icon{width:32px;height:32px;border-radius:9px;background:var(--milk);display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0;}
.stat-val{font-size:2.3rem;font-weight:800;color:var(--espresso);line-height:1.1;margin-top:.55rem;}
.stat-sub{font-size:.78rem;color:var(--muted);margin-top:.3rem;}

[data-testid\$="olumn"]:has(.qa-body){position:relative;box-sizing:border-box;background:var(--card);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);padding:1.2rem 1.3rem !important;cursor:pointer;transition:transform .2s ease, box-shadow .2s ease, border-color .2s ease;}
[data-testid\$="olumn"]:has(.qa-body):hover{transform:translateY(-2px);box-shadow:var(--shadow-h);border-color:var(--taba);}
[data-testid\$="olumn"]:has(.qa-body) .element-container:has(.stButton),
[data-testid\$="olumn"]:has(.qa-body) [data-testid="stElementContainer"]:has(.stButton){position:absolute;inset:0;width:100% !important;height:100%;margin:0;z-index:3;}
[data-testid\$="olumn"]:has(.qa-body) .stButton{width:100%;height:100%;}
[data-testid\$="olumn"]:has(.qa-body) .stButton > button{width:100%;height:100%;min-height:0;opacity:0;cursor:pointer;border-radius:14px;}
.qa-body{display:flex;align-items:center;gap:1rem;min-height:74px;}
.qa-ico{width:52px;height:52px;border-radius:12px;background:var(--milk);border:1px solid var(--line);display:flex;align-items:center;justify-content:center;font-size:1.55rem;flex-shrink:0;}
.qa-t{font-weight:800;font-size:1.05rem;color:var(--espresso);}
.qa-d{font-size:.85rem;color:var(--muted);margin-top:.15rem;line-height:1.35;}
.qa-go{font-size:1.25rem;color:var(--taba);font-weight:800;}

.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button{border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--espresso);font-weight:700;transition:all .2s ease;}
.stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover{border-color:var(--taba);color:var(--taba-d);transform:translateY(-1px);box-shadow:0 6px 14px rgba(0,0,0,.06);}
.stButton > button[kind="primary"], [data-testid="stBaseButton-primary"], [data-testid="stFormSubmitButton"] > button[kind="primary"], [data-testid="stBaseButton-primaryFormSubmit"]{background:linear-gradient(135deg,#10B981,#059669) !important;color:#FFFFFF !important;border:none !important;}

.kvkk-footer {
  text-align: center;
  padding: 1.5rem;
  color: var(--muted);
  font-size: 0.8rem;
  border-top: 1px solid var(--line);
  margin-top: 3rem;
}
</style>
"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)


def hero(title, subtitle, icon_name):
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:1rem;margin-bottom:1.5rem;">'
        f'<div style="font-size:2rem;">{_ico(icon_name, 32)}</div>'
        f'<div><h2 style="margin:0;">{esc(title)}</h2><p style="margin:0;color:var(--muted);">{esc(subtitle)}</p></div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def welcome_banner():
    st.markdown(
        '<div class="welcome"><div class="welcome-main">'
        '<div class="welcome-eyebrow" style="color:#10B981;font-weight:800;letter-spacing:.1em;text-transform:uppercase;font-size:.75rem">Akıl ve Zeka Oyunları · Operasyon Paneli</div>'
        '<div class="welcome-t">Günaydın, Öğretmenim</div>'
        '<div class="welcome-s">Şubelerinizdeki turnuva akışını ve öğrencilerin kişisel gelişim süreçlerini tek platformda izleyin.</div>'
        '</div><div class="welcome-art" style="font-size:2.5rem">' + icon("puzzle", 54) + icon("dice", 54) + '</div></div>',
        unsafe_allow_html=True,
    )


def stat_cards(items):
    cols = st.columns(len(items))
    for col, item in zip(cols, items):
        icon_name, label, value = item[0], item[1], item[2]
        sub = item[3] if len(item) > 3 else ""
        col.markdown(
            f'<div class="stat"><div class="stat-top"><div class="stat-label">{esc(label)}</div>'
            f'<div class="stat-icon">{_ico(icon_name, 18)}</div></div>'
            f'<div class="stat-val">{esc(str(value))}</div>'
            + (f'<div class="stat-sub">{esc(sub)}</div>' if sub else "")
            + "</div>",
            unsafe_allow_html=True,
        )


def goto(page, match_tab=None):
    st.session_state["nav"] = page
    if match_tab:
        st.session_state["match_tab"] = match_tab


def reset_match_tab():
    st.session_state["match_tab"] = "manual"


def quick_action(col, key, icon_name, title, desc, page, match_tab=None):
    with col:
        st.markdown(
            f'<div class="qa-body"><div class="qa-ico">{_ico(icon_name, 26)}</div>'
            f'<div class="qa-txt"><div class="qa-t">{esc(title)}</div><div class="qa-d">{esc(desc)}</div></div>'
            f'<div class="qa-go">→</div></div>',
            unsafe_allow_html=True,
        )
        st.button(title, key=key, on_click=goto, args=(page, match_tab), use_container_width=True)


def bar_html(pct):
    return f'<div class="bar" style="height:6px;background:#F1F3F2;border-radius:99px;overflow:hidden"><div style="width:{max(0, min(100, pct))}%;height:100%;background:#10B981;border-radius:99px"></div></div>'


def flash(kind, msg):
    st.session_state["flash"] = (kind, msg)


def show_flash():
    f = st.session_state.pop("flash", None)
    if f:
        kind, msg = f
        {"success": st.success, "error": st.error, "warning": st.warning}.get(kind, st.info)(msg)


def empty_state():
    st.info("Henüz öğrenci listesi yüklenmemiş. Genel Bakış sayfasından veya Yönetici Paneli'nden Excel/CSV yükleyebilirsiniz.")


def student_labels(df):
    return {int(r.id): f"{r.ad_soyad}  ·  {r.sinif}-{r.sube}" for r in df.itertuples()}


def class_branch_filter(df, key):
    c1, c2 = st.columns(2)
    siniflar = ["Tümü"] + sorted(int(x) for x in df.sinif.unique())
    sel_s = c1.selectbox(
        "Sınıf", siniflar, key=f"{key}_sinif",
        format_func=lambda x: x if x == "Tümü" else f"{x}. Sınıf",
    )
    base = df if sel_s == "Tümü" else df[df.sinif == sel_s]
    subeler = ["Tümü"] + sorted(base.sube.unique().tolist(), key=tr_key)
    sel_b = c2.selectbox("Şube", sub

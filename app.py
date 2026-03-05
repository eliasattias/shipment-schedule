import base64
import os
import sqlite3
from pathlib import Path
from datetime import date

import pandas as pd
import requests
import streamlit as st

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = BASE_DIR / "orders.db"
LOGO_PATH = BASE_DIR / "assets" / "sensimedical-logo.png"
FILE_PATTERNS = ("Pending Orders *.csv", "Pending Orders *.xlsx", "*.csv", "*.xlsx")

# ── Resend email config ───────────────────────────────────────────────────────
RESEND_API_KEY = st.secrets.get("RESEND_API_KEY", os.getenv("RESEND_API_KEY", ""))
NOTIFY_EMAILS  = ["elias.a@sensimedical.com"]
NOTIFY_FROM    = "SensiMedical Schedule <schedule@sensimedical.com>"

SENSIMEDICAL_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

    /* ─── Reset & Base ───────────────────────────────────── */
    html, body, .stApp {
        background-color: #f4f6f9 !important;
        font-family: 'DM Sans', sans-serif !important;
    }

    /* ─── Hide Streamlit chrome ──────────────────────────── */
    #MainMenu, footer, header { visibility: hidden; }
    [data-testid="stSidebar"] { display: none; }
    [data-testid="stSidebar"] ~ div { margin-left: 0 !important; }
    [data-testid="stDecoration"] { display: none; }

    /* ─── Top Navigation Bar ─────────────────────────────── */
    .sm-navbar {
        position: fixed;
        top: 0; left: 0; right: 0;
        z-index: 999;
        background: #0c1f3a;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 2rem;
        height: 56px;
        border-bottom: 1px solid rgba(255,255,255,0.06);
        box-shadow: 0 2px 16px rgba(0,0,0,0.25);
    }
    .sm-navbar-brand {
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .sm-navbar-brand img { height: 28px; width: auto; }
    .sm-navbar-title {
        font-family: 'DM Sans', sans-serif;
        font-weight: 600;
        font-size: 0.85rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: rgba(255,255,255,0.55);
    }
    .sm-navbar-badge {
        background: linear-gradient(135deg, #0ea5e9, #0d9488);
        color: white;
        font-size: 0.7rem;
        font-weight: 600;
        letter-spacing: 0.06em;
        padding: 3px 10px;
        border-radius: 20px;
        text-transform: uppercase;
    }

    /* ─── Main content offset for fixed nav ─────────────── */
    .main .block-container {
        padding-top: 4rem !important;
        padding-left: 2.5rem !important;
        padding-right: 2.5rem !important;
        max-width: 1400px !important;
    }

    /* ─── Hero Header ────────────────────────────────────── */
    .sm-hero {
        text-align: center;
        padding: 0.6rem 1rem 0.8rem;
        margin-bottom: 0.8rem;
        border-bottom: 1px solid #e2e8f0;
    }
    .sm-hero h1 {
        font-family: 'DM Sans', sans-serif !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #0c1f3a !important;
        margin: 0 0 0.35rem 0 !important;
        letter-spacing: -0.03em;
    }
    .sm-hero-date {
        font-size: 1rem;
        color: #64748b;
        margin-bottom: 0.8rem;
        font-weight: 400;
    }
    .sm-hero-stats {
        display: inline-flex;
        gap: 0;
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.06);
        overflow: hidden;
    }
    .sm-stat {
        padding: 0.85rem 2rem;
        border-right: 1px solid #e2e8f0;
        min-width: 120px;
    }
    .sm-stat:last-child { border-right: none; }
    .sm-stat-value {
        font-family: 'DM Mono', monospace;
        font-size: 1.6rem;
        font-weight: 600;
        color: #0c1f3a;
        line-height: 1;
        margin-bottom: 0.2rem;
    }
    .sm-stat-value.teal  { color: #0d9488; }
    .sm-stat-value.amber { color: #f59e0b; }
    .sm-stat-value.green { color: #10b981; }
    .sm-stat-label {
        font-size: 0.68rem;
        font-weight: 600;
        letter-spacing: 0.09em;
        text-transform: uppercase;
        color: #94a3b8;
    }

    /* ─── Stat Cards ─────────────────────────────────────── */
    .sm-cards {
        display: flex;
        gap: 1rem;
        margin-bottom: 1rem;
    }
    .sm-card {
        flex: 1;
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem 1.3rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        position: relative;
        overflow: hidden;
    }
    .sm-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
    }
    .sm-card.blue::before  { background: linear-gradient(90deg, #0c1f3a, #2d5a87); }
    .sm-card.teal::before  { background: linear-gradient(90deg, #0d9488, #0ea5e9); }
    .sm-card.amber::before { background: linear-gradient(90deg, #f59e0b, #f97316); }
    .sm-card.green::before { background: linear-gradient(90deg, #10b981, #0d9488); }
    .sm-card-label {
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.09em;
        text-transform: uppercase;
        color: #94a3b8;
        margin-bottom: 0.3rem;
    }
    .sm-card-value {
        font-size: 1.5rem;
        font-weight: 600;
        color: #0c1f3a;
        font-family: 'DM Mono', monospace;
        line-height: 1;
    }
    .sm-card-sub {
        font-size: 0.75rem;
        color: #94a3b8;
        margin-top: 0.2rem;
    }

    /* ─── Table wrapper ──────────────────────────────────── */
    .sm-table-wrapper {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.2rem 1.4rem 1rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        margin-bottom: 1.2rem;
    }
    .sm-table-label {
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.09em;
        text-transform: uppercase;
        color: #64748b;
        margin-bottom: 0.8rem;
        display: flex;
        align-items: center;
        gap: 6px;
    }

    /* ─── Data editor polish ─────────────────────────────── */
    [data-testid="stDataEditor"] {
        border-radius: 8px !important;
        border: 1px solid #e2e8f0 !important;
        font-family: 'DM Sans', sans-serif !important;
        font-size: 0.85rem !important;
        background: #ffffff !important;
    }
    [data-testid="stDataEditor"] th {
        background: #f8fafc !important;
        color: #475569 !important;
        font-size: 0.72rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.07em !important;
        text-transform: uppercase !important;
        border-bottom: 1px solid #e2e8f0 !important;
    }
    [data-testid="stDataEditor"] td {
        color: #0f172a !important;
        border-bottom: 1px solid #f1f5f9 !important;
    }

    /* Center alignment for Row, Cases #, Created Date, Scheduled date columns */
    [data-testid="stDataEditor"] table tbody tr td:nth-child(1),
    [data-testid="stDataEditor"] table thead tr th:nth-child(1),
    [data-testid="stDataEditor"] table tbody tr td:nth-child(3),
    [data-testid="stDataEditor"] table thead tr th:nth-child(3),
    [data-testid="stDataEditor"] table tbody tr td:nth-child(5),
    [data-testid="stDataEditor"] table thead tr th:nth-child(5),
    [data-testid="stDataEditor"] table tbody tr td:nth-child(6),
    [data-testid="stDataEditor"] table thead tr th:nth-child(6) {
        text-align: center !important;
    }
    /* Narrow Row index column */
    [data-testid="stDataEditor"] table thead tr th:nth-child(1),
    [data-testid="stDataEditor"] table tbody tr td:nth-child(1) {
        width: 28px !important;
        min-width: 28px !important;
        max-width: 28px !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
        font-size: 0.7rem !important;
        color: #94a3b8 !important;
        text-align: center !important;
    }

    /* ─── Selectbox / inputs ─────────────────────────────── */
    .stSelectbox > div > div {
        background: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 8px !important;
        font-family: 'DM Sans', sans-serif !important;
        font-size: 0.87rem !important;
        color: #0f172a !important;
    }
    .stTextInput input {
        background: #f8fafc !important;
        border-radius: 8px !important;
    }

    /* ─── Save Button ────────────────────────────────────── */
    .stButton > button {
        background: linear-gradient(135deg, #0c1f3a 0%, #1e3a5f 100%) !important;
        color: white !important;
        border: none !important;
        font-family: 'DM Sans', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        letter-spacing: 0.04em !important;
        border-radius: 8px !important;
        padding: 0.55rem 1.8rem !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 2px 8px rgba(12,31,58,0.25) !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #0d9488 0%, #0ea5e9 100%) !important;
        box-shadow: 0 4px 16px rgba(13,148,136,0.3) !important;
        transform: translateY(-1px) !important;
    }

    /* ─── Alert overrides ────────────────────────────────── */
    [data-testid="stSuccess"] {
        background: #f0fdf9 !important;
        border-left: 3px solid #0d9488 !important;
        border-radius: 8px !important;
        font-family: 'DM Sans', sans-serif !important;
    }
    [data-testid="stInfo"] {
        background: #f0f9ff !important;
        border-left: 3px solid #0ea5e9 !important;
        border-radius: 8px !important;
        font-family: 'DM Sans', sans-serif !important;
    }
    [data-testid="stWarning"] {
        background: #fffbeb !important;
        border-left: 3px solid #f59e0b !important;
        border-radius: 8px !important;
        font-family: 'DM Sans', sans-serif !important;
    }
    [data-testid="stError"] {
        background: #fef2f2 !important;
        border-left: 3px solid #ef4444 !important;
        border-radius: 8px !important;
    }

    /* ─── Caption / small text ───────────────────────────── */
    .stCaption, [data-testid="stCaptionContainer"] {
        font-family: 'DM Sans', sans-serif !important;
        color: #94a3b8 !important;
        font-size: 0.78rem !important;
    }

    /* ─── Suppress default Streamlit title (we render our own) ── */
    h1:first-of-type { display: none; }
    h2, h3 {
        color: #0c1f3a !important;
        font-weight: 600 !important;
        letter-spacing: -0.01em;
    }
</style>
"""


def get_all_pending_files() -> list[Path]:
    if not DATA_DIR.exists():
        return []
    seen = set()
    files = []
    for pattern in FILE_PATTERNS:
        for p in DATA_DIR.glob(pattern):
            if p.name not in seen:
                seen.add(p.name)
                files.append(p)
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)


def get_latest_file() -> Path | None:
    files = get_all_pending_files()
    return files[0] if files else None


def init_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS followup_overrides (
            order_key TEXT PRIMARY KEY,
            follow_up TEXT,
            comments TEXT
        )
        """
    )
    try:
        conn.execute("ALTER TABLE followup_overrides ADD COLUMN comments TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        conn.rollback()
    conn.commit()
    return conn


def load_base_data(path: Path | None = None) -> pd.DataFrame:
    if path is None:
        path = get_latest_file()
    if not path:
        return pd.DataFrame()

    if path.suffix.lower() == ".xlsx":
        df = pd.read_excel(path, engine="openpyxl")
    else:
        df = pd.read_csv(path)

    # Normalize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Parse dates
    if "Created Date" in df.columns:
        df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")

    # Scheduled date
    date_col = next((c for c in df.columns if c in ("Follow up", "Scheduled Date", "Scheduled date")), None)
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.date
        if date_col != "Scheduled date":
            df = df.rename(columns={date_col: "Scheduled date"})
    if "Scheduled date" not in df.columns:
        df["Scheduled date"] = pd.NaT
    if "Comments" not in df.columns:
        df["Comments"] = ""

    if "SO Number" in df.columns and "Mfg Ref" in df.columns:
        df["_cust_key"] = df["Customer"].astype(str).str.strip()
        df["_created_key"] = df["Created Date"]
        qty = pd.to_numeric(df["Qty Order"], errors="coerce").fillna(0)
        price = pd.to_numeric(df["Sales Price"], errors="coerce").fillna(0)
        df["_sales_line"] = qty * price
        grouped = df.groupby(["_cust_key", "_created_key"], as_index=False).agg(
            Customer=("_cust_key", "first"),
            **{"Created Date": ("_created_key", "first")},
            Cases_num=("Qty Order", "sum"),
            Sales=("_sales_line", "sum"),
        )
        grouped = grouped.rename(columns={"Cases_num": "Cases #"})
        grouped["Sales"] = grouped["Sales"].round(2)
        grouped["Created Date"] = pd.to_datetime(grouped["Created Date"], errors="coerce")
        grouped["Scheduled date"] = pd.NaT
        grouped["Comments"] = ""
        grouped["order_key"] = (
            grouped["Customer"].astype(str).str.strip()
            + "|"
            + grouped["Created Date"].dt.strftime("%Y-%m-%d")
        )
        df = grouped[
            ["Customer", "Cases #", "Sales", "Created Date", "Scheduled date", "Comments", "order_key"]
        ]
    else:
        df["order_key"] = (
            df["Customer"].astype(str).str.strip()
            + "|" + df["Created Date"].dt.strftime("%Y-%m-%d")
        )

    return df


def apply_overrides(df: pd.DataFrame, conn: sqlite3.Connection) -> pd.DataFrame:
    NEW_ORDER_COMMENT = "estimated date??"
    PAST_DUE_COMMENT = "past due. please explain."
    today = date.today()
    try:
        overrides = pd.read_sql_query(
            "SELECT order_key, follow_up, comments FROM followup_overrides", conn
        )
    except sqlite3.OperationalError:
        overrides = pd.read_sql_query(
            "SELECT order_key, follow_up FROM followup_overrides", conn
        )
        overrides["comments"] = ""
    if not overrides.empty:
        overrides["follow_up"] = pd.to_datetime(overrides["follow_up"], errors="coerce").dt.date
        key_to_date = {}
        key_to_comments = {}
        for _, row in overrides.iterrows():
            k = row["order_key"]
            key_to_date[k] = row["follow_up"]
            key_to_comments[k] = str(row.get("comments") or "")
            parts = k.split("|")
            if len(parts) >= 2:
                k_cd = "|".join(parts[:2])
                if k_cd not in key_to_date:
                    key_to_date[k_cd] = row["follow_up"]
                    key_to_comments[k_cd] = str(row.get("comments") or "")
        df["Scheduled date"] = df["order_key"].map(key_to_date).combine_first(df["Scheduled date"])
        df["Comments"] = df["order_key"].map(key_to_comments).combine_first(df["Comments"].fillna("").astype(str))
        df["Scheduled date"] = pd.to_datetime(df["Scheduled date"], errors="coerce").dt.date
    # ── Auto-fill comments ────────────────────────────────────────────────────
    # 1. Past-due: has a scheduled date strictly before today
    past_due_mask = df["Scheduled date"].apply(
        lambda d: (
            d is not None
            and not pd.isnull(d)
            and isinstance(d, date)
            and d < today
        )
    )
    # Only overwrite comment if it's blank, the new-order placeholder, or already past-due text
    auto_comment = df["Comments"].fillna("").astype(str).str.strip().isin(
        {"", NEW_ORDER_COMMENT, PAST_DUE_COMMENT}
    )
    df.loc[past_due_mask & auto_comment, "Comments"] = PAST_DUE_COMMENT

    # 2. New orders (no date, no comment yet)
    new_order = df["Scheduled date"].isna()
    empty_comment = df["Comments"].fillna("").astype(str).str.strip() == ""
    df.loc[new_order & empty_comment, "Comments"] = NEW_ORDER_COMMENT

    # Flag used by JS to highlight rows yellow
    df["_past_due"] = past_due_mask

    return df


def save_overrides(
    original: pd.DataFrame, edited: pd.DataFrame, conn: sqlite3.Connection
) -> int:
    cols = ["order_key", "Scheduled date", "Comments"]
    for c in cols:
        if c not in edited.columns:
            return 0
    orig = original[cols].rename(columns={"Scheduled date": "_sd_orig", "Comments": "_com_orig"})
    merged = edited[cols].merge(
        orig[["order_key", "_sd_orig", "_com_orig"]],
        on="order_key",
    )
    merged["_sd_str"] = merged["Scheduled date"].astype(str)
    merged["_com_str"] = merged["Comments"].fillna("").astype(str)
    changed = merged[
        (merged["_sd_str"] != merged["_sd_orig"].astype(str))
        | (merged["_com_str"] != merged["_com_orig"].fillna("").astype(str))
    ]

    if changed.empty:
        return 0

    NEW_ORDER_COMMENT = "estimated date??"
    PAST_DUE_COMMENT = "past due. please explain."
    today = date.today()

    cur = conn.cursor()
    for _, row in changed.iterrows():
        sd = row["Scheduled date"]
        comment = str(row["Comments"] or "")
        # If a date is now set (and it's not past-due itself) and the comment is
        # still one of the auto-filled placeholders, clear it automatically.
        date_is_current = pd.notna(sd) and (
            not isinstance(sd, date) or sd >= today
        )
        if date_is_current and comment.strip() in (NEW_ORDER_COMMENT, PAST_DUE_COMMENT):
            comment = ""
        cur.execute(
            """
            INSERT INTO followup_overrides (order_key, follow_up, comments)
            VALUES (?, ?, ?)
            ON CONFLICT(order_key) DO UPDATE SET
                follow_up=excluded.follow_up,
                comments=excluded.comments
            """,
            (row["order_key"], str(sd) if pd.notna(sd) else None, comment),
        )
    conn.commit()
    return len(changed)


def send_update_email(df: pd.DataFrame) -> tuple[bool, str]:
    """Send a schedule-update notification via Resend."""
    if not RESEND_API_KEY:
        return False, "RESEND_API_KEY is not set. Add it to Streamlit secrets or a local .env file."

    today = date.today().strftime("%B %d, %Y")
    total = len(df)
    scheduled = int(df["Scheduled date"].notna().sum())
    past_due = int(df.get("_past_due", pd.Series(dtype=bool)).sum())

    # Build a simple HTML table of all orders
    rows_html = ""
    for _, row in df.iterrows():
        sd = row.get("Scheduled date", "")
        comment = str(row.get("Comments", "") or "")
        past = bool(row.get("_past_due", False))
        bg = ' style="background:#fff1f2;"' if past else ""
        rows_html += (
            f"<tr{bg}>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #e2e8f0;'>{row.get('Customer','')}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #e2e8f0;text-align:center;'>{row.get('Cases #','')}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #e2e8f0;text-align:center;'>{sd}</td>"
            f"<td style='padding:6px 10px;border-bottom:1px solid #e2e8f0;'>{comment}</td>"
            f"</tr>"
        )

    html_body = f"""
    <div style="font-family:'DM Sans',Arial,sans-serif;max-width:720px;margin:0 auto;">
      <div style="background:#0c1f3a;padding:18px 24px;border-radius:8px 8px 0 0;">
        <span style="color:white;font-size:1.1rem;font-weight:600;letter-spacing:-0.01em;">
          SensiMedical — Shipment Schedule Update
        </span>
      </div>
      <div style="background:#f4f6f9;padding:16px 24px;border-bottom:1px solid #e2e8f0;">
        <span style="color:#64748b;font-size:0.85rem;">{today}</span>
        &nbsp;&nbsp;·&nbsp;&nbsp;
        <span style="color:#0c1f3a;font-weight:600;">{total} orders</span>
        &nbsp;&nbsp;·&nbsp;&nbsp;
        <span style="color:#0d9488;font-weight:600;">{scheduled} scheduled</span>
        {"&nbsp;&nbsp;·&nbsp;&nbsp;<span style='color:#ef4444;font-weight:600;'>" + str(past_due) + " past due</span>" if past_due else ""}
      </div>
      <table style="width:100%;border-collapse:collapse;background:#ffffff;">
        <thead>
          <tr style="background:#f8fafc;">
            <th style="padding:8px 10px;text-align:left;font-size:0.72rem;letter-spacing:0.07em;text-transform:uppercase;color:#475569;border-bottom:2px solid #e2e8f0;">Customer</th>
            <th style="padding:8px 10px;text-align:center;font-size:0.72rem;letter-spacing:0.07em;text-transform:uppercase;color:#475569;border-bottom:2px solid #e2e8f0;">Cases #</th>
            <th style="padding:8px 10px;text-align:center;font-size:0.72rem;letter-spacing:0.07em;text-transform:uppercase;color:#475569;border-bottom:2px solid #e2e8f0;">Scheduled Date</th>
            <th style="padding:8px 10px;text-align:left;font-size:0.72rem;letter-spacing:0.07em;text-transform:uppercase;color:#475569;border-bottom:2px solid #e2e8f0;">Comments</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
        </tbody>
      </table>
      <div style="background:#f8fafc;padding:14px 24px;border-radius:0 0 8px 8px;border-top:1px solid #e2e8f0;display:flex;align-items:center;justify-content:space-between;">
        <span style="color:#94a3b8;font-size:0.75rem;">Sent automatically from SensiMedical Shipment Console</span>
        <a href="https://sensimedical-shipment-schedule.streamlit.app/" style="display:inline-block;background:linear-gradient(135deg,#0c1f3a,#1e3a5f);color:white;font-size:0.75rem;font-weight:600;text-decoration:none;padding:6px 14px;border-radius:6px;letter-spacing:0.03em;">View Console →</a>
      </div>
    </div>
    """

    try:
        resp = requests.post(
            "https://api.resend.com/emails",
            headers={
                "Authorization": f"Bearer {RESEND_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "from": NOTIFY_FROM,
                "to": NOTIFY_EMAILS,
                "subject": f"Schedule Update — {today} ({total} orders, {scheduled} scheduled)",
                "html": html_body,
            },
            timeout=10,
        )
        if resp.status_code in (200, 201):
            return True, "Email sent successfully."
        return False, f"Resend error {resp.status_code}: {resp.text}"
    except Exception as exc:
        return False, f"Request failed: {exc}"


def render_navbar(logo_path: Path) -> None:
    logo_src = ""
    if logo_path.exists():
        raw = logo_path.read_bytes()
        b64 = base64.b64encode(raw).decode()
        mime = "image/png" if logo_path.suffix.lower() == ".png" else "image/jpeg"
        logo_src = f"data:{mime};base64,{b64}"

    logo_html = (
        f'<img src="{logo_src}" alt="SensiMedical" />'
        if logo_src
        else '<span style="color:white;font-weight:700;font-size:1rem;letter-spacing:-0.02em;">SensiMedical</span>'
    )

    st.markdown(
        f"""
        <div class="sm-navbar">
            <div class="sm-navbar-brand">
                {logo_html}
                <span class="sm-navbar-title">Shipment Console</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_hero(df: pd.DataFrame) -> None:
    today_str = date.today().strftime("%B %d, %Y")
    st.markdown(
        f"""
        <div class="sm-hero">
            <h1>SensiMedical Shipment Schedule</h1>
            <div class="sm-hero-date">{today_str}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    # Stat cards directly below the hero title
    render_stat_cards(df)


def render_stat_cards(df: pd.DataFrame) -> None:
    total = len(df)
    scheduled = int(df["Scheduled date"].notna().sum())
    unscheduled = total - scheduled
    pct = int(scheduled / total * 100) if total > 0 else 0

    if "Sales" in df.columns:
        sales_val = pd.to_numeric(df["Sales"], errors="coerce").fillna(0).sum()
        sales_display = f"${sales_val:,.0f}"
    else:
        sales_display = "—"

    st.markdown(
        f"""
        <div class="sm-cards">
            <div class="sm-card blue">
                <div class="sm-card-label">Total Orders</div>
                <div class="sm-card-value">{total}</div>
                <div class="sm-card-sub">pending shipments</div>
            </div>
            <div class="sm-card teal">
                <div class="sm-card-label">Scheduled</div>
                <div class="sm-card-value">{scheduled}</div>
                <div class="sm-card-sub">{pct}% of orders</div>
            </div>
            <div class="sm-card amber">
                <div class="sm-card-label">Needs Date</div>
                <div class="sm-card-value">{unscheduled}</div>
                <div class="sm-card-sub">awaiting schedule</div>
            </div>
            <div class="sm-card green">
                <div class="sm-card-label">Total Sales</div>
                <div class="sm-card-value">{sales_display}</div>
                <div class="sm-card-sub">order value</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="SensiMedical™ – Pending Orders",
        layout="wide",
        page_icon="📦",
    )
    st.markdown(SENSIMEDICAL_CSS, unsafe_allow_html=True)

    if not DATA_DIR.exists():
        render_navbar(LOGO_PATH)
        st.error(f"**`data` folder not found.** Please create: `{DATA_DIR}` and put your daily CSV there.")
        return

    conn = init_db()
    all_files = get_all_pending_files()

    if not all_files:
        render_navbar(LOGO_PATH)
        st.warning("No 'Pending Orders' file found in `data/`. Expected: `Pending Orders *.csv` or `*.xlsx`")
        return

    # Always load the most recent file
    latest_path = all_files[0]
    base_df = load_base_data(latest_path)
    if base_df.empty:
        render_navbar(LOGO_PATH)
        return

    df = apply_overrides(base_df.copy(), conn)

    # Sort by Created Date (oldest first)
    if "Created Date" in df.columns:
        df = df.sort_values("Created Date", ascending=True, kind="mergesort").reset_index(drop=True)
        df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce").dt.date

    # Add 1-based row number
    df.insert(0, "Row", range(1, len(df) + 1))

    # Ensure Sales is numeric
    if "Sales" in df.columns:
        df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")

    # ─── Navbar ──────────────────────────────────────────────
    render_navbar(LOGO_PATH)

    # ─── Hero header + stats ──────────────────────────────────
    render_hero(df)

    # ─── Table ───────────────────────────────────────────────
    st.markdown(
        '<div class="sm-table-wrapper">'
        '<div class="sm-table-label"><span>🗓️</span> Orders · Edit Scheduled Date &amp; Comments</div>',
        unsafe_allow_html=True,
    )

    display_df = df.drop(columns=["order_key"], errors="ignore").copy()

    # Add a visible past-due status column — reliable without any JS
    display_df.insert(
        display_df.columns.get_loc("Scheduled date") + 1,
        "⚠️",
        display_df.index.map(
            lambda i: "⚠️ Past Due" if ("_past_due" in df.columns and df.at[i, "_past_due"]) else ""
        ),
    )
    display_df = display_df.drop(columns=["_past_due"], errors="ignore")

    column_config = {
        "Row": st.column_config.NumberColumn("·", disabled=True, width="small"),
        "Cases #": st.column_config.NumberColumn("Cases #", format="%,d", disabled=True),
        "Sales": st.column_config.NumberColumn("Sales", format="$%,.2f", disabled=True),
        "Scheduled date": st.column_config.DateColumn("Scheduled date"),
        "⚠️": st.column_config.TextColumn("⚠️", disabled=True, width="small"),
        "Comments": st.column_config.TextColumn("Comments", width="large"),
    }
    for col in display_df.columns:
        if col not in ("Row", "Cases #", "Sales", "Scheduled date", "⚠️", "Comments"):
            column_config[col] = st.column_config.Column(col, disabled=True)

    edited_display = st.data_editor(
        display_df,
        column_config=column_config,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        height=700,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # Reattach order_key for save logic; drop the display-only status column
    edited_df = edited_display.drop(columns=["⚠️"], errors="ignore").copy()
    edited_df["order_key"] = df["order_key"].values

    # ─── Save + Send Update ──────────────────────────────────
    col_save, col_mid, col_send = st.columns([2, 6, 2])
    with col_save:
        if st.button("💾  Save Changes", use_container_width=True):
            n = save_overrides(base_df, edited_df, conn)
            if n > 0:
                st.success(f"✓ Saved {n} updated row(s).")
            else:
                st.info("No changes to save.")
    with col_mid:
        st.caption("Changes are stored locally and applied automatically across file versions.")
    with col_send:
        if st.button("✉  Send Update", use_container_width=True, key="send_update"):
            ok, msg = send_update_email(df)
            if ok:
                st.toast("✓ Update email sent!", icon="✉️")
            else:
                st.toast(f"Failed: {msg}", icon="⚠️")


if __name__ == "__main__":
    main()
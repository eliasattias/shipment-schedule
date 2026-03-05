import base64
import sqlite3
from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = BASE_DIR / "orders.db"
LOGO_PATH = BASE_DIR / "assets" / "sensimedical-logo.png"
FILE_PATTERNS = ("Pending Orders *.csv", "Pending Orders *.xlsx", "*.csv", "*.xlsx")

# SensiMedical theme – using DM Sans / DM Mono and prettier layout
SENSIMEDICAL_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

    /* Reset & base */
    html, body, .stApp {
        background-color: #f4f6f9 !important;
        font-family: 'DM Sans', sans-serif !important;
    }

    /* Hide Streamlit chrome */
    #MainMenu, footer, header { visibility: hidden; }
    [data-testid="stSidebar"] { display: none; }
    [data-testid="stSidebar"] ~ div { margin-left: 0 !important; }
    [data-testid="stDecoration"] { display: none; }

    /* Top header bar – reuse existing sensimedical-header container */
    .sensimedical-header {
        background: linear-gradient(90deg, #0c1f3a 0%, #2d5a87 100%);
        padding: 0.6rem 1.5rem;
        margin-left: calc(-50vw + 50%);
        margin-right: calc(-50vw + 50%);
        margin-bottom: 1rem;
        width: 100vw;
        box-sizing: border-box;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 2px 16px rgba(0,0,0,0.25);
    }
    .sensimedical-header img {
        height: 32px;
        width: auto;
        display: block;
    }

    /* Main content container offset */
    .main .block-container {
        padding-top: 1.5rem !important;
        padding-left: 2.5rem !important;
        padding-right: 2.5rem !important;
        max-width: 1400px !important;
    }

    /* Headings */
    h1, h2, h3 {
        color: #0c1f3a !important;
        font-weight: 600 !important;
        letter-spacing: -0.01em;
    }

    /* Data editor polish */
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

    /* Center alignment for specific columns in main table:
       Row (1), Cases # (3), Created Date (5), Scheduled date (6) */
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
    /* Make Row column thin */
    [data-testid="stDataEditor"] table thead tr th:nth-child(1),
    [data-testid="stDataEditor"] table tbody tr td:nth-child(1) {
        width: 3rem !important;
        max-width: 3rem !important;
        padding-left: 0.25rem;
        padding-right: 0.25rem;
    }

    /* Selectbox / inputs */
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

    /* Primary button */
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

    /* Alerts */
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

    /* Caption / small text */
    .stCaption, [data-testid="stCaptionContainer"] {
        font-family: 'DM Sans', sans-serif !important;
        color: #94a3b8 !important;
        font-size: 0.78rem !important;
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

    # Parse dates (Excel may already give datetime)
    if "Created Date" in df.columns:
        df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce")

    # Scheduled date: daily files typically don't have it; we add it and fill from DB later
    date_col = next((c for c in df.columns if c in ("Follow up", "Scheduled Date", "Scheduled date")), None)
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.date
        if date_col != "Scheduled date":
            df = df.rename(columns={date_col: "Scheduled date"})
    if "Scheduled date" not in df.columns:
        df["Scheduled date"] = pd.NaT
    if "Comments" not in df.columns:
        df["Comments"] = ""

    # Build a stable order key and optionally group to the clean summary format.
    if "SO Number" in df.columns and "Mfg Ref" in df.columns:
        # Original/detail file: group by Customer + Created Date so the summary
        # matches the clean Pending Orders layout (Customer, Cases #, Sales, Created Date).
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
        grouped["Created Date"] = pd.to_datetime(
            grouped["Created Date"], errors="coerce"
        )
        grouped["Scheduled date"] = pd.NaT
        grouped["Comments"] = ""
        # order_key = Customer | YYYY-MM-DD (same as summary files)
        grouped["order_key"] = (
            grouped["Customer"].astype(str).str.strip()
            + "|"
            + grouped["Created Date"].dt.strftime("%Y-%m-%d")
        )
        df = grouped[
            ["Customer", "Cases #", "Sales", "Created Date", "Scheduled date", "Comments", "order_key"]
        ]
    else:
        # Pending Orders summary: Customer + Created Date (qty can change on partial dispatch)
        def _norm_num(val):
            if pd.isna(val):
                return ""
            s = str(val).strip().replace("$", "").replace(",", "").strip()
            try:
                return str(int(float(s)))
            except (ValueError, TypeError):
                return s

        df["order_key"] = (
            df["Customer"].astype(str).str.strip()
            + "|" + df["Created Date"].dt.strftime("%Y-%m-%d")
        )

    return df


def apply_overrides(df: pd.DataFrame, conn: sqlite3.Connection) -> pd.DataFrame:
    NEW_ORDER_COMMENT = "estimated date??"
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
        # Build lookup: support both older keys that included extra parts and the new
        # 2-part Customer|CreatedDate keys
        key_to_date = {}
        key_to_comments = {}
        for _, row in overrides.iterrows():
            k = row["order_key"]
            key_to_date[k] = row["follow_up"]
            key_to_comments[k] = str(row.get("comments") or "")
            parts = k.split("|")
            # If key had extra parts (e.g. Customer|Date|Cases|Sales), also map the
            # first two segments (Customer|Date) used by the new summary keys.
            if len(parts) >= 2:
                k_cd = "|".join(parts[:2])
                if k_cd not in key_to_date:
                    key_to_date[k_cd] = row["follow_up"]
                    key_to_comments[k_cd] = str(row.get("comments") or "")
        df["Scheduled date"] = df["order_key"].map(key_to_date).combine_first(df["Scheduled date"])
        df["Comments"] = df["order_key"].map(key_to_comments).combine_first(df["Comments"].fillna("").astype(str))
        # Ensure Scheduled date is a proper date type for display
        df["Scheduled date"] = pd.to_datetime(df["Scheduled date"], errors="coerce").dt.date
    # New orders (no saved scheduled date): auto-fill Comments so team knows to set a date
    new_order = df["Scheduled date"].isna()
    empty_comment = (df["Comments"].fillna("").astype(str).str.strip() == "")
    df.loc[new_order & empty_comment, "Comments"] = NEW_ORDER_COMMENT
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

    cur = conn.cursor()
    for _, row in changed.iterrows():
        sd = row["Scheduled date"]
        cur.execute(
            """
            INSERT INTO followup_overrides (order_key, follow_up, comments)
            VALUES (?, ?, ?)
            ON CONFLICT(order_key) DO UPDATE SET
                follow_up=excluded.follow_up,
                comments=excluded.comments
            """,
            (row["order_key"], str(sd) if pd.notna(sd) else None, str(row["Comments"] or "")),
        )
    conn.commit()
    return len(changed)


def render_top_header() -> None:
    """Render SensiMedical top header bar with logo (same convention as sales-lot-tool)."""
    if not LOGO_PATH.exists():
        return
    raw = LOGO_PATH.read_bytes()
    b64 = base64.b64encode(raw).decode()
    mime = "image/png" if LOGO_PATH.suffix.lower() == ".png" else "image/jpeg"
    src = f"data:{mime};base64,{b64}"
    st.markdown(
        f'<div class="sensimedical-header"><img src="{src}" alt="SensiMedical" /></div>',
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="SensiMedical™ – Pending Orders",
        layout="wide",
        page_icon="📦",
    )
    st.markdown(SENSIMEDICAL_CSS, unsafe_allow_html=True)
    render_top_header()

    st.title("Pending Orders – Schedule Manager")
    st.caption("SensiMedical™ Shipment Schedule")

    if not DATA_DIR.exists():
        st.error(
            f"`data` folder not found.\n\n"
            f"Please create: `{DATA_DIR}` and put your daily CSV there."
        )
        return

    conn = init_db()
    all_files = get_all_pending_files()
    if not all_files:
        st.warning(
            "No 'Pending Orders' file found in the `data` folder.\n\n"
            "Expected: `Pending Orders *.csv` or `Pending Orders *.xlsx`"
        )
        return

    # Let user pick which file to load (e.g. load March 3 to save dates, then March 4 to see them)
    file_options = [f.name for f in all_files]
    default_idx = 0
    selected_name = st.selectbox(
        "File to load",
        file_options,
        index=default_idx,
        help="Choose which pending orders file to view. Use the latest for today; pick an older file to copy its dates into the app and Save, then load the latest again.",
    )
    selected_path = all_files[file_options.index(selected_name)]
    base_df = load_base_data(selected_path)
    if base_df.empty:
        return

    df = apply_overrides(base_df.copy(), conn)
    # Sort by Created Date (oldest first) by default
    if "Created Date" in df.columns:
        df = df.sort_values(
            "Created Date", ascending=True, kind="mergesort"
        ).reset_index(drop=True)
        # Display Created Date without time portion
        df["Created Date"] = pd.to_datetime(
            df["Created Date"], errors="coerce"
        ).dt.date
    # Add a 1-based row number for readability
    df.insert(0, "Row", range(1, len(df) + 1))
    # Ensure Sales is numeric so we can format as currency
    if "Sales" in df.columns:
        df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")

    st.write("Edit **Scheduled date** and **Comments** below.")

    # Show table without order_key (internal use only); only Scheduled date and Comments are editable
    display_df = df.drop(columns=["order_key"])
    column_config = {
        "Row": st.column_config.NumberColumn(
            "Row", disabled=True, width="small"
        ),
        "Sales": st.column_config.NumberColumn(
            "Sales", format="$%.2f", disabled=True
        ),
        "Scheduled date": st.column_config.DateColumn("Scheduled date"),
        "Comments": st.column_config.TextColumn("Comments", width="large"),
    }
    for col in display_df.columns:
        if col not in ("Row", "Sales", "Scheduled date", "Comments"):
            column_config[col] = st.column_config.Column(col, disabled=True)
    edited_display = st.data_editor(
        display_df,
        column_config=column_config,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        height=700,
    )
    # Reattach order_key for save logic
    edited_df = edited_display.copy()
    edited_df["order_key"] = df["order_key"].values

    if st.button("Save changes"):
        n = save_overrides(base_df, edited_df, conn)
        if n > 0:
            st.success(f"Saved {n} updated row(s).")
        else:
            st.info("No changes to save.")


if __name__ == "__main__":
    main()

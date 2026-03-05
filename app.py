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

# SensiMedical theme – aligned with sensimedical.com / sales-lot-tool
SENSIMEDICAL_CSS = """
<style>
    /* Main – clean white */
    .stApp { background-color: #ffffff; }
    /* Top header bar – same convention as SensiMedical sales lot tool */
    .sensimedical-header {
        background: linear-gradient(90deg, #1e3a5f 0%, #2d5a87 100%);
        padding: 0.6rem 1.5rem;
        margin-left: calc(-50vw + 50%);
        margin-right: calc(-50vw + 50%);
        margin-bottom: 1rem;
        width: 100vw;
        box-sizing: border-box;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .sensimedical-header img { height: 32px; width: auto; display: block; }
    /* Hide sidebar */
    [data-testid="stSidebar"] { display: none; }
    [data-testid="stSidebar"] ~ div { margin-left: 0 !important; }
    /* Main content headers – dark blue */
    h1, h2, h3 { color: #1e3a5f !important; font-weight: 600; }
    /* Primary button – SensiMedical teal accent */
    .stButton > button {
        background: linear-gradient(90deg, #0d9488 0%, #0f766e 100%) !important;
        color: white !important;
        border: none !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }
    .stButton > button:hover {
        background: #0f766e !important;
        color: white !important;
    }
    /* Inputs – light border */
    .stTextInput input, .stDataFrame { border-radius: 6px; }
    /* Expander – light blue tint */
    .streamlit-expanderHeader { background-color: #f0f9ff; color: #1e3a5f; border-radius: 6px; }
    /* Alerts */
    [data-testid="stSuccess"] { border-left: 4px solid #0d9488; background: #f0fdfa; }
    [data-testid="stWarning"] { border-left: 4px solid #2d5a87; background: #f0f9ff; }
    [data-testid="stError"] { border-left: 4px solid #b91c1c; }
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
        # Build lookup: support both 3-part keys (Customer|Date|Cases) and old 4-part keys (with Sales)
        key_to_date = {}
        key_to_comments = {}
        for _, row in overrides.iterrows():
            k = row["order_key"]
            key_to_date[k] = row["follow_up"]
            key_to_comments[k] = str(row.get("comments") or "")
            parts = k.split("|")
            if len(parts) >= 4:
                k_short = "|".join(parts[:3])
                if k_short not in key_to_date:
                    key_to_date[k_short] = row["follow_up"]
                    key_to_comments[k_short] = str(row.get("comments") or "")
            if len(parts) >= 2:
                k_so = parts[0]
                if k_so not in key_to_date:
                    key_to_date[k_so] = row["follow_up"]
                    key_to_comments[k_so] = str(row.get("comments") or "")
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

    st.write("Edit **Scheduled date** and **Comments** below.")

    # Show table without order_key (internal use only); only Scheduled date and Comments are editable
    display_df = df.drop(columns=["order_key"])
    column_config = {
        "Scheduled date": st.column_config.DateColumn("Scheduled date"),
        "Comments": st.column_config.TextColumn("Comments", width="large"),
    }
    for col in display_df.columns:
        if col not in ("Scheduled date", "Comments"):
            column_config[col] = st.column_config.Column(col, disabled=True)
    edited_display = st.data_editor(
        display_df,
        column_config=column_config,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
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


from typing import Tuple
import math
from html import escape
import textwrap

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
import io
from pathlib import Path

# -------------------------------------------------------------------
# Config / constants
# -------------------------------------------------------------------

SPONSORSHIP_SHEET = "Sponsorships"
PUBLIC_INVESTMENT_SHEET = "Public Investment"
CONTACT_DETAIL_SHEET = "Contact Detail"
DATA_DICTIONARY_SHEET = "Data_Dictionary"
KEY_CONVERSATIONS_SHEET = "Key Conversations"
KEY_CONVERSATION_COL = "Brand"

# Optional workbook sheet used to break out the BizX portion of the closed sponsorship goal.
# Recommended sheet name: BizX Details
# Supported column names include Brand, Detail, Amount, and Active.
BIZX_DETAILS_SHEET = "BizX Details"
BIZX_DETAILS_SHEET_ALIASES = ["BizX Details", "BizX Breakdown", "BizX", "BizX Investment"]
BIZX_BRAND_COL = "Brand"
BIZX_DETAIL_COL = "Detail"
BIZX_AMOUNT_COL = "Amount"
BIZX_ACTIVE_COL = "Active"

PARTNER_TYPE_COL = "Partner Type"
CURRENT_INVESTMENT_COL = "Current Proposed Investment"
CONTRACTED_COL = "Contracted Annual Revenue ($)"
PROBABILITY_COL = "Probability (%)"
EXPECTED_COL = "Expected Value ($)"
TERM_COL = "Term (years)"
INTEREST_COL = "Interest"

STAGE_ORDER = ["Lead", "Under 50%", "50-75%", "Over 75%", "Contracted"]
TOTALS_STAGE_ORDER = ["Under 50%", "50-75%", "Over 75%", "Contracted"]

STAGE_COLORS = {
    "Lead": "#64748b",
    "Under 50%": "#38bdf8",
    "50-75%": "#f59e0b",
    "Over 75%": "#a855f7",
    "Contracted": "#22c55e",
    "Dead": "#475569",
}

TYPE_COLORS = {
    "Sponsorship": "#8b5cf6",
    "Public Investment": "#06b6d4",
}

PROPERTY_COLORS = {
    "Bumbershoot": "#8b5cf6",
    "Cannonball Arts": "#06b6d4",
}

# Short display names for the closed-business bubbles in the 2026 Goal card.
# Extend this dictionary if a row name in the workbook should display more cleanly.
CLOSED_BUSINESS_DISPLAY_NAMES = {
    "WW Toyota Dealers": "Toyota",
    "Monster Energy": "Monster",
    "Westland Distillery": "Westland",
}


# -------------------------------------------------------------------
# Styling
# -------------------------------------------------------------------

def apply_custom_css() -> None:
    st.markdown(
        """
        <style>
            :root {
                --bg: #0b1220;
                --panel: #111827;
                --panel-2: #172033;
                --panel-3: #1f2937;
                --border: rgba(148, 163, 184, 0.18);
                --text: #e5e7eb;
                --muted: #94a3b8;
                --good: #22c55e;
                --shadow: 0 16px 40px rgba(2, 8, 23, 0.32);
            }

            .stApp {
                background:
                    radial-gradient(circle at top left, rgba(139,92,246,0.14), transparent 28%),
                    radial-gradient(circle at top right, rgba(6,182,212,0.12), transparent 24%),
                    linear-gradient(180deg, #0b1220 0%, #0f172a 100%);
                color: var(--text);
            }

            .block-container {
                padding-top: 2rem;
                padding-bottom: 2rem;
            }

            .dashboard-section-title {
                font-size: 1.15rem;
                font-weight: 700;
                letter-spacing: 0.01em;
                color: var(--text);
                margin: 0 0 0.65rem 0;
            }

            .dashboard-section-subtitle {
                color: var(--muted);
                font-size: 0.92rem;
                margin: 0 0 1.05rem 0;
            }

            .kpi-card, .goal-card, .deal-card-shell {
                background: linear-gradient(180deg, rgba(23,32,51,.96), rgba(17,24,39,.96));
                border: 1px solid var(--border);
                border-radius: 20px;
                box-shadow: var(--shadow);
            }

            .kpi-card {
                padding: 18px 18px 16px 18px;
                min-height: 118px;
            }

            .kpi-label {
                color: var(--muted);
                font-size: 0.82rem;
                text-transform: uppercase;
                letter-spacing: .08em;
                margin-bottom: .5rem;
            }

            .kpi-value {
                color: var(--text);
                font-size: 2rem;
                font-weight: 700;
                line-height: 1.05;
                margin-bottom: .35rem;
            }

            .kpi-sub {
                color: var(--muted);
                font-size: 0.92rem;
            }

            .goal-card {
                padding: 18px 20px 18px 20px;
                min-height: 178px;
            }

            .goal-header {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 12px;
                margin-bottom: 14px;
            }

            .goal-title {
                font-size: 1rem;
                font-weight: 700;
                color: var(--text);
            }

            .goal-pill {
                display: inline-flex;
                align-items: center;
                padding: 6px 10px;
                border-radius: 999px;
                font-size: 0.78rem;
                background: rgba(148, 163, 184, 0.10);
                color: var(--muted);
                border: 1px solid var(--border);
            }

            .goal-track {
                position: relative;
                width: 100%;
                height: 26px;
                border-radius: 999px;
                background: rgba(148, 163, 184, 0.14);
                overflow: hidden;
                border: 1px solid rgba(148, 163, 184, 0.10);
            }

            .goal-fill {
                position: absolute;
                top: 0;
                left: 0;
                height: 100%;
                border-radius: 999px;
            }

            .goal-fill-standard {
                background: linear-gradient(90deg, #16a34a, #22c55e);
            }

            .goal-fill-bizx {
                background: linear-gradient(90deg, #0d9488, #2dd4bf);
                box-shadow: inset 1px 0 0 rgba(255,255,255,0.18);
            }

            .bizx-goal-note {
                margin-top: 10px;
                display: flex;
                align-items: center;
                gap: 8px;
                color: #99f6e4;
                font-size: 0.88rem;
                font-weight: 650;
            }

            .bizx-legend-dot {
                width: 10px;
                height: 10px;
                border-radius: 999px;
                display: inline-block;
                background: linear-gradient(90deg, #0d9488, #2dd4bf);
                box-shadow: 0 0 0 3px rgba(45, 212, 191, 0.12);
            }

            .bizx-detail-bubbles {
                display: flex;
                align-items: center;
                gap: 7px;
                flex-wrap: wrap;
                margin-top: 8px;
            }

            .bizx-detail-bubble {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 6px 11px;
                border-radius: 999px;
                font-size: 0.75rem;
                font-weight: 750;
                line-height: 1;
                white-space: nowrap;
                background: rgba(13, 148, 136, 0.18);
                color: #99f6e4;
                border: 1px solid rgba(45, 212, 191, 0.38);
                box-shadow: 0 7px 14px rgba(2, 8, 23, 0.16);
            }

            .goal-scale {
                display: flex;
                justify-content: space-between;
                gap: 10px;
                margin-top: 10px;
                color: var(--muted);
                font-size: 0.86rem;
            }

            .goal-main-number {
                margin-top: 14px;
                font-size: 1.85rem;
                line-height: 1.05;
                font-weight: 750;
                color: var(--text);
            }

            .goal-main-row {
                display: flex;
                align-items: center;
                gap: 12px;
                flex-wrap: wrap;
                margin-top: 14px;
            }

            .goal-main-row .goal-main-number {
                margin-top: 0;
            }

            .goal-closed-bubbles {
                display: inline-flex;
                align-items: center;
                gap: 7px;
                flex-wrap: wrap;
            }

            .goal-closed-bubble {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 6px 11px;
                border-radius: 999px;
                font-size: 0.76rem;
                font-weight: 750;
                line-height: 1;
                white-space: nowrap;
                border: 1px solid rgba(255,255,255,0.10);
                box-shadow: 0 7px 14px rgba(2, 8, 23, 0.16);
            }

            .goal-note {
                margin-top: 6px;
                color: var(--muted);
                font-size: 0.92rem;
            }

            .key-conversations-card {
                background: linear-gradient(180deg, rgba(23,32,51,.96), rgba(17,24,39,.96));
                border: 1px solid var(--border);
                border-radius: 20px;
                box-shadow: var(--shadow);
                padding: 18px 20px;
            }

            .key-conversations-row {
                display: flex;
                align-items: center;
                gap: 10px;
                flex-wrap: wrap;
            }

            .key-conversation-bubble {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 9px 16px;
                border-radius: 999px;
                font-size: 0.92rem;
                font-weight: 750;
                line-height: 1;
                white-space: nowrap;
                border: 1px solid rgba(255,255,255,0.10);
                box-shadow: 0 8px 18px rgba(2, 8, 23, 0.18);
            }

            .key-conversations-empty {
                color: var(--muted);
                font-size: 0.92rem;
            }

            .interest-pills {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                flex-wrap: wrap;
            }

            .interest-pill {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 5px 12px;
                border-radius: 999px;
                font-size: 0.76rem;
                font-weight: 700;
                line-height: 1;
                white-space: nowrap;
                border: 1px solid transparent;
            }

            .interest-pill.bumbershoot {
                background: rgba(139, 92, 246, 0.18);
                color: #c4b5fd;
                border-color: rgba(139, 92, 246, 0.35);
            }

            .interest-pill.cannonball {
                background: rgba(6, 182, 212, 0.18);
                color: #67e8f9;
                border-color: rgba(6, 182, 212, 0.35);
            }

            .deal-card-shell {
                padding: 16px 16px 8px 16px;
                min-height: 640px;
            }

            .deal-panel-title {
                font-size: 1rem;
                font-weight: 700;
                color: var(--text);
                margin-bottom: 10px;
            }

            .deal-row {
                padding: 10px 0 12px 0;
                border-bottom: 1px solid rgba(148, 163, 184, 0.10);
            }

            .deal-row:last-child {
                border-bottom: none;
            }

            .deal-header {
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                gap: 12px;
                margin-bottom: 8px;
            }

            .deal-left {
                min-width: 0;
            }

            .deal-rank-name {
                display: flex;
                align-items: center;
                gap: 10px;
                flex-wrap: wrap;
                min-width: 0;
            }

            .deal-rank {
                color: var(--muted);
                font-size: 0.82rem;
                font-weight: 700;
                min-width: 22px;
            }

            .deal-name {
                color: var(--text);
                font-size: 0.98rem;
                font-weight: 650;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                max-width: 350px;
            }

            .stage-pill {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                padding: 5px 10px;
                border-radius: 999px;
                font-size: 0.75rem;
                color: var(--text);
                border: 1px solid rgba(255,255,255,0.08);
                white-space: nowrap;
            }

            .deal-interest-row {
                margin-top: 6px;
                margin-left: 32px;
            }

            .deal-track {
                width: 100%;
                height: 12px;
                border-radius: 999px;
                background: rgba(148, 163, 184, 0.12);
                overflow: hidden;
                margin-bottom: 7px;
            }

            .deal-fill {
                height: 100%;
                border-radius: 999px;
            }

            .deal-meta {
                color: var(--muted);
                font-size: 0.86rem;
            }

            .deal-meta strong {
                color: var(--text);
                font-weight: 650;
            }

            [data-testid="stExpander"] {
                border: 1px solid var(--border);
                border-radius: 18px;
                background: rgba(17,24,39,0.82);
                overflow: hidden;
            }

            [data-testid="stDataFrame"], [data-testid="stTable"] {
                border: 1px solid var(--border);
                border-radius: 14px;
                overflow: hidden;
            }

            .mini-note {
                color: var(--muted);
                font-size: 0.84rem;
                margin-top: -0.15rem;
                margin-bottom: 0.8rem;
            }

            .interest-pills {
                display: inline-flex;
                align-items: center;
                gap: 8px;
                flex-wrap: wrap;
            }

            .interest-pill {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 5px 12px;
                border-radius: 999px;
                font-size: 0.76rem;
                font-weight: 700;
                line-height: 1;
                white-space: nowrap;
                border: 1px solid transparent;
            }

            .deal-interest-row {
                margin-top: 6px;
                margin-left: 32px;
            }
            
        </style>
        """,
        unsafe_allow_html=True,
    )


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------

def format_currency(value: float) -> str:
    return f"${float(value):,.0f}"


def format_percent(value: float) -> str:
    return f"{float(value):,.0f}%"


def nice_ceiling(value: float) -> float:
    value = float(value or 0)
    if value <= 0:
        return 50000
    if value <= 250000:
        step = 25000
    elif value <= 1000000:
        step = 50000
    else:
        step = 100000
    return math.ceil(value / step) * step


def _normalize_partner_type(val: str):
    if pd.isna(val):
        return None
    s = str(val).strip().lower()
    if not s:
        return None
    if s.startswith("sponsor"):
        return "Sponsorship"
    if s.startswith("public"):
        return "Public Investment"
    return None


def _is_flag(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s in ("x", "1", "true", "yes", "y")


def _coalesce_column(df: pd.DataFrame, aliases: list[str], target: str) -> pd.DataFrame:
    for alias in aliases:
        if alias in df.columns and alias != target:
            df = df.rename(columns={alias: target})
            break
    return df


def _normalize_interest_flags(val) -> tuple[bool, bool]:
    if pd.isna(val):
        return False, False

    s = str(val).strip().lower()
    if not s or s in {"none", "no", "n/a", "na"}:
        return False, False

    has_bumbershoot = "bumbershoot" in s
    has_cannonball = "cannonball" in s

    return has_bumbershoot, has_cannonball


def _closed_value_series(df: pd.DataFrame) -> pd.Series:
    contracted = pd.to_numeric(df.get(CONTRACTED_COL, 0), errors="coerce").fillna(0.0)
    proposed = pd.to_numeric(df.get(CURRENT_INVESTMENT_COL, 0), errors="coerce").fillna(0.0)
    return np.where(contracted > 0, contracted, proposed)


def _filter_contacts_to_visible_prospects(
    contacts: pd.DataFrame,
    filtered_prospects: pd.DataFrame,
) -> pd.DataFrame:
    if contacts.empty or filtered_prospects.empty:
        return contacts.iloc[0:0].copy()

    visible_names = set(
        filtered_prospects["Prospect (Account Name)"]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )
    if not visible_names:
        return contacts.iloc[0:0].copy()

    mask = contacts["Prospect (Account Name)"].astype(str).str.strip().isin(visible_names)
    return contacts[mask].copy()


def _extract_key_conversation_brands(raw_df: pd.DataFrame) -> list[str]:
    if raw_df.empty:
        return []

    cols = [str(col).strip() for col in raw_df.columns]
    raw_df = raw_df.copy()
    raw_df.columns = cols

    brand_col = KEY_CONVERSATION_COL if KEY_CONVERSATION_COL in raw_df.columns else raw_df.columns[0]

    brands: list[str] = []
    seen = set()

    for value in raw_df[brand_col].dropna().tolist():
        name = str(value).strip()
        key = name.lower()
        if not name or key in {"none", "n/a", "na"}:
            continue
        if key not in seen:
            brands.append(name)
            seen.add(key)

    return brands


def _first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the first matching column name from a list of likely aliases."""
    normalized = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        match = normalized.get(candidate.strip().lower())
        if match is not None:
            return match
    return None


def _extract_bizx_details(raw_df: pd.DataFrame) -> pd.DataFrame:
    """Normalize the optional BizX Details sheet into brand/detail/amount rows.

    Recommended workbook sheet: `BizX Details`
    Recommended columns: `Brand`, `Detail`, `Amount`, `Active`

    `Active` is optional. If used, mark rows with Yes/True/X/1 to include them.
    `Amount` is optional, but the app needs it to size the BizX bar segment.
    """
    output_cols = [BIZX_BRAND_COL, BIZX_DETAIL_COL, BIZX_AMOUNT_COL]
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=output_cols)

    df = raw_df.copy()
    df.columns = [str(col).strip() for col in df.columns]

    brand_col = _first_existing_column(
        df,
        [
            "Brand",
            "BizX Brand",
            "Name",
            "Line Item",
            "Investment Source",
            "Description",
        ],
    )
    detail_col = _first_existing_column(
        df,
        [
            "Detail",
            "Details",
            "BizX Detail",
            "Description",
            "Notes",
            "Category",
        ],
    )
    amount_col = _first_existing_column(
        df,
        [
            "Amount",
            "Value",
            "BizX Amount",
            "Investment",
            "BizX Investment",
            "Closed Value",
            "Closed Amount",
            "Contracted Annual Revenue ($)",
            "Current Proposed Investment",
        ],
    )
    active_col = _first_existing_column(df, ["Active", "Include", "Included", "Show", "Use"])

    if brand_col is None:
        brand_col = df.columns[0]

    if active_col is not None:
        df = df[df[active_col].apply(_is_flag)].copy()

    rows = []
    for _, row in df.iterrows():
        brand = str(row.get(brand_col, "") or "").strip()
        detail = str(row.get(detail_col, "") or "").strip() if detail_col else ""
        amount = pd.to_numeric(pd.Series([row.get(amount_col, 0) if amount_col else 0]), errors="coerce").fillna(0).iloc[0]

        if not brand or brand.lower() in {"none", "n/a", "na"}:
            continue

        rows.append(
            {
                BIZX_BRAND_COL: brand,
                BIZX_DETAIL_COL: detail,
                BIZX_AMOUNT_COL: float(amount),
            }
        )

    return pd.DataFrame(rows, columns=output_cols)


def _extract_bizx_details_from_closed_rows(closed_df: pd.DataFrame) -> pd.DataFrame:
    """Fallback for workbooks that tag BizX directly in the Sponsorships sheet."""
    output_cols = [BIZX_BRAND_COL, BIZX_DETAIL_COL, BIZX_AMOUNT_COL]
    if closed_df.empty or "Prospect (Account Name)" not in closed_df.columns:
        return pd.DataFrame(columns=output_cols)

    sponsorship_df = closed_df[closed_df[PARTNER_TYPE_COL] == "Sponsorship"].copy()
    if sponsorship_df.empty:
        return pd.DataFrame(columns=output_cols)

    bizx_flag_cols = [
        col
        for col in sponsorship_df.columns
        if str(col).strip().lower() in {
            "bizx",
            "bizx?",
            "is bizx",
            "bizx investment",
            "bizx flag",
        }
    ]

    mask = sponsorship_df["Prospect (Account Name)"].astype(str).str.contains("bizx", case=False, na=False)
    if INTEREST_COL in sponsorship_df.columns:
        mask = mask | sponsorship_df[INTEREST_COL].astype(str).str.contains("bizx", case=False, na=False)

    for col in bizx_flag_cols:
        mask = mask | sponsorship_df[col].apply(_is_flag)

    bizx_rows = sponsorship_df[mask].copy()
    if bizx_rows.empty:
        return pd.DataFrame(columns=output_cols)

    bizx_rows["Closed Value"] = _closed_value_series(bizx_rows)

    rows = []
    for _, row in bizx_rows.iterrows():
        rows.append(
            {
                BIZX_BRAND_COL: _display_closed_business_brand_name(row.get("Prospect (Account Name)", "BizX")),
                BIZX_DETAIL_COL: "",
                BIZX_AMOUNT_COL: float(row.get("Closed Value", 0) or 0),
            }
        )

    return pd.DataFrame(rows, columns=output_cols)


def _resolve_bizx_breakdown(
    closed_df: pd.DataFrame,
    bizx_details: pd.DataFrame,
) -> tuple[float, pd.DataFrame]:
    """Return the BizX closed amount and display rows.

    Priority:
    1. Use the optional BizX Details sheet if it contains dollar amounts.
    2. If that sheet only has labels, pair those labels with the amount from BizX-tagged closed sponsorship rows.
    3. If no detail sheet exists, use BizX-tagged closed sponsorship rows as a fallback.
    """
    detail_df = _extract_bizx_details(bizx_details) if bizx_details is not None else pd.DataFrame(columns=[BIZX_BRAND_COL, BIZX_DETAIL_COL, BIZX_AMOUNT_COL])
    row_df = _extract_bizx_details_from_closed_rows(closed_df)

    detail_amount = float(detail_df[BIZX_AMOUNT_COL].sum()) if not detail_df.empty else 0.0
    row_amount = float(row_df[BIZX_AMOUNT_COL].sum()) if not row_df.empty else 0.0

    if detail_amount > 0:
        return detail_amount, detail_df

    if not detail_df.empty and row_amount > 0:
        return row_amount, detail_df

    if row_amount > 0:
        return row_amount, row_df

    return 0.0, detail_df


def _build_bizx_details_html(bizx_details: pd.DataFrame) -> str:
    if bizx_details is None or bizx_details.empty:
        return ""

    bubbles = []
    seen = set()

    for _, row in bizx_details.iterrows():
        brand = str(row.get(BIZX_BRAND_COL, "") or "").strip()
        detail = str(row.get(BIZX_DETAIL_COL, "") or "").strip()
        amount = float(row.get(BIZX_AMOUNT_COL, 0) or 0)

        label = brand
        if detail and detail.lower() != brand.lower():
            label = f"{brand}: {detail}"
        if amount > 0:
            label = f"{label} · {format_currency(amount)}"

        key = label.lower()
        if not label or key in seen:
            continue

        bubbles.append(f'<span class="bizx-detail-bubble">{escape(label)}</span>')
        seen.add(key)

    if not bubbles:
        return ""

    return f'<div class="bizx-detail-bubbles">{"".join(bubbles)}</div>'


# -------------------------------------------------------------------
# Data loading
# -------------------------------------------------------------------

def load_workbook(xlsx_file) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, list[str], pd.DataFrame]:
    try:
        if hasattr(xlsx_file, "seek"):
            xlsx_file.seek(0)
        xls = pd.ExcelFile(xlsx_file)
        sponsorships = pd.read_excel(xls, sheet_name=SPONSORSHIP_SHEET)
        public = pd.read_excel(xls, sheet_name=PUBLIC_INVESTMENT_SHEET)
        contacts = pd.read_excel(xls, sheet_name=CONTACT_DETAIL_SHEET)
        data_dict = pd.read_excel(xls, sheet_name=DATA_DICTIONARY_SHEET)

        if KEY_CONVERSATIONS_SHEET in xls.sheet_names:
            key_conversation_df = pd.read_excel(xls, sheet_name=KEY_CONVERSATIONS_SHEET)
            key_conversations = _extract_key_conversation_brands(key_conversation_df)
        else:
            key_conversations = []

        bizx_details = pd.DataFrame(columns=[BIZX_BRAND_COL, BIZX_DETAIL_COL, BIZX_AMOUNT_COL])
        for sheet_name in BIZX_DETAILS_SHEET_ALIASES:
            if sheet_name in xls.sheet_names:
                bizx_details = _extract_bizx_details(pd.read_excel(xls, sheet_name=sheet_name))
                break
    except Exception as e:
        st.error("There was a problem reading one or more sheets from the workbook.")
        st.exception(e)
        st.stop()

    # Normalize columns so the app works with both the old and new workbook labels.
    for df in (sponsorships, public):
        df = _coalesce_column(
            df,
            ["Projected Annual Revenue ($)", "Projected Annual Value ($)", "Current Proposed Investment"],
            CURRENT_INVESTMENT_COL,
        )
        df = _coalesce_column(
            df,
            ["Contracted Annual Revenue ($)", "Contracted Annual Value ($)"],
            CONTRACTED_COL,
        )

    sponsorships = _coalesce_column(
        sponsorships,
        ["Projected Annual Revenue ($)", "Projected Annual Value ($)", "Current Proposed Investment"],
        CURRENT_INVESTMENT_COL,
    )
    sponsorships = _coalesce_column(
        sponsorships,
        ["Contracted Annual Revenue ($)", "Contracted Annual Value ($)"],
        CONTRACTED_COL,
    )

    public = _coalesce_column(
        public,
        ["Projected Annual Revenue ($)", "Projected Annual Value ($)", "Current Proposed Investment"],
        CURRENT_INVESTMENT_COL,
    )
    public = _coalesce_column(
        public,
        ["Contracted Annual Revenue ($)", "Contracted Annual Value ($)"],
        CONTRACTED_COL,
    )

    sponsorships[PARTNER_TYPE_COL] = "Sponsorship"
    public[PARTNER_TYPE_COL] = "Public Investment"

    prospects = pd.concat([sponsorships, public], ignore_index=True)

    for col in ["Prospect ID", "Prospect (Account Name)"]:
        if col not in prospects.columns:
            st.error(f"Expected column `{col}` not found in prospect sheets.")
            st.stop()

    prospects = prospects.dropna(
        subset=["Prospect ID", "Prospect (Account Name)"], how="all"
    ).copy()

    numeric_cols = [
        CURRENT_INVESTMENT_COL,
        CONTRACTED_COL,
        PROBABILITY_COL,
        EXPECTED_COL,
        TERM_COL,
    ]
    for col in numeric_cols:
        if col in prospects.columns:
            prospects[col] = pd.to_numeric(prospects[col], errors="coerce").fillna(0.0)
        else:
            prospects[col] = 0.0

    if PROBABILITY_COL in prospects.columns and not prospects[PROBABILITY_COL].isna().all():
        max_prob = prospects[PROBABILITY_COL].max()
        if max_prob <= 1.0:
            prospects[PROBABILITY_COL] = prospects[PROBABILITY_COL] * 100.0

    def _compute_stage_bucket(row: pd.Series) -> str:
        if "Dead" in row and _is_flag(row["Dead"]):
            return "Dead"

        contracted_rev = row.get(CONTRACTED_COL, 0.0)
        try:
            contracted_rev = float(contracted_rev)
        except (TypeError, ValueError):
            contracted_rev = 0.0

        if ("Contracted" in row and _is_flag(row["Contracted"])) or contracted_rev > 0:
            return "Contracted"

        p_raw = row.get(PROBABILITY_COL)
        try:
            p = float(p_raw)
        except (TypeError, ValueError):
            p = float("nan")

        if not math.isnan(p):
            if p == 0:
                return "Lead"
            if p < 50:
                return "Under 50%"
            if p < 75:
                return "50-75%"
            if p < 100:
                return "Over 75%"
            return "Contracted"

        stage_cols_map = {
            "Lead": "Lead",
            "Prospect": "Lead",
            "Under 50%": "Under 50%",
            "50-75%": "50-75%",
            "Over 75%": "Over 75%",
        }
        for col, bucket in stage_cols_map.items():
            if col in row and _is_flag(row[col]):
                return bucket

        return "Lead"

    prospects["Stage Bucket"] = prospects.apply(_compute_stage_bucket, axis=1)

    if INTEREST_COL not in prospects.columns:
        prospects[INTEREST_COL] = ""

    interest_flags = prospects[INTEREST_COL].apply(_normalize_interest_flags)
    prospects["Has Bumbershoot Interest"] = interest_flags.apply(lambda x: x[0])
    prospects["Has Cannonball Interest"] = interest_flags.apply(lambda x: x[1])

    if "Prospect (Account Name)" not in contacts.columns:
        st.error("Expected column `Prospect (Account Name)` not found in Contact Detail.")
        st.stop()

    contacts = contacts.dropna(
        subset=["Prospect (Account Name)", "Contact Date"], how="all"
    ).copy()

    for col in ["Contact Date", "Follow-up Date"]:
        if col in contacts.columns:
            contacts[col] = pd.to_datetime(contacts[col], errors="coerce")

    return prospects, contacts, data_dict, key_conversations, bizx_details


def _display_closed_business_brand_name(name: str) -> str:
    clean_name = str(name or "").strip()
    return CLOSED_BUSINESS_DISPLAY_NAMES.get(clean_name, clean_name)


def _extract_closed_business_brands(closed_df: pd.DataFrame, partner_type: str) -> list[str]:
    if closed_df.empty or PARTNER_TYPE_COL not in closed_df.columns:
        return []

    sub = closed_df[closed_df[PARTNER_TYPE_COL] == partner_type].copy()
    if sub.empty or "Prospect (Account Name)" not in sub.columns:
        return []

    brands: list[str] = []
    seen = set()

    for value in sub["Prospect (Account Name)"].dropna().tolist():
        display_name = _display_closed_business_brand_name(value)
        key = display_name.lower()
        if not display_name or key in seen:
            continue
        brands.append(display_name)
        seen.add(key)

    return brands


def _build_goal_closed_bubbles_html(brands: list[str]) -> str:
    if not brands:
        return ""

    bubble_styles = [
        {"bg": "rgba(139, 92, 246, 0.22)", "text": "#ddd6fe", "border": "rgba(139, 92, 246, 0.45)"},
        {"bg": "rgba(6, 182, 212, 0.20)", "text": "#a5f3fc", "border": "rgba(6, 182, 212, 0.42)"},
        {"bg": "rgba(34, 197, 94, 0.18)", "text": "#bbf7d0", "border": "rgba(34, 197, 94, 0.38)"},
        {"bg": "rgba(249, 115, 22, 0.18)", "text": "#fed7aa", "border": "rgba(249, 115, 22, 0.38)"},
        {"bg": "rgba(236, 72, 153, 0.18)", "text": "#fbcfe8", "border": "rgba(236, 72, 153, 0.38)"},
    ]

    bubbles = []
    for i, brand in enumerate(brands):
        style = bubble_styles[i % len(bubble_styles)]
        bubbles.append(
            f'<span class="goal-closed-bubble" '
            f'style="background:{style["bg"]}; color:{style["text"]}; border-color:{style["border"]};">'
            f'{escape(str(brand))}'
            f'</span>'
        )

    return f'<div class="goal-closed-bubbles">{"".join(bubbles)}</div>'


# -------------------------------------------------------------------
# Section builders
# -------------------------------------------------------------------

def build_goal_section(prospects: pd.DataFrame, bizx_details: pd.DataFrame | None = None) -> None:
    st.markdown('<div class="dashboard-section-title">2026 Goal</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="dashboard-section-subtitle">Closed dollars at a glance for Sponsorship and Public Investment.</div>',
        unsafe_allow_html=True,
    )

    active_df = prospects[prospects["Stage Bucket"] != "Dead"].copy()
    closed_df = active_df[active_df["Stage Bucket"] == "Contracted"].copy()
    if closed_df.empty:
        closed_df = active_df.iloc[0:0].copy()

    closed_df["Closed Value"] = _closed_value_series(closed_df)

    # Cash sponsorship closed dollars from the main Sponsorships sheet.
    # BizX is treated as an ADDITIONAL closed-value layer from the optional BizX Details sheet,
    # not as a subset to subtract from the cash sponsorship total.
    cash_sponsorship_closed = float(
        closed_df.loc[closed_df[PARTNER_TYPE_COL] == "Sponsorship", "Closed Value"].sum() or 0.0
    )

    public_closed = float(
        closed_df.loc[closed_df[PARTNER_TYPE_COL] == "Public Investment", "Closed Value"].sum() or 0.0
    )

    sponsorship_closed_brands = _extract_closed_business_brands(closed_df, "Sponsorship")
    sponsorship_closed_bubbles_html = _build_goal_closed_bubbles_html(sponsorship_closed_brands)

    sponsorship_goal = 1_000_000
    public_scale = nice_ceiling(max(public_closed * 1.15, 50000))

    # Pull BizX detail rows from the BizX Details sheet when present.
    # Example: $239K cash sponsorship + $190K BizX = $429K shown against the $1M goal.
    raw_bizx_closed, bizx_breakdown_df = _resolve_bizx_breakdown(closed_df, bizx_details)
    bizx_closed = max(float(raw_bizx_closed or 0.0), 0.0)

    total_sponsorship_goal_value = cash_sponsorship_closed + bizx_closed

    standard_sponsorship_pct = min((cash_sponsorship_closed / sponsorship_goal) * 100, 100) if sponsorship_goal else 0
    bizx_pct = min((bizx_closed / sponsorship_goal) * 100, max(100 - standard_sponsorship_pct, 0)) if sponsorship_goal else 0
    sponsorship_pct = min((total_sponsorship_goal_value / sponsorship_goal) * 100, 100) if sponsorship_goal else 0
    public_pct = min((public_closed / public_scale) * 100, 100) if public_scale else 0

    bizx_details_html = _build_bizx_details_html(bizx_breakdown_df)
    bizx_summary_html = ""
    if bizx_closed > 0:
        bizx_summary_html = (
            f'<div class="bizx-goal-note">'
            f'<span class="bizx-legend-dot"></span>'
            f'<span>BizX portion: {format_currency(bizx_closed)}</span>'
            f'</div>'
            f'{bizx_details_html}'
        )

    left, right = st.columns(2)

    with left:
        st.markdown(
            f"""
            <div class="goal-card">
                <div class="goal-header">
                    <div class="goal-title">Sponsorship closed vs. $1M goal</div>
                    <div class="goal-pill">{sponsorship_pct:.0f}% of goal</div>
                </div>
                <div class="goal-track">
                    <div class="goal-fill goal-fill-standard" style="width:{standard_sponsorship_pct:.2f}%;"></div>
                    <div class="goal-fill goal-fill-bizx" style="left:{standard_sponsorship_pct:.2f}%; width:{bizx_pct:.2f}%;"></div>
                </div>
                <div class="goal-scale">
                    <span>{format_currency(0)}</span>
                    <span>{format_currency(sponsorship_goal)}</span>
                </div>
                <div class="goal-main-row">
                    <div class="goal-main-number">{format_currency(total_sponsorship_goal_value)}</div>
                    {sponsorship_closed_bubbles_html}
                </div>
                {bizx_summary_html}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown(
            f"""
            <div class="goal-card">
                <div class="goal-header">
                    <div class="goal-title">Public Investment Value</div>
                    <div class="goal-pill">No cap - bar grows with new closed deals</div>
                </div>
                <div class="goal-track">
                    <div class="goal-fill" style="width:{public_pct:.2f}%; background:linear-gradient(90deg, #0891b2, #06b6d4);"></div>
                </div>
                <div class="goal-scale">
                    <span>{format_currency(0)}</span>
                    <span>{format_currency(public_scale)}</span>
                </div>
                <div class="goal-main-number">{format_currency(public_closed)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

def build_key_conversations_section(key_conversations: list[str]) -> None:
    st.markdown('<div class="dashboard-section-title">Key Conversations</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="dashboard-section-subtitle">Weekly priority brands pulled directly from the Key Conversations sheet in your workbook.</div>',
        unsafe_allow_html=True,
    )

    if not key_conversations:
        st.markdown(
            """
            <div class="key-conversations-card">
                <div class="key-conversations-empty">No key conversations added yet. Add brand names to the <strong>Key Conversations</strong> sheet in the workbook.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    bubble_styles = [
        {"bg": "rgba(139, 92, 246, 0.22)", "text": "#ddd6fe", "border": "rgba(139, 92, 246, 0.45)"},
        {"bg": "rgba(6, 182, 212, 0.20)", "text": "#a5f3fc", "border": "rgba(6, 182, 212, 0.42)"},
        {"bg": "rgba(34, 197, 94, 0.18)", "text": "#bbf7d0", "border": "rgba(34, 197, 94, 0.38)"},
        {"bg": "rgba(249, 115, 22, 0.18)", "text": "#fed7aa", "border": "rgba(249, 115, 22, 0.38)"},
        {"bg": "rgba(236, 72, 153, 0.18)", "text": "#fbcfe8", "border": "rgba(236, 72, 153, 0.38)"},
    ]

    bubbles = []
    for i, brand in enumerate(key_conversations):
        style = bubble_styles[i % len(bubble_styles)]
        bubbles.append(
            f'<span class="key-conversation-bubble" '
            f'style="background:{style["bg"]}; color:{style["text"]}; border-color:{style["border"]};">'
            f'{escape(str(brand))}'
            f'</span>'
        )

    st.markdown(
        f"""
        <div class="key-conversations-card">
            <div class="key-conversations-row">{''.join(bubbles)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_snapshot_section(prospects: pd.DataFrame) -> None:
    st.markdown('<div class="dashboard-section-title">Total Pipeline Snapshot</div>', unsafe_allow_html=True)

    active_df = prospects[prospects["Stage Bucket"] != "Dead"].copy()

    total_expected = active_df[EXPECTED_COL].sum()
    total_current = active_df[CURRENT_INVESTMENT_COL].sum()
    total_contracted = active_df[CONTRACTED_COL].sum()
    deal_count = active_df["Prospect ID"].nunique()

    cards = [
        ("Current Proposed Investment", format_currency(total_current), "Latest proposed value across active opportunities"),
        ("Expected Value", format_currency(total_expected), "Probability weighted pipeline value"),
        ("Contracted Value", format_currency(total_contracted), "Booked dollars already under contract"),
        ("Active Prospects", f"{int(deal_count)}", "Unique accounts excluding deals marked dead"),
    ]

    cols = st.columns(4)
    for col, (label, value, sub) in zip(cols, cards):
        with col:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-label">{escape(label)}</div>
                    <div class="kpi-value">{escape(value)}</div>
                    <div class="kpi-sub">{escape(sub)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def build_pipeline_totals(prospects: pd.DataFrame) -> None:
    st.markdown('<div class="dashboard-section-title">Total Pipeline Value by Stage</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="dashboard-section-subtitle">Expected value only. Lead deals are intentionally excluded for a cleaner pipeline view.</div>',
        unsafe_allow_html=True,
    )

    df = prospects[
        (prospects["Stage Bucket"] != "Dead")
        & (prospects["Stage Bucket"].isin(TOTALS_STAGE_ORDER))
    ].copy()

    if df.empty:
        st.caption("No active pipeline data for the selected filters.")
        return

    agg = (
        df.groupby(["Stage Bucket", PARTNER_TYPE_COL], as_index=False)
        .agg(
            expected_total=(EXPECTED_COL, "sum"),
            deal_count=("Prospect ID", "nunique"),
        )
    )

    chart = (
        alt.Chart(agg)
        .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6, size=30)
        .encode(
            y=alt.Y("Stage Bucket:N", sort=TOTALS_STAGE_ORDER, title=None),
            x=alt.X("expected_total:Q", title="Expected Value ($)"),
            color=alt.Color(
                f"{PARTNER_TYPE_COL}:N",
                scale=alt.Scale(
                    domain=list(TYPE_COLORS.keys()),
                    range=list(TYPE_COLORS.values()),
                ),
                legend=alt.Legend(title="Type"),
            ),
            tooltip=[
                alt.Tooltip("Stage Bucket:N", title="Stage"),
                alt.Tooltip(f"{PARTNER_TYPE_COL}:N", title="Type"),
                alt.Tooltip("expected_total:Q", title="Expected Value", format="$,.0f"),
                alt.Tooltip("deal_count:Q", title="Deals", format=",.0f"),
            ],
        )
        .properties(height=250)
    )

    st.altair_chart(chart, use_container_width=True)


def _build_interest_pills_html(row: pd.Series) -> str:
    raw = str(row.get(INTEREST_COL, "") or "").strip()
    if not raw:
        return ""

    items = [item.strip() for item in raw.split(";") if item.strip()]
    if not items:
        return ""

    INTEREST_STYLES = {
        "Bumbershoot": {
            "bg": "rgba(139, 92, 246, 0.18)",
            "text": "#c4b5fd",
            "border": "rgba(139, 92, 246, 0.35)",
        },
        "Cannonball Arts": {
            "bg": "rgba(6, 182, 212, 0.18)",
            "text": "#67e8f9",
            "border": "rgba(6, 182, 212, 0.35)",
        },
        "Level1": {
            "bg": "rgba(34, 197, 94, 0.18)",
            "text": "#86efac",
            "border": "rgba(34, 197, 94, 0.35)",
        },
        "Art Install": {
            "bg": "rgba(249, 115, 22, 0.18)",
            "text": "#fdba74",
            "border": "rgba(249, 115, 22, 0.35)",
        },
        "IP": {
            "bg": "rgba(236, 72, 153, 0.18)",
            "text": "#f9a8d4",
            "border": "rgba(236, 72, 153, 0.35)",
        },
        "Fashion District": {
            "bg": "rgba(168, 85, 247, 0.18)",
            "text": "#d8b4fe",
            "border": "rgba(168, 85, 247, 0.35)",
        },
        "Recess": {
            "bg": "rgba(234, 179, 8, 0.18)",
            "text": "#fde68a",
            "border": "rgba(234, 179, 8, 0.35)",
        },
        "Entitlement": {
            "bg": "rgba(239, 68, 68, 0.18)",
            "text": "#fca5a5",
            "border": "rgba(239, 68, 68, 0.35)",
        },
        "Distribution": {
            "bg": "rgba(20, 184, 166, 0.18)",
            "text": "#99f6e4",
            "border": "rgba(20, 184, 166, 0.35)",
        },
    }

    default_style = {
        "bg": "rgba(148, 163, 184, 0.16)",
        "text": "#cbd5e1",
        "border": "rgba(148, 163, 184, 0.28)",
    }

    pills = []
    for item in items:
        style = INTEREST_STYLES.get(item, default_style)
        pills.append(
            f'<span class="interest-pill" '
            f'style="background:{style["bg"]}; color:{style["text"]}; border-color:{style["border"]};">'
            f'{escape(item)}'
            f'</span>'
        )

    return f'<div class="deal-interest-row"><div class="interest-pills">{"".join(pills)}</div></div>'

def _render_deal_panel(data: pd.DataFrame, title: str) -> None:
    if data.empty:
        st.markdown(
            f'<div class="deal-card-shell"><div class="deal-panel-title">{escape(title)}</div><div class="mini-note">No deals in this view yet.</div></div>',
            unsafe_allow_html=True,
        )
        return

    max_expected = max(float(data[EXPECTED_COL].max()), 1.0)

    rows_html = []
    for i, row in data.reset_index(drop=True).iterrows():
        stage = row.get("Stage Bucket", "Lead")
        stage_color = STAGE_COLORS.get(stage, "#64748b")
        name = escape(str(row.get("Prospect (Account Name)", "Untitled Prospect")))
        current_value = format_currency(row.get(CURRENT_INVESTMENT_COL, 0))
        expected_value = format_currency(row.get(EXPECTED_COL, 0))
        raw_expected = pd.to_numeric(pd.Series([row.get(EXPECTED_COL, 0)]), errors="coerce").fillna(0).iloc[0]
        width_pct = max((float(raw_expected) / max_expected) * 100, 2 if float(raw_expected) > 0 else 0)
        interest_html = _build_interest_pills_html(row)

        row_html = (
            f'<div class="deal-row">'
            f'<div class="deal-header">'
            f'<div class="deal-left">'
            f'<div class="deal-rank-name">'
            f'<span class="deal-rank">#{i + 1}</span>'
            f'<span class="deal-name">{name}</span>'
            f'</div>'
            f'{interest_html}'
            f'</div>'
            f'<span class="stage-pill" style="background:{stage_color}22; color:{stage_color};">{escape(str(stage))}</span>'
            f'</div>'
            f'<div class="deal-track">'
            f'<div class="deal-fill" style="width:{width_pct:.2f}%; background:{stage_color};"></div>'
            f'</div>'
            f'<div class="deal-meta"><strong>Current Proposed Investment</strong> {escape(current_value)} &nbsp;&nbsp;&bull;&nbsp;&nbsp; <strong>Expected Value</strong> {escape(expected_value)}</div>'
            f'</div>'
        )
        rows_html.append(row_html)

    panel_html = (
        f'<div class="deal-card-shell"><div class="deal-panel-title">{escape(title)}</div>'
        + "".join(rows_html)
        + "</div>"
    )
    st.markdown(panel_html, unsafe_allow_html=True)


def build_top_deals(prospects: pd.DataFrame) -> None:
    st.markdown('<div class="dashboard-section-title">Top 15 Deals by Expected Value</div>', unsafe_allow_html=True)

    def _top_n(df: pd.DataFrame, partner_type: str, n: int = 15) -> pd.DataFrame:
        sub = df[df[PARTNER_TYPE_COL] == partner_type].copy()
        if sub.empty:
            return sub
        return sub.sort_values(EXPECTED_COL, ascending=False).head(n)

    left, right = st.columns(2)

    with left:
        _render_deal_panel(_top_n(prospects, "Sponsorship"), "Top Sponsorship Deals")

    with right:
        _render_deal_panel(_top_n(prospects, "Public Investment"), "Top Public Investment Deals")


def build_pipeline_board(prospects: pd.DataFrame) -> None:
    with st.expander("Pipeline Stage", expanded=False):
        st.markdown(
            '<div class="dashboard-section-subtitle">Detailed stage-by-stage deal list for teams that want the full board view.</div>',
            unsafe_allow_html=True,
        )

        df = prospects[prospects["Stage Bucket"] != "Dead"].copy()
        if df.empty:
            st.caption("No active prospects in the pipeline.")
            return

        partner_types = sorted(df[PARTNER_TYPE_COL].dropna().unique().tolist())
        tab_labels = ["All"] + partner_types
        tabs = st.tabs(tab_labels)

        for label, tab in zip(tab_labels, tabs):
            with tab:
                subset = df.copy() if label == "All" else df[df[PARTNER_TYPE_COL] == label].copy()

                if subset.empty:
                    st.caption("No deals in this segment yet.")
                    continue

                cols = st.columns(len(STAGE_ORDER))

                for col, stage in zip(cols, STAGE_ORDER):
                    with col:
                        stage_df = subset[subset["Stage Bucket"] == stage].copy()

                        cols_to_show = [
                            "Prospect (Account Name)",
                            "Owner",
                            CURRENT_INVESTMENT_COL,
                            CONTRACTED_COL,
                            EXPECTED_COL,
                        ]
                        cols_to_show = [c for c in cols_to_show if c in stage_df.columns]

                        stage_df = stage_df[cols_to_show].sort_values(
                            EXPECTED_COL, ascending=False
                        )

                        st.markdown(
                            f"**{stage}**  \n"
                            f"<span style='font-size:12px;color:#94a3b8;'>{len(stage_df)} deals</span>",
                            unsafe_allow_html=True,
                        )

                        if stage_df.empty:
                            st.caption("No deals at this stage.")
                        else:
                            st.dataframe(
                                stage_df.style.format(
                                    {
                                        CURRENT_INVESTMENT_COL: "${:,.0f}",
                                        CONTRACTED_COL: "${:,.0f}",
                                        EXPECTED_COL: "${:,.0f}",
                                    }
                                ),
                                hide_index=True,
                                use_container_width=True,
                            )


# -------------------------------------------------------------------
# Main app
# -------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title="Bumbershoot / Cannonball - Partnership Revenue Pipeline",
        page_icon="🎟️",
        layout="wide",
    )

    apply_custom_css()

    st.title("Bumbershoot & Cannonball - Partnership Revenue Pipeline")
    st.caption(
        "Upload the latest Excel workbook for a split view of "
        "**Sponsorship** and **Public Investment**."
    )

    # Server-side fallback file
    LAST_UPLOAD_PATH = Path("data/last_uploaded_pipeline.xlsx")
    LAST_UPLOAD_PATH.parent.mkdir(parents=True, exist_ok=True)

    # Session state setup
    if "workbook_bytes" not in st.session_state:
        st.session_state.workbook_bytes = None
    if "workbook_name" not in st.session_state:
        st.session_state.workbook_name = None

    st.sidebar.header("Data source")
    uploaded = st.sidebar.file_uploader(
        "Upload latest Excel workbook (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        key="pipeline_workbook_upload",
    )

    # If user uploads a new workbook, save it to session + disk
    if uploaded is not None:
        file_bytes = uploaded.getvalue()
        st.session_state.workbook_bytes = file_bytes
        st.session_state.workbook_name = uploaded.name
        LAST_UPLOAD_PATH.write_bytes(file_bytes)

    # If nothing is in session, try loading the most recent saved workbook from disk
    if st.session_state.workbook_bytes is None and LAST_UPLOAD_PATH.exists():
        st.session_state.workbook_bytes = LAST_UPLOAD_PATH.read_bytes()
        st.session_state.workbook_name = LAST_UPLOAD_PATH.name

    # Stop only if we truly have nothing available
    if st.session_state.workbook_bytes is None:
        st.info(
            "⬅️ Please upload the latest Excel workbook to see the dashboard.\n\n"
            "Recommended: use the file you update weekly before your team meeting."
        )
        st.stop()

    st.sidebar.caption(f"Using saved workbook: **{st.session_state.workbook_name}**")

    if st.sidebar.button("Clear saved workbook"):
        st.session_state.workbook_bytes = None
        st.session_state.workbook_name = None
        if LAST_UPLOAD_PATH.exists():
            LAST_UPLOAD_PATH.unlink()
        st.rerun()

    # Load from saved bytes instead of directly from uploaded
    prospects, _, _, key_conversations, bizx_details = load_workbook(
        io.BytesIO(st.session_state.workbook_bytes)
    )

    st.sidebar.header("Filters")

    if PARTNER_TYPE_COL not in prospects.columns:
        st.error(f"Expected column `{PARTNER_TYPE_COL}` not found in prospect data.")
        st.stop()

    partner_types = sorted(prospects[PARTNER_TYPE_COL].dropna().unique().tolist())
    selected_partner_types = st.sidebar.multiselect(
        "Partner Type",
        options=partner_types,
        default=partner_types,
    )

    if "Owner" in prospects.columns:
        owner_options = sorted(prospects["Owner"].dropna().unique().tolist())
        selected_owners = st.sidebar.multiselect(
            "Owner (AE)",
            options=owner_options,
            default=owner_options,
        )
        mask_owner = prospects["Owner"].isin(selected_owners)
    else:
        mask_owner = True

    filtered_prospects = prospects[
        prospects[PARTNER_TYPE_COL].isin(selected_partner_types) & mask_owner
    ].copy()

    if filtered_prospects.empty:
        st.warning("No deals match the selected filters.")
        st.stop()

    build_goal_section(filtered_prospects, bizx_details)
    st.divider()

    build_key_conversations_section(key_conversations)
    st.divider()

    build_snapshot_section(filtered_prospects)
    st.divider()

    build_pipeline_totals(filtered_prospects)
    st.divider()

    build_top_deals(filtered_prospects)
    st.divider()

    build_pipeline_board(filtered_prospects)


if __name__ == "__main__":
    main()

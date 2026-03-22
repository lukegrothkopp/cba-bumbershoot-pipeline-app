
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

            .goal-note {
                margin-top: 6px;
                color: var(--muted);
                font-size: 0.92rem;
            }

            .legend-row {
                display: flex;
                flex-wrap: wrap;
                gap: 18px;
                align-items: center;
                margin: 0 0 10px 0;
                color: var(--muted);
                font-size: 0.9rem;
            }

            .legend-item {
                display: inline-flex;
                align-items: center;
                gap: 8px;
            }

            .legend-dot {
                width: 10px;
                height: 10px;
                border-radius: 999px;
                display: inline-block;
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

            .property-dots {
                display: inline-flex;
                align-items: center;
                gap: 6px;
            }

            .property-dot {
                width: 10px;
                height: 10px;
                border-radius: 999px;
                display: inline-block;
                border: 1px solid rgba(255,255,255,0.14);
            }

            .property-dot.inactive {
                background: rgba(148, 163, 184, 0.16) !important;
                border-color: rgba(148, 163, 184, 0.12);
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


# -------------------------------------------------------------------
# Data loading
# -------------------------------------------------------------------

def load_workbook(xlsx_file) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    try:
        sponsorships = pd.read_excel(xlsx_file, sheet_name=SPONSORSHIP_SHEET)
        public = pd.read_excel(xlsx_file, sheet_name=PUBLIC_INVESTMENT_SHEET)
        contacts = pd.read_excel(xlsx_file, sheet_name=CONTACT_DETAIL_SHEET)
        data_dict = pd.read_excel(xlsx_file, sheet_name=DATA_DICTIONARY_SHEET)
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

    return prospects, contacts, data_dict


# -------------------------------------------------------------------
# Section builders
# -------------------------------------------------------------------

def build_goal_section(prospects: pd.DataFrame) -> None:
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

    sponsorship_closed = closed_df.loc[
        closed_df[PARTNER_TYPE_COL] == "Sponsorship", "Closed Value"
    ].sum()

    public_closed = closed_df.loc[
        closed_df[PARTNER_TYPE_COL] == "Public Investment", "Closed Value"
    ].sum()

    sponsorship_goal = 1_000_000
    public_scale = nice_ceiling(max(public_closed * 1.15, 50000))

    sponsorship_pct = min((sponsorship_closed / sponsorship_goal) * 100, 100) if sponsorship_goal else 0
    public_pct = min((public_closed / public_scale) * 100, 100) if public_scale else 0

    left, right = st.columns(2)

    with left:
        st.markdown(
            f"""
            <div class="goal-card">
                <div class="goal-header">
                    <div class="goal-title">Sponsorship closed vs. $1M goal</div>
                    <div class="goal-pill">{(sponsorship_closed / sponsorship_goal) * 100:.0f}% of goal</div>
                </div>
                <div class="goal-track">
                    <div class="goal-fill" style="width:{sponsorship_pct:.2f}%; background:linear-gradient(90deg, #16a34a, #22c55e);"></div>
                </div>
                <div class="goal-scale">
                    <span>{format_currency(0)}</span>
                    <span>{format_currency(sponsorship_goal)}</span>
                </div>
                <div class="goal-main-number">{format_currency(sponsorship_closed)}</div>
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


def _property_dots_html(row: pd.Series) -> str:
    b_class = "property-dot" if bool(row.get("Has Bumbershoot Interest", False)) else "property-dot inactive"
    c_class = "property-dot" if bool(row.get("Has Cannonball Interest", False)) else "property-dot inactive"
    return (
        f'<span class="property-dots">'
        f'<span class="{b_class}" style="background:{PROPERTY_COLORS["Bumbershoot"]};"></span>'
        f'<span class="{c_class}" style="background:{PROPERTY_COLORS["Cannonball Arts"]};"></span>'
        f'</span>'
    )


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

        row_html = textwrap.dedent(
            f"""
            <div class="deal-row">
                <div class="deal-header">
                    <div class="deal-left">
                        <div class="deal-rank-name">
                            <span class="deal-rank">#{i + 1}</span>
                            <span class="deal-name">{name}</span>
                            {_property_dots_html(row)}
                        </div>
                    </div>
                    <span class="stage-pill" style="background:{stage_color}22; color:{stage_color};">{escape(str(stage))}</span>
                </div>
                <div class="deal-track">
                    <div class="deal-fill" style="width:{width_pct:.2f}%; background:{stage_color};"></div>
                </div>
                <div class="deal-meta"><strong>Current Proposed Investment</strong> {escape(current_value)} &nbsp;&nbsp;&bull;&nbsp;&nbsp; <strong>Expected Value</strong> {escape(expected_value)}</div>
            </div>
            """
        ).strip()
        rows_html.append(row_html)

    panel_html = (
        f'<div class="deal-card-shell"><div class="deal-panel-title">{escape(title)}</div>'
        + "".join(rows_html)
        + "</div>"
    )
    st.markdown(panel_html, unsafe_allow_html=True)


def build_top_deals(prospects: pd.DataFrame) -> None:
    st.markdown('<div class="dashboard-section-title">Top 15 Deals by Expected Value</div>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="legend-row">
            <span class="legend-item"><span class="legend-dot" style="background:{PROPERTY_COLORS['Bumbershoot']};"></span>Bumbershoot selected</span>
            <span class="legend-item"><span class="legend-dot" style="background:{PROPERTY_COLORS['Cannonball Arts']};"></span>Cannonball Arts selected</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

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
    prospects, _, _ = load_workbook(
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

    build_goal_section(filtered_prospects)
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


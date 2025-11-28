from typing import Tuple
import math

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

# -------------------------------------------------------------------
# Config / constants
# -------------------------------------------------------------------

SPONSORSHIP_SHEET = "Sponsorships"
PUBLIC_INVESTMENT_SHEET = "Public Investment"
CONTACT_DETAIL_SHEET = "Contact Detail"
DATA_DICTIONARY_SHEET = "Data_Dictionary"

PARTNER_TYPE_COL = "Partner Type"

STAGE_ORDER = ["Lead", "Under 50%", "50–75%", "Over 75%", "Contracted"]


# -------------------------------------------------------------------
# Helper functions
# -------------------------------------------------------------------

def _normalize_partner_type(val: str):
    """Map free-text values from Contact Detail to normalized partner types."""
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


def load_workbook(xlsx_file) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Load Sponsorships + Public Investment + Contact Detail + Data_Dictionary
    from the uploaded Excel workbook, combine to a unified prospects table,
    and compute Stage Bucket.
    """
    try:
        sponsorships = pd.read_excel(xlsx_file, sheet_name=SPONSORSHIP_SHEET)
        public = pd.read_excel(xlsx_file, sheet_name=PUBLIC_INVESTMENT_SHEET)
        contacts = pd.read_excel(xlsx_file, sheet_name=CONTACT_DETAIL_SHEET)
        data_dict = pd.read_excel(xlsx_file, sheet_name=DATA_DICTIONARY_SHEET)
    except Exception as e:
        st.error("There was a problem reading one or more sheets from the workbook.")
        st.exception(e)
        st.stop()

    # Tag each with Partner Type
    sponsorships[PARTNER_TYPE_COL] = "Sponsorship"
    public[PARTNER_TYPE_COL] = "Public Investment"

    prospects = pd.concat([sponsorships, public], ignore_index=True)

    # Drop completely empty prospect rows
    for col in ["Prospect ID", "Prospect (Account Name)"]:
        if col not in prospects.columns:
            st.error(f"Expected column `{col}` not found in prospects sheets.")
            st.stop()

    prospects = prospects.dropna(
        subset=["Prospect ID", "Prospect (Account Name)"], how="all"
    ).copy()

    # Coerce numeric columns
    numeric_cols = [
        "Projected Annual Revenue ($)",
        "Contracted Annual Revenue ($)",
        "Probability (%)",
        "Expected Value ($)",
        "Term (years)",
    ]
    for col in numeric_cols:
        if col in prospects.columns:
            prospects[col] = pd.to_numeric(prospects[col], errors="coerce").fillna(0.0)

    # Probability (%) currently stored as 0.0–0.75; convert to 0–100 scale if needed
    if "Probability (%)" in prospects.columns and not prospects["Probability (%)"].isna().all():
        max_prob = prospects["Probability (%)"].max()
        if max_prob <= 1.0:
            prospects["Probability (%)"] = prospects["Probability (%)"] * 100.0

    # Helpers for stage derivation
    def _is_flag(val) -> bool:
        if pd.isna(val):
            return False
        s = str(val).strip().lower()
        return s in ("x", "1", "true", "yes", "y")

    def _compute_stage_bucket(row: pd.Series) -> str:
        """
        Derive a single Stage Bucket using:
        - Dead / Contracted flags + contracted revenue
        - Probability (%)
        - Stage flag columns (Lead, Prospect, Under 50%, 50-75%, Over 75%)
        """

        # 1) Dead overrides everything
        if "Dead" in row and _is_flag(row["Dead"]):
            return "Dead"

        # 2) Contracted by flag or revenue
        contracted_rev = row.get("Contracted Annual Revenue ($)", 0.0)
        try:
            contracted_rev = float(contracted_rev)
        except (TypeError, ValueError):
            contracted_rev = 0.0

        if ("Contracted" in row and _is_flag(row["Contracted"])) or contracted_rev > 0:
            return "Contracted"

        # 3) Probability-based logic (0–100 scale)
        p_raw = row.get("Probability (%)")
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
                return "50–75%"
            if p < 100:
                return "Over 75%"
            return "Contracted"

        # 4) Fallback to stage flags if no probability
        stage_cols_map = {
            "Lead": "Lead",
            "Prospect": "Lead",      # treat Prospect as Lead bucket for roll-ups
            "Under 50%": "Under 50%",
            "50-75%": "50–75%",
            "Over 75%": "Over 75%",
        }
        for col, bucket in stage_cols_map.items():
            if col in row and _is_flag(row[col]):
                return bucket

        # Default
        return "Lead"

    prospects["Stage Bucket"] = prospects.apply(_compute_stage_bucket, axis=1)

    # --- Contacts cleaning -------------------------------------------
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
# Visualization blocks
# -------------------------------------------------------------------

def build_top_deals(prospects: pd.DataFrame) -> None:
    """2) Top 3 Sponsorship & Top 3 Public Investment deals by Expected Value ($)."""
    st.markdown("### Top Deals by Expected Value")

    def _top_n(df: pd.DataFrame, partner_type: str, n: int = 3) -> pd.DataFrame:
        sub = df[df[PARTNER_TYPE_COL] == partner_type].copy()
        if sub.empty:
            return sub
        sub = sub.sort_values("Expected Value ($)", ascending=False).head(n)
        cols = [
            "Prospect (Account Name)",
            "Owner",
            "Stage Bucket",
            "Expected Value ($)",
            "Projected Annual Revenue ($)",
            "Probability (%)",
        ]
        cols = [c for c in cols if c in sub.columns]
        return sub[cols]

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("##### Top Sponsorship Deals")
        top_spons = _top_n(prospects, "Sponsorship")
        if top_spons.empty:
            st.caption("No Sponsorship deals yet.")
        else:
            st.dataframe(
                top_spons.style.format(
                    {
                        "Expected Value ($)": "${:,.0f}",
                        "Projected Annual Revenue ($)": "${:,.0f}",
                        "Probability (%)": "{:.0f}%",
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )

    with col_right:
        st.markdown("##### Top Public Investment Deals")
        top_public = _top_n(prospects, "Public Investment")
        if top_public.empty:
            st.caption("No Public Investment deals yet.")
        else:
            st.dataframe(
                top_public.style.format(
                    {
                        "Expected Value ($)": "${:,.0f}",
                        "Projected Annual Revenue ($)": "${:,.0f}",
                        "Probability (%)": "{:.0f}%",
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )


def build_pipeline_board(prospects: pd.DataFrame) -> None:
    """3) Pipeline stages board."""
    st.markdown("### Pipeline by Stage")

    # Ignore Dead in the main board
    df = prospects[prospects["Stage Bucket"] != "Dead"].copy()

    if df.empty:
        st.caption("No active prospects in the pipeline.")
        return

    # Tabs by partner type
    partner_types = sorted(df[PARTNER_TYPE_COL].dropna().unique().tolist())
    tab_labels = ["All"] + partner_types

    tabs = st.tabs(tab_labels)

    for label, tab in zip(tab_labels, tabs):
        with tab:
            if label == "All":
                subset = df.copy()
            else:
                subset = df[df[PARTNER_TYPE_COL] == label].copy()

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
                        "Projected Annual Revenue ($)",
                        "Contracted Annual Revenue ($)",
                        "Expected Value ($)",
                    ]
                    cols_to_show = [c for c in cols_to_show if c in stage_df.columns]

                    stage_df = stage_df[cols_to_show].sort_values(
                        "Expected Value ($)", ascending=False
                    )

                    st.markdown(
                        f"**{stage}**  \n"
                        f"<span style='font-size:12px;color:gray;'>"
                        f"{len(stage_df)} deals"
                        f"</span>",
                        unsafe_allow_html=True,
                    )

                    if stage_df.empty:
                        st.caption("No deals at this stage.")
                    else:
                        st.dataframe(
                            stage_df.style.format(
                                {
                                    "Projected Annual Revenue ($)": "${:,.0f}",
                                    "Contracted Annual Revenue ($)": "${:,.0f}",
                                    "Expected Value ($)": "${:,.0f}",
                                }
                            ),
                            hide_index=True,
                            use_container_width=True,
                        )

def build_pipeline_totals(prospects: pd.DataFrame) -> None:
    """4) Total Pipeline Value by Stage."""
    st.markdown("### Total Pipeline Value by Stage")

    # Exclude Dead from pipeline roll-ups
    df = prospects[prospects["Stage Bucket"] != "Dead"].copy()

    if df.empty:
        st.caption("No active pipeline data.")
        return

    agg = (
        df.groupby(["Stage Bucket", PARTNER_TYPE_COL])
        .agg(
            expected_total=("Expected Value ($)", "sum"),
            projected_total=("Projected Annual Revenue ($)", "sum"),
            contracted_total=("Contracted Annual Revenue ($)", "sum"),
            deal_count=("Prospect ID", "nunique"),
        )
        .reset_index()
    )

    # Summary table across all partner types
    overall = (
        df.groupby("Stage Bucket")
        .agg(
            expected_total=("Expected Value ($)", "sum"),
            deal_count=("Prospect ID", "nunique"),
        )
        .reindex(STAGE_ORDER)
        .reset_index()
    )

    # Fill NaNs (from missing stages) before casting to int
    overall["deal_count"] = overall["deal_count"].fillna(0).astype(int)
    overall["expected_total"] = overall["expected_total"].fillna(0.0)

    with st.expander("Summary table", expanded=True):
        st.dataframe(
            overall.style.format({
                "expected_total": "${:,.0f}",
                "deal_count": "{:,.0f}",  # whole numbers, no decimal places
            }),
            hide_index=True,
            use_container_width=True,
        )

    chart = (
        alt.Chart(agg)
        .mark_bar()
        .encode(
            x=alt.X("Stage Bucket:N", sort=STAGE_ORDER, title="Stage"),
            y=alt.Y("expected_total:Q", title="Total Expected Value ($)"),
            color=alt.Color(f"{PARTNER_TYPE_COL}:N", title="Type"),
            tooltip=[
                "Stage Bucket",
                PARTNER_TYPE_COL,
                alt.Tooltip("expected_total:Q", title="Expected Total", format="$.0f"),
                "deal_count",
            ],
        )
    )

    st.altair_chart(chart, use_container_width=True)

def build_recent_activity_table(contacts: pd.DataFrame) -> None:
    """5) Recent activity feed (last 10 contact events)."""
    st.markdown("### Recent Activity")

    if contacts.empty:
        st.caption("No contact activity logged yet.")
        return

    cols = [
        "Prospect (Account Name)",
        "Prospect (Sponsorship/Public)",
        "Contact Date",
        "Contact Type (email/phone/zoom/in-person)",
        "Contact Owner",
        "Contact Name",
        "Outcome (left VM/spoke/meeting set/sent deck/etc.)",
        "Next Step",
    ]
    cols = [c for c in cols if c in contacts.columns]

    recent = (
        contacts.sort_values("Contact Date", ascending=False)
        .head(10)
        .loc[:, cols]
        .copy()
    )

    st.dataframe(
        recent.style.format({"Contact Date": "{:%Y-%m-%d}"}),
        hide_index=True,
        use_container_width=True,
    )


def build_data_dictionary(data_dict: pd.DataFrame) -> None:
    """Show Data_Dictionary sheet, if present."""
    if data_dict is None or data_dict.empty:
        return

    with st.expander("Data Dictionary (field definitions)", expanded=False):
        st.dataframe(
            data_dict,
            hide_index=True,
            use_container_width=True,
        )


# -------------------------------------------------------------------
# Main app
# -------------------------------------------------------------------

def main() -> None:
    st.set_page_config(
        page_title="Bumbershoot / Cannonball – Partnership Revenue Pipeline",
        page_icon="🎟️",
        layout="wide",
    )

    st.title("Bumbershoot & Cannonball – Partnership Revenue Pipeline")
    st.caption(
        "Upload the latest Excel workbook for a split view of "
        "**Sponsorship (Corporate Partnerships)** and **Public Investment**."
    )

    # Sidebar: data source
    st.sidebar.header("Data source")

    uploaded = st.sidebar.file_uploader(
        "Upload latest Excel workbook (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
    )

    if uploaded is None:
        st.info(
            "⬅️ Please upload the latest Excel workbook to see the dashboard.\n\n"
            "Recommended: use the file you update weekly before your team meeting."
        )
        st.stop()

    # Load data
    prospects, contacts, data_dict = load_workbook(uploaded)

    # Sidebar filters
    st.sidebar.header("Filters")

    # Partner type filter
    if PARTNER_TYPE_COL not in prospects.columns:
        st.error(f"Expected column `{PARTNER_TYPE_COL}` not found in prospect data.")
        st.stop()

    partner_types = sorted(prospects[PARTNER_TYPE_COL].dropna().unique().tolist())
    selected_partner_types = st.sidebar.multiselect(
        "Partner Type",
        options=partner_types,
        default=partner_types,
    )

    # Owner filter (if present)
    if "Owner" in prospects.columns:
        owner_options = sorted(prospects["Owner"].dropna().unique().tolist())
        selected_owners = st.sidebar.multiselect(
            "Owner (AE)",
            options=owner_options,
            default=owner_options,
        )
        mask_owner = prospects["Owner"].isin(selected_owners)
    else:
        selected_owners = []
        mask_owner = True

    # Apply filters
    filtered_prospects = prospects[
        prospects[PARTNER_TYPE_COL].isin(selected_partner_types) & mask_owner
    ].copy()

    # 1) Snapshot KPIs
    st.markdown("### Snapshot")

    total_expected = filtered_prospects.get("Expected Value ($)", pd.Series(dtype=float)).sum()
    total_projected = filtered_prospects.get("Projected Annual Revenue ($)", pd.Series(dtype=float)).sum()
    total_contracted = filtered_prospects.get("Contracted Annual Revenue ($)", pd.Series(dtype=float)).sum()
    deal_count = filtered_prospects.get("Prospect ID", pd.Series(dtype=object)).nunique()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Expected Value", f"${total_expected:,.0f}")
    col2.metric("Projected Annual", f"${total_projected:,.0f}")
    col3.metric("Contracted Annual", f"${total_contracted:,.0f}")
    col4.metric("Active Prospects", int(deal_count))

    st.divider()

    # 2) Top Deals by Expected Value
    build_top_deals(filtered_prospects)
    st.divider()

    # 3) Pipeline by Stage
    build_pipeline_board(filtered_prospects)
    st.divider()

    # 4) Total Pipeline Value by Stage
    build_pipeline_totals(filtered_prospects)
    st.divider()

    # 5) Recent Activity
    build_recent_activity_table(contacts)
    st.divider()

    # Extra: Data Dictionary at the bottom
    build_data_dictionary(data_dict)


if __name__ == "__main__":
    main()


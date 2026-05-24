"""Iowa Early Vote Tracker — Streamlit app.

Drop a new Absentee County*.pdf into the data/ folder, commit, and push
to GitHub. Streamlit Cloud re-runs automatically.
"""

import io
from pathlib import Path

import streamlit as st

import breakdown

DATA_DIR = Path(__file__).parent / "data"

# ── Cached data loading ───────────────────────────────────────────────────────

@st.cache_data
def _pdf_mtime() -> float:
    p = breakdown.find_inputs(DATA_DIR).get("county")
    return float(p.stat().st_mtime) if (p and p.exists()) else 0.0


@st.cache_data
def _load_data(mtime: float):
    inputs    = breakdown.find_inputs(DATA_DIR)
    raw       = breakdown.parse_county_pdf(inputs["county"])
    raw["market"] = raw["county"].map(breakdown.COUNTY_TO_MARKET)
    voter_reg = breakdown.parse_voter_reg_pdf(inputs["voter_reg"])
    share_df  = breakdown.build_share_analysis(raw, voter_reg)
    county_df = breakdown.build_county_overunder(raw, voter_reg)
    return raw, voter_reg, share_df, county_df, inputs["county"].name, inputs.get("cd")


@st.cache_resource
def _load_geo():
    geojson    = breakdown._load_geojson()
    boundaries = breakdown._build_market_boundaries(geojson)
    return geojson, boundaries


@st.cache_data
def _build_figs(mtime: float) -> dict:
    raw, voter_reg, share_df, county_df, _, _ = _load_data(mtime)
    geojson, boundaries = _load_geo()

    df  = county_df.copy()
    mkt = share_df[share_df["Market"] != "TOTAL"].set_index("Market")

    def ou(v):
        return f"{'+' if v >= 0 else ''}{v:.2f} pp"

    df["gop_ou_label"]  = df["gop_overunder"].apply(ou)
    df["dem_ou_label"]  = df["dem_overunder"].apply(ou)
    df["gop_early_fmt"] = df["gop_early_pct"].apply(lambda v: f"{v:.2f}%")
    df["gop_reg_fmt"]   = df["gop_reg_pct"].apply(lambda v: f"{v:.2f}%")
    df["dem_early_fmt"] = df["dem_early_pct"].apply(lambda v: f"{v:.2f}%")
    df["dem_reg_fmt"]   = df["dem_reg_pct"].apply(lambda v: f"{v:.2f}%")

    mdf = df[["county", "fips", "market"]].copy()
    for col, src in [
        ("gop_mkt_ou",    "GOP Over/Under"), ("dem_mkt_ou",    "Dem Over/Under"),
        ("gop_mkt_early", "GOP Early %"),    ("dem_mkt_early", "Dem Early %"),
        ("gop_mkt_reg",   "GOP Reg %"),      ("dem_mkt_reg",   "Dem Reg %"),
    ]:
        mdf[col] = mdf["market"].map(mkt[src])
    mdf["gop_mkt_ou_label"]  = mdf["gop_mkt_ou"].apply(ou)
    mdf["dem_mkt_ou_label"]  = mdf["dem_mkt_ou"].apply(ou)
    mdf["gop_mkt_early_fmt"] = mdf["gop_mkt_early"].apply(lambda v: f"{v:.2f}%")
    mdf["gop_mkt_reg_fmt"]   = mdf["gop_mkt_reg"].apply(lambda v: f"{v:.2f}%")
    mdf["dem_mkt_early_fmt"] = mdf["dem_mkt_early"].apply(lambda v: f"{v:.2f}%")
    mdf["dem_mkt_reg_fmt"]   = mdf["dem_mkt_reg"].apply(lambda v: f"{v:.2f}%")

    county_labels = ["Early %", "Reg %", "Over/Under"]
    market_labels = ["Mkt Early %", "Mkt Reg %", "Over/Under"]

    return {
        "gop-county": breakdown._make_choropleth(
            df, geojson, "gop_overunder",
            ["gop_early_fmt", "gop_reg_fmt", "gop_ou_label"],
            "GOP Early Vote Over/Under — by County", boundaries, county_labels),
        "gop-market": breakdown._make_choropleth(
            mdf, geojson, "gop_mkt_ou",
            ["gop_mkt_early_fmt", "gop_mkt_reg_fmt", "gop_mkt_ou_label"],
            "GOP Early Vote Over/Under — by Media Market", boundaries, market_labels),
        "dem-county": breakdown._make_choropleth(
            df, geojson, "dem_overunder",
            ["dem_early_fmt", "dem_reg_fmt", "dem_ou_label"],
            "Dem Early Vote Over/Under — by County", boundaries, county_labels),
        "dem-market": breakdown._make_choropleth(
            mdf, geojson, "dem_mkt_ou",
            ["dem_mkt_early_fmt", "dem_mkt_reg_fmt", "dem_mkt_ou_label"],
            "Dem Early Vote Over/Under — by Media Market", boundaries, market_labels),
    }


@st.cache_data
def _excel_bytes(mtime: float) -> bytes:
    raw, voter_reg, share_df, _, _, cd_pdf = _load_data(mtime)
    market_df = breakdown.build_summary(raw, "market")
    if cd_pdf and cd_pdf.exists():
        cd_raw = breakdown.parse_cd_pdf(cd_pdf)
        cd_df  = breakdown.build_summary(cd_raw, "cd")
    else:
        raw = raw.copy()
        raw["cd"] = raw["county"].map(breakdown.COUNTY_TO_CD)
        cd_df = breakdown.build_summary(raw, "cd")
    cty_df = breakdown.build_summary(raw, "county")
    cty_df.insert(1, "Media Market",
                  cty_df["Geography"].map(breakdown.COUNTY_TO_MARKET).fillna(""))
    cty_df.insert(2, "Cong. District",
                  cty_df["Geography"].map(breakdown.COUNTY_TO_CD).fillna(""))
    buf = io.BytesIO()
    breakdown.write_excel(market_df, cd_df, cty_df, share_df, buf)
    return buf.getvalue()


# ── Layout ───────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Iowa Early Vote Tracker", layout="wide")

mtime = _pdf_mtime()
_, _, share_df, _, pdf_name, _ = _load_data(mtime)
figs = _build_figs(mtime)

# Sidebar
with st.sidebar:
    st.title("Iowa Early Vote")
    st.divider()
    party = st.radio("Party", ["GOP", "Dem"])
    view  = st.radio("View",  ["County", "Market"])
    st.divider()
    st.caption(f"Source: {pdf_name}")
    xl_name = pdf_name.replace(".pdf", "_Breakdown.xlsx").replace(" ", "_")
    st.download_button(
        label="Download Excel",
        data=_excel_bytes(mtime),
        file_name=xl_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Map
fig_key = f"{party.lower()}-{view.lower()}"
st.plotly_chart(figs[fig_key], width="stretch")

# Market share table
st.subheader(f"{party} — Media Market Share")
p = party  # "GOP" or "Dem"
tbl = share_df[["Market", f"{p} Received", f"{p} Early %", f"{p} Reg %", f"{p} Over/Under"]].copy()
tbl.columns = ["Market", "Received", "Early %", "Reg %", "Over/Under"]

def _color_ou(val):
    if isinstance(val, (int, float)):
        if val > 0:
            return "color: #1a9850; font-weight: bold"
        if val < 0:
            return "color: #d73027; font-weight: bold"
    return ""

st.dataframe(
    tbl.style
        .format({"Received": "{:,.0f}", "Early %": "{:.2f}%",
                 "Reg %": "{:.2f}%", "Over/Under": "{:+.2f}%"})
        .map(_color_ou, subset=["Over/Under"]),
    width="stretch",
    hide_index=True,
)

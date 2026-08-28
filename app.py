# ==============================================================================
# app.py — RO Membrane Performance Dashboard (DSSRO & LT4)
# ==============================================================================

import operator

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ------------------------------------------------------------------------------
# Page config
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="RO Membrane Performance Dashboard",
    page_icon="🔬",
    layout="wide",
)

st.title("🔬 RO Membrane Performance Analysis Dashboard")
st.markdown(
    """
    Upload your RO operational data (Excel) and this dashboard will:
    1. Clean & standardize raw sensor columns (works for **DSSRO** or **LT4** systems)
    2. Remove non-physical (NaN / negative) readings
    3. Isolate truly **operational** membrane conditions (filter out flush/CIP cycles)
    4. Compute normalized performance indicators (SSPn, QSPn, DPn, WTCn, STCn, ASPn)
    5. Plot each indicator (with optional regression trendline) against the
       Recovery Setpoint over time
    """
)

# ==============================================================================
# SIDEBAR — 0) System Type
# ==============================================================================
st.sidebar.header("⚙️ System Type")
system_type = st.sidebar.radio(
    "Select the membrane system this data comes from:",
    ["DSSRO", "LT4"],
    horizontal=True,
)
st.sidebar.caption(f"Currently configured for: **{system_type}**")

# ==============================================================================
# SIDEBAR — 1) Upload Data
# ==============================================================================
st.sidebar.header("1️⃣ Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if system_type == "DSSRO":
    expected_raw_cols_display = """
    - `Timestamp (yyyy-MM-dd HH:mm:ss)`
    - `RO-Mixed feed Temperature-°F`
    - `RO-Permeate flowrate-m³/hr`
    - `RO-Concentrate Recircualtion Flow-m³/hr`
    - `RO-Concentrate pressure-bar`
    - `RO-Element feed Pressure-bar`
    - `RO-Permeate line pressure-bar`
    - `RO-Recovery Actual Setpoint-%`
    - `RO-Raw feed pH-pH`
    - `RO-Raw feed Conductivity-mS/cm`
    - `RO-Permeate Conductivity-mS/cm`
    """
else:
    expected_raw_cols_display = """
    - `Timestamp (yyyy-MM-dd HH:mm:ss)`
    - `ProFlex RO FS-421-Feed Temperature-°C`
    - `ProFlex RO FS-421-Permeate Flow Rate-L/s`
    - `ProFlex RO FS-421-Recycle Flow Rate-L/s`
    - `ProFlex RO FS-421-Concentrate Outlet Pressure-kPa`
    - `ProFlex RO FS-421-RO Feed Inlet Pressure-kPa`
    - `ProFlex RO FS-421-Recovery-%`
    - `ProFlex RO FS-421-Feed pH-pH`
    - `ProFlex RO FS-421-Feed Conductivity-µS/cm`
    - `ProFlex RO FS-421-Permeate Conductivity-µS/cm`

    *(Permeate line pressure is not sensed on LT4 — it is set to a constant 1 psi, matching the original analysis script.)*
    """

with st.sidebar.expander("ℹ️ Expected raw column names for this system"):
    st.write(expected_raw_cols_display)

if uploaded_file is None:
    st.info("👈 Upload an Excel file and select a system type in the sidebar to begin.")
    st.stop()

# ==============================================================================
# SIDEBAR — 2) Constants & Conversions (mode-dependent defaults)
# ==============================================================================
st.sidebar.header("2️⃣ Constants & Conversions")

with st.sidebar.expander("Unit Conversions"):
    CONC_PPM_CONV = st.number_input("uS/cm → ppm factor", value=2.143, format="%.4f")

    if system_type == "DSSRO":
        M3H_TO_GPM = st.number_input("m3/h → gpm factor", value=4.40286, format="%.5f")
        BAR_TO_PSI = st.number_input("bar → psi factor", value=14.5, format="%.2f")
        KPA_TO_PSI = None
        LPS_TO_GPM = None
    else:
        KPA_TO_PSI = st.number_input("kPa → psi divisor", value=6.895, format="%.3f")
        LPS_TO_GPM = st.number_input("L/s → gpm factor", value=15.85, format="%.2f")
        M3H_TO_GPM = None
        BAR_TO_PSI = None

with st.sidebar.expander("System Constants"):
    if system_type == "DSSRO":
        EPV = st.number_input("Elements per Vessel (EPV)", value=5.00, key=f"EPV_{system_type}")
        V = st.number_input("Number of Vessels (V)", value=1.00, key=f"V_{system_type}")
        EMAe = st.number_input("Element Membrane Area (sq ft)", value=90.00, key=f"EMAe_{system_type}")
    else:
        EPV = st.number_input("Elements per Vessel (EPV)", value=1.00, key=f"EPV_{system_type}")
        V = st.number_input("Number of Vessels (V)", value=7.00, key=f"V_{system_type}")
        EMAe = st.number_input("Element Membrane Area (sq ft)", value=90.00, key=f"EMAe_{system_type}")

with st.sidebar.expander("Standard Test Conditions"):
    if system_type == "DSSRO":
        Qpe = st.number_input("Standard Permeate Flow (GPD)", value=2600.00, key=f"Qpe_{system_type}")
        NDPe = st.number_input("Standard NDP (psi)", value=225.00, key=f"NDPe_{system_type}")
        Cfe = st.number_input("Standard Feed Concentration (ppm)", value=2000.00, key=f"Cfe_{system_type}")
    else:
        Qpe = st.number_input("Standard Permeate Flow (GPD)", value=2225.00, key=f"Qpe_{system_type}")
        NDPe = st.number_input("Standard NDP (psi)", value=115.00, key=f"NDPe_{system_type}")
        Cfe = st.number_input("Standard Feed Concentration (ppm)", value=500.00, key=f"Cfe_{system_type}")
    Ke = st.number_input("Constant Ke", value=2700.00)

with st.sidebar.expander("Temperature Assumptions"):
    TR = st.number_input("Room Temperature (K)", value=298.15)
    TS = st.number_input("Standard Temperature (K)", value=273.15)

# ==============================================================================
# LOAD & STANDARDIZE RAW DATA
# ==============================================================================
try:
    raw_df_input = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Could not read Excel file: {e}")
    st.stop()


def load_dssro(df_in, m3h_to_gpm, bar_to_psi):
    rename_map = {
        'Timestamp (yyyy-MM-dd HH:mm:ss)': 'Date',
        'RO-Mixed feed Temperature-°F': 'T (C)',
        'RO-Permeate flowrate-m³/hr': 'Qp_raw',
        'RO-Concentrate Recircualtion Flow-m³/hr': 'Qc_raw',
        'RO-Concentrate pressure-bar': 'Pc_raw',
        'RO-Element feed Pressure-bar': 'Pf_raw',
        'RO-Permeate line pressure-bar': 'Pp_raw',
        'RO-Recovery Actual Setpoint-%': 'Rsp (%)',
        'RO-Raw feed pH-pH': 'pH',
        'RO-Raw feed Conductivity-mS/cm': 'Cf (uS/cm)',
        'RO-Permeate Conductivity-mS/cm': 'Cp (uS/cm)',
    }
    missing = [c for c in rename_map if c not in df_in.columns]
    df = df_in.rename(columns=rename_map)

    expected_raw = ['Date', 'T (C)', 'pH', 'Cf (uS/cm)', 'Cp (uS/cm)',
                     'Qp_raw', 'Qc_raw', 'Pf_raw', 'Pc_raw', 'Pp_raw', 'Rsp (%)']
    df = df.reindex(columns=expected_raw)

    df['Cf (uS/cm)'] = df['Cf (uS/cm)'] * 1000  # Insight inconsistency fix
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    df['Qp (gpm)'] = df['Qp_raw'] * m3h_to_gpm
    df['Qc (gpm)'] = df['Qc_raw'] * m3h_to_gpm
    df['Pf (psi)'] = df['Pf_raw'] * bar_to_psi
    df['Pc (psi)'] = df['Pc_raw'] * bar_to_psi
    df['Pp (psi)'] = df['Pp_raw'] * bar_to_psi

    final_cols = ['Date', 'T (C)', 'pH', 'Cf (uS/cm)', 'Cp (uS/cm)',
                  'Qp (gpm)', 'Qc (gpm)', 'Pf (psi)', 'Pc (psi)', 'Pp (psi)', 'Rsp (%)']
    return df[final_cols], missing


def load_lt4(df_in, kpa_to_psi, lps_to_gpm):
    rename_map = {
        'Timestamp (yyyy-MM-dd HH:mm:ss)': 'Date',
        'ProFlex RO FS-421-Feed Temperature-°C': 'T (C)',
        'ProFlex RO FS-421-Permeate Flow Rate-L/s': 'Qp_raw',
        'ProFlex RO FS-421-Recycle Flow Rate-L/s': 'Qc_raw',
        'ProFlex RO FS-421-Concentrate Outlet Pressure-kPa': 'Pc_raw',
        'ProFlex RO FS-421-RO Feed Inlet Pressure-kPa': 'Pf_raw',
        'ProFlex RO FS-421-Recovery-%': 'Rsp (%)',
        'ProFlex RO FS-421-Feed pH-pH': 'pH',
        'ProFlex RO FS-421-Feed Conductivity-µS/cm': 'Cf (uS/cm)',
        'ProFlex RO FS-421-Permeate Conductivity-µS/cm': 'Cp (uS/cm)',
    }
    missing = [c for c in rename_map if c not in df_in.columns]
    df = df_in.rename(columns=rename_map)

    expected_raw = ['Date', 'T (C)', 'pH', 'Cf (uS/cm)', 'Cp (uS/cm)',
                     'Qp_raw', 'Qc_raw', 'Pf_raw', 'Pc_raw', 'Rsp (%)']
    df = df.reindex(columns=expected_raw)

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    df['Qp (gpm)'] = df['Qp_raw'] * lps_to_gpm
    df['Qc (gpm)'] = df['Qc_raw'] * lps_to_gpm
    df['Pf (psi)'] = df['Pf_raw'] / kpa_to_psi
    df['Pc (psi)'] = df['Pc_raw'] / kpa_to_psi
    df['Pp (psi)'] = 1.0  # Not sensed on LT4; constant per original script

    final_cols = ['Date', 'T (C)', 'pH', 'Cf (uS/cm)', 'Cp (uS/cm)',
                  'Qp (gpm)', 'Qc (gpm)', 'Pf (psi)', 'Pc (psi)', 'Pp (psi)', 'Rsp (%)']
    return df[final_cols], missing


if system_type == "DSSRO":
    raw_df, missing_cols = load_dssro(raw_df_input, M3H_TO_GPM, BAR_TO_PSI)
else:
    raw_df, missing_cols = load_lt4(raw_df_input, KPA_TO_PSI, LPS_TO_GPM)

if missing_cols:
    st.sidebar.warning(f"⚠️ Missing expected raw columns: {missing_cols}")

# ------------------------------------------------------------------------------
# Remove non-physical values (NaN / negative)
# ------------------------------------------------------------------------------
org_df = raw_df.copy()
rows_before = len(org_df)
org_df.dropna(inplace=True)
removed_nan = rows_before - len(org_df)

rows_before_neg = len(org_df)
numeric_cols = org_df.select_dtypes(include=['number'])
org_df = org_df[numeric_cols.ge(0).all(axis=1)]
removed_neg = rows_before_neg - len(org_df)

if len(org_df) == 0:
    st.error("No data remains after cleaning. Check your file / system type selection.")
    st.stop()

# ==============================================================================
# SIDEBAR — 3) Filtering Conditions (mode-dependent)
# ==============================================================================
st.sidebar.header("3️⃣ Filtering Conditions")

_standard_cols = ['T (C)', 'pH', 'Cf (uS/cm)', 'Cp (uS/cm)', 'Qp (gpm)',
                   'Qc (gpm)', 'Pf (psi)', 'Pc (psi)', 'Pp (psi)', 'Rsp (%)']

if system_type == "DSSRO":
    apply_cond1 = st.sidebar.checkbox(
        "Cond. 1: Flush/CIP detection (Pf > mult × next Pf AND Pf > threshold)",
        value=True, key="dssro_c1"
    )
    c1_mult = st.sidebar.number_input("Flush multiplier", value=3.0, key="dssro_c1_mult")
    c1_thresh = st.sidebar.number_input("Pf threshold (psi)", value=87.0, key="dssro_c1_thresh")

    apply_cond2 = st.sidebar.checkbox("Cond. 2: Pp > 0", value=True, key="dssro_c2")

    apply_cond3 = st.sidebar.checkbox("Cond. 3: Qc ≥ threshold", value=True, key="dssro_c3")
    c3_thresh = st.sidebar.number_input("Qc threshold (gpm)", value=4.403, key="dssro_c3_thresh")

    apply_cond4 = apply_cond5 = False  # unused in DSSRO mode

else:  # LT4
    apply_cond1 = st.sidebar.checkbox("Cond. 1: Pf ≥ threshold", value=True, key="lt4_c1")
    c1_thresh = st.sidebar.number_input("Pf threshold (psi)", value=80.0, key="lt4_c1_thresh")

    apply_cond2 = st.sidebar.checkbox("Cond. 2: Pc ≥ threshold", value=True, key="lt4_c2")
    c2_thresh = st.sidebar.number_input("Pc threshold (psi)", value=80.0, key="lt4_c2_thresh")

    apply_cond3 = st.sidebar.checkbox("Cond. 3: Qc ≥ threshold", value=True, key="lt4_c3")
    c3_thresh = st.sidebar.number_input("Qc threshold (gpm)", value=2.0, key="lt4_c3_thresh")

    apply_cond4 = st.sidebar.checkbox("Cond. 4: Qp ≥ threshold", value=True, key="lt4_c4")
    c4_thresh = st.sidebar.number_input("Qp threshold (gpm)", value=4.0, key="lt4_c4_thresh")

    apply_cond5 = st.sidebar.checkbox("Cond. 5: pH ≥ threshold", value=True, key="lt4_c5")
    c5_thresh = st.sidebar.number_input("pH threshold", value=7.0, key="lt4_c5_thresh")

st.sidebar.markdown("**Optional custom condition**")
add_condition_enabled = st.sidebar.checkbox("Enable custom condition", value=False)
custom_col = custom_op = None
custom_val = 0.0
if add_condition_enabled:
    custom_col = st.sidebar.selectbox("Column", _standard_cols)
    custom_op = st.sidebar.selectbox("Operator", [">", ">=", "<", "<=", "==", "!="])
    custom_val = st.sidebar.number_input("Threshold value", value=0.0)

# ==============================================================================
# SIDEBAR — 4) Plot Ranges & Date Selection
# ==============================================================================
st.sidebar.header("4️⃣ Plot Ranges & Date Selection")

with st.sidebar.expander("📅 Date Range", expanded=True):
    min_date = org_df['Date'].min().date()
    max_date = org_df['Date'].max().date()
    date_range = st.slider(
        "Select date range to analyze",
        min_value=min_date,
        max_value=max_date,
        value=(min_date, max_date),
        format="YYYY-MM-DD",
    )
    start_date, end_date = date_range

with st.sidebar.expander("Adjust plot Y-axis ranges"):
    if system_type == "DSSRO":
        ylim_sspn = st.slider("SSPn (%)", 0.0, 20.0, (0.0, 5.0))
        ylim_qspn = st.slider("QSPn (gpm)", 0.0, 20.0, (0.0, 7.0))
        ylim_dpn = st.slider("DPn (psi)", 0.0, 100.0, (0.0, 35.0))
    else:
        ylim_sspn = st.slider("SSPn (%)", 0.0, 20.0, (0.0, 2.0))
        ylim_qspn = st.slider("QSPn (gpm)", 0.0, 20.0, (0.0, 10.0))
        ylim_dpn = st.slider("DPn (psi)", 0.0, 100.0, (0.0, 35.0))
    ylim_wtcn = st.slider("WTCn (m/s-kPa) x1e-8", 0.0, 5.0, (0.0, 1.2))
    ylim_stcn = st.slider("STCn (m/s) x1e-7", 0.0, 10.0, (0.0, 2.5))
    ylim_aspn = st.slider("ASPn (%)", 0.0, 20.0, (0.0, 5.0))

with st.sidebar.expander("Chart display options"):
    show_rangeslider = st.checkbox("Show mini range-slider under each chart", value=True)

# ==============================================================================
# SIDEBAR — 5) Regression Settings
# ==============================================================================
st.sidebar.header("5️⃣ Regression Settings")
regression_type = st.sidebar.selectbox("Trendline type", ["None", "Linear", "Polynomial"])
poly_degree = 2
if regression_type == "Polynomial":
    poly_degree = st.sidebar.slider("Polynomial degree", 2, 6, 2)

# ==============================================================================
# STEP 1 — RAW / CLEANED DATA (display)
# ==============================================================================
st.header("Step 1 — Raw Data")
with st.expander("Show standardized (renamed) data", expanded=False):
    st.dataframe(raw_df, use_container_width=True)
    st.caption(f"Rows: {len(raw_df)}")

c1, c2, c3 = st.columns(3)
c1.metric("Removed (NaN rows)", removed_nan)
c2.metric("Removed (Negative-value rows)", removed_neg)
c3.metric("Remaining rows", len(org_df))

st.subheader("Cleaned data (org_df)")
st.dataframe(org_df, use_container_width=True)
st.download_button(
    "⬇️ Download cleaned data (CSV)",
    org_df.to_csv(index=False).encode(),
    file_name="cleaned_data.csv",
)

# ------------------------------------------------------------------------------
# Apply the sidebar Date Range filter
# ------------------------------------------------------------------------------
org_df = org_df[
    (org_df['Date'].dt.date >= start_date) & (org_df['Date'].dt.date <= end_date)
].copy()

st.caption(f"📅 Date range applied: **{start_date} → {end_date}** — {len(org_df)} rows remain")

if len(org_df) == 0:
    st.error("No data in the selected date range — widen the Date Range slider in the sidebar.")
    st.stop()

# ==============================================================================
# STEP 2 — FILTER FOR OPERATIONAL CONDITIONS (mode-dependent)
# ==============================================================================
st.header("Step 2 — Operational Filtering")

combined_condition = pd.Series(True, index=org_df.index)

if system_type == "DSSRO":
    condition1 = (org_df['Pf (psi)'] > c1_mult * org_df['Pf (psi)'].shift(-1)) & (org_df['Pf (psi)'] > c1_thresh)
    condition2 = (org_df['Pp (psi)'] > 0)
    condition3 = (org_df['Qc (gpm)'] >= c3_thresh)

    if apply_cond1:
        combined_condition &= condition1
    if apply_cond2:
        combined_condition &= condition2
    if apply_cond3:
        combined_condition &= condition3

else:  # LT4
    condition1 = (org_df['Pf (psi)'] >= c1_thresh)
    condition2 = (org_df['Pc (psi)'] >= c2_thresh)
    condition3 = (org_df['Qc (gpm)'] >= c3_thresh)
    condition4 = (org_df['Qp (gpm)'] >= c4_thresh)
    condition5 = (org_df['pH'] >= c5_thresh)

    if apply_cond1:
        combined_condition &= condition1
    if apply_cond2:
        combined_condition &= condition2
    if apply_cond3:
        combined_condition &= condition3
    if apply_cond4:
        combined_condition &= condition4
    if apply_cond5:
        combined_condition &= condition5

if add_condition_enabled and custom_col:
    ops = {
        ">": operator.gt, ">=": operator.ge,
        "<": operator.lt, "<=": operator.le,
        "==": operator.eq, "!=": operator.ne,
    }
    combined_condition &= ops[custom_op](org_df[custom_col], custom_val)

fil_df = org_df[combined_condition].copy()

st.metric("Rows removed by filtering", len(org_df) - len(fil_df))
st.subheader("Filtered / operational data (fil_df)")
st.dataframe(fil_df, use_container_width=True)
st.download_button(
    "⬇️ Download filtered data (CSV)",
    fil_df.to_csv(index=False).encode(),
    file_name="filtered_data.csv",
)

if len(fil_df) == 0:
    st.error("No rows remain after filtering — relax the conditions in the sidebar.")
    st.stop()

# ==============================================================================
# STEP 3 — CALCULATED METRICS (unified formulas — same for both systems)
# ==============================================================================
st.header("Step 3 — Calculated Performance Metrics")

calc_df = fil_df[['Date', 'T (C)']].copy()

calc_df['Rec (%)'] = fil_df['Qp (gpm)'] / (fil_df['Qp (gpm)'] + fil_df['Qc (gpm)'])
calc_df['DP (psi)'] = fil_df['Pf (psi)'] - fil_df['Pc (psi)']

calc_df['Cfo (ppm)'] = fil_df['Cf (uS/cm)'] / CONC_PPM_CONV
calc_df['Cpo (ppm)'] = fil_df['Cp (uS/cm)'] / CONC_PPM_CONV

calc_df['Qco (gpm)'] = fil_df['Qc (gpm)']
calc_df['Qpo (gpm)'] = fil_df['Qp (gpm)']

calc_df['Pfo (psi)'] = fil_df['Pf (psi)']
calc_df['Pco (psi)'] = fil_df['Pc (psi)']
calc_df['Ppo (psi)'] = fil_df['Pp (psi)']

calc_df['ConF (ppm)'] = np.log(1 / (1 - calc_df['Rec (%)'])) / calc_df['Rec (%)']
calc_df['Cf ave (ppm)'] = calc_df['Cfo (ppm)'] * calc_df['ConF (ppm)']

calc_df['FOP ave (psi)'] = (
    calc_df['Cf ave (ppm)'] * 0.03851 * (TS + calc_df['T (C)'])
    / (1000 - calc_df['Cf ave (ppm)'] / 1000)
)
calc_df['POP (psi)'] = (
    calc_df['Cpo (ppm)'] * 0.03851 * (TS + calc_df['T (C)'])
    / (1000 - calc_df['Cpo (ppm)'] / 1000)
)

calc_df['DPo (psi)'] = calc_df['Pfo (psi)'] - calc_df['Pco (psi)']

calc_df['NDP (psi)'] = (
    calc_df['Pfo (psi)']
    - 0.5 * calc_df['DPo (psi)']
    - calc_df['Ppo (psi)']
    - calc_df['FOP ave (psi)']
    + calc_df['POP (psi)']
)

calc_df['SFX (gfd)'] = 1440 * calc_df['Qpo (gpm)'] / (EPV * V * EMAe)

calc_df['SSP (%)'] = 100 * calc_df['Cpo (ppm)'] / calc_df['Cf ave (ppm)']
calc_df['SSR (%)'] = 100 - calc_df['SSP (%)']

calc_df['TCF (no units)'] = np.exp(Ke * (1 / TR - 1 / (TS + calc_df['T (C)'])))

calc_df['ASPn (%)'] = (
    calc_df['SSP (%)'] * (calc_df['SFX (gfd)'] / (Qpe / EMAe)) / calc_df['TCF (no units)']
)

st.dataframe(calc_df, use_container_width=True)
st.download_button(
    "⬇️ Download calculated metrics (CSV)",
    calc_df.to_csv(index=False).encode(),
    file_name="calculated_metrics.csv",
)

# ==============================================================================
# STEP 4 — REFERENCE POINTS & NORMALIZED METRICS (unified formulas)
# ==============================================================================
st.header("Step 4 — Normalized Metrics")

Qpr = fil_df['Qp (gpm)'].iloc[0]
TCFr = calc_df['TCF (no units)'].iloc[0]
NDPr = calc_df['NDP (psi)'].iloc[0]
Qcr = fil_df['Qc (gpm)'].iloc[0]

st.caption(
    f"Reference point (first operational row): "
    f"Qpr = {Qpr:.3f} gpm, TCFr = {TCFr:.4f}, NDPr = {NDPr:.2f} psi, Qcr = {Qcr:.3f} gpm"
)

norm_df = pd.DataFrame({'Date': fil_df['Date']}).copy()

norm_df['SSPn (%)'] = (
    calc_df['SSP (%)'] * (fil_df['Qp (gpm)'] / Qpr) * (TCFr / calc_df['TCF (no units)'])
)

norm_df['QSPn (gpm)'] = (
    fil_df['Qp (gpm)'] * NDPr / calc_df['NDP (psi)'] * (TCFr / calc_df['TCF (no units)'])
)

norm_df['DPn (psi)'] = (
    calc_df['DP (psi)']
    * (Qpr / 2 + Qcr) ** 1.4
    / ((fil_df['Qp (gpm)'] / 2 + fil_df['Qc (gpm)'])) ** 1.4
) * (1 + 0.01 * (calc_df['T (C)'] - 25))

norm_df['WTCn (m/s-kPa)'] = (
    0.00000006849 * calc_df['SFX (gfd)'] / calc_df['NDP (psi)'] / calc_df['TCF (no units)']
)

norm_df['STCn (m/s)'] = (
    calc_df['Qpo (gpm)'] * calc_df['Cpo (ppm)'] / calc_df['TCF (no units)']
    / 264.17 / 60 / (EPV * V * EMAe * 0.0929)
    / (calc_df['Cf ave (ppm)'] - calc_df['Cpo (ppm)'])
)

st.dataframe(norm_df, use_container_width=True)
st.download_button(
    "⬇️ Download normalized metrics (CSV)",
    norm_df.to_csv(index=False).encode(),
    file_name="normalized_metrics.csv",
)

# ==============================================================================
# STEP 5 — INTERACTIVE PLOTS WITH REGRESSION
# ==============================================================================
st.header("Step 5 — Diagnostic Plots")


def compute_regression(x_numeric, y, degree):
    mask = ~(np.isnan(x_numeric) | np.isnan(y))
    x_clean = x_numeric[mask]
    y_clean = y[mask]

    if len(x_clean) < degree + 1:
        return None, None, None

    coeffs = np.polyfit(x_clean, y_clean, degree)
    poly = np.poly1d(coeffs)
    y_fit = poly(x_numeric)

    y_pred_clean = poly(x_clean)
    ss_res = np.sum((y_clean - y_pred_clean) ** 2)
    ss_tot = np.sum((y_clean - np.mean(y_clean)) ** 2)
    r2 = 1 - ss_res / ss_tot if ss_tot != 0 else float('nan')

    n = len(coeffs)
    terms = []
    for i, c in enumerate(coeffs):
        power = n - i - 1
        if power == 0:
            terms.append(f"{c:.4g}")
        elif power == 1:
            terms.append(f"{c:.4g}x")
        else:
            terms.append(f"{c:.4g}x^{power}")
    equation = "y = " + " + ".join(terms)

    return y_fit, equation, r2


def plot_normalized_data_interactive(
    data_df, fil_df, norm_col, norm_label, title, norm_ylim,
    regression_type, poly_degree, show_rangeslider
):
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Scatter(
            x=data_df['Date'], y=data_df[norm_col],
            mode='markers', name=norm_label,
            marker=dict(color='darkblue', symbol='diamond', size=7),
        ),
        secondary_y=False,
    )

    if regression_type != "None":
        x_numeric = (data_df['Date'] - data_df['Date'].min()).dt.total_seconds().values
        y = data_df[norm_col].values.astype(float)
        degree = 1 if regression_type == "Linear" else poly_degree

        y_fit, equation, r2 = compute_regression(x_numeric, y, degree)

        if y_fit is not None:
            order = np.argsort(x_numeric)
            fig.add_trace(
                go.Scatter(
                    x=data_df['Date'].values[order], y=y_fit[order],
                    mode='lines',
                    name=f'{regression_type} fit (R²={r2:.3f})',
                    line=dict(color='orange', width=3, dash='dash'),
                ),
                secondary_y=False,
            )
            fig.add_annotation(
                text=f"{equation}<br>R² = {r2:.4f}",
                xref="paper", yref="paper",
                x=0.01, y=0.99, showarrow=False,
                bgcolor="rgba(255,255,255,0.75)",
                bordercolor="orange", borderwidth=1,
                align="left",
            )

    fig.add_trace(
        go.Scatter(
            x=fil_df['Date'], y=fil_df['Rsp (%)'],
            mode='markers', name='Rsp (%)',
            marker=dict(color='red', symbol='square', size=6),
        ),
        secondary_y=True,
    )

    fig.update_layout(
        title=title,
        xaxis_title='Date',
        hovermode='x unified',
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        height=600,
        xaxis=dict(
            rangeslider=dict(visible=show_rangeslider),
            rangeselector=dict(
                buttons=list([
                    dict(count=7, label="1W", step="day", stepmode="backward"),
                    dict(count=1, label="1M", step="month", stepmode="backward"),
                    dict(count=3, label="3M", step="month", stepmode="backward"),
                    dict(count=6, label="6M", step="month", stepmode="backward"),
                    dict(count=1, label="YTD", step="year", stepmode="todate"),
                    dict(step="all", label="All"),
                ]),
                bgcolor="rgba(230,230,230,0.7)",
            ),
        ),
    )
    fig.update_yaxes(title_text=norm_label, range=list(norm_ylim), secondary_y=False)
    fig.update_yaxes(title_text='R_sp (%)', range=[0, 100], secondary_y=True)

    return fig


plots_config = [
    (norm_df, 'SSPn (%)', 'SSPn (%)',
     'Normalized % Salt Passage (SSPn) vs Recovery Setpoint',
     'SSPn_and_R_sp_Over_Time.html', ylim_sspn),
    (norm_df, 'QSPn (gpm)', 'Permeate Flow Rate (gpm)',
     'Permeate Flow Rate (QSPn) vs Recovery Setpoint',
     'QSPn_and_R_sp_Over_Time.html', ylim_qspn),
    (norm_df, 'DPn (psi)', 'DPn (psi)',
     'Normalized Differential Pressure (DPn) vs Recovery Setpoint',
     'DPn_and_R_sp_Over_Time.html', ylim_dpn),
    (norm_df, 'WTCn (m/s-kPa)', 'WTCn (m/s-kPa)',
     'Water Transport Coefficient (WTCn) vs Recovery Setpoint',
     'WTCn_and_R_sp_Over_Time.html', tuple(v * 1e-8 for v in ylim_wtcn)),
    (norm_df, 'STCn (m/s)', 'STCn (m/s)',
     'Salt Transport Coefficient (STCn) vs Recovery Setpoint',
     'STCn_and_R_sp_Over_Time.html', tuple(v * 1e-7 for v in ylim_stcn)),
    (calc_df, 'ASPn (%)', 'ASPn (%)',
     '% Salt Passage (ASPn) vs Recovery Setpoint',
     'ASPn_and_R_sp_Over_Time.html', ylim_aspn),
]

tab_labels = ["SSPn", "QSPn", "DPn", "WTCn", "STCn", "ASPn"]
tabs = st.tabs(tab_labels)

for tab, (df_source, col, label, title, fname, ylim) in zip(tabs, plots_config):
    with tab:
        fig = plot_normalized_data_interactive(
            df_source, fil_df, col, label, title, ylim,
            regression_type, poly_degree, show_rangeslider
        )
        st.plotly_chart(fig, use_container_width=True)

        html_bytes = fig.to_html(include_plotlyjs='cdn').encode()
        st.download_button(
            f"⬇️ Download interactive chart ('{fname}')",
            data=html_bytes,
            file_name=fname,
            mime='text/html',
        )

st.success(
    f"✅ Analysis complete for **{system_type}** system. "
    "Adjust settings in the sidebar and the dashboard updates automatically."
)

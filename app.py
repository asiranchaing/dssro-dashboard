# ==============================================================================
# app.py — DSSRO / RO Membrane Performance Analysis Dashboard
# ==============================================================================

import operator
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime

# ------------------------------------------------------------------------------
# Page config
# ------------------------------------------------------------------------------
from PIL import Image

icon = Image.open("veolia.png")

st.set_page_config(
    page_title="DSSRO Membrane Performance Dashboard",
    page_icon="icon",
    layout="wide",
)

st.title("🔬 DSSRO Membrane Performance Analysis Dashboard")
st.markdown(
    """
    Upload your RO/DSSRO operational data (Excel) and this dashboard will:
    1. Clean & rename raw sensor columns
    2. Remove non-physical (NaN / negative) readings
    3. Isolate truly **operational** membrane conditions (filter out flush/CIP cycles)
    4. Compute normalized performance indicators (SSPn, QSPn, DPn, WTCn, STCn, ASPn)
    5. Plot each indicator against the Recovery Setpoint over time
    """
)

# ==============================================================================
# SIDEBAR — Inputs
# ==============================================================================
st.sidebar.header("1️⃣ Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

with st.sidebar.expander("ℹ️ Expected raw column names"):
    st.write(
        """
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
    )

st.sidebar.header("2️⃣ Constants & Conversions")
with st.sidebar.expander("Unit Conversions"):
    CONC_PPM_CONV = st.number_input("uS/cm → ppm factor", value=2.143, format="%.4f")
    M3H_TO_GPM = st.number_input("m3/h → gpm factor", value=4.40286, format="%.5f")
    BAR_TO_PSI = st.number_input("bar → psi factor", value=14.5, format="%.2f")

with st.sidebar.expander("System Constants"):
    EPV = st.number_input("Elements per Vessel (EPV)", value=5.00)
    V = st.number_input("Number of Vessels (V)", value=1.00)
    EMAe = st.number_input("Element Membrane Area (sq ft)", value=90.00)

with st.sidebar.expander("Standard Test Conditions"):
    Qpe = st.number_input("Standard Permeate Flow (GPD)", value=2600.00)
    Ke = st.number_input("Constant Ke", value=2700.00)
    NDPe = st.number_input("Standard NDP (psi)", value=225.00)
    Cfe = st.number_input("Standard Feed Concentration (ppm)", value=2000.00)

with st.sidebar.expander("Temperature Assumptions"):
    TR = st.number_input("Room Temperature (K)", value=298.15)
    TS = st.number_input("Standard Temperature (K)", value=273.15)

st.sidebar.header("3️⃣ Filtering Conditions")
apply_cond1 = st.sidebar.checkbox(
    "Cond. 1: Pf > 3x(next Pf) AND Pf > 6 bar  (flush/CIP detection)", value=True
)
apply_cond2 = st.sidebar.checkbox("Cond. 2: Pp > 0", value=True)
apply_cond3 = st.sidebar.checkbox("Cond. 3: Qc >= 1 m3/h", value=True)

st.sidebar.markdown("**Optional custom condition**")
add_condition_enabled = st.sidebar.checkbox("Enable custom condition", value=False)
custom_col = custom_op = None
custom_val = 0.0
_expected_cols = [
    'T (C)', 'pH', 'Cf (uS/cm)', 'Cp (uS/cm)', 'Qp (m3/h)',
    'Qc (m3/h)', 'Pf (bar)', 'Pc (bar)', 'Pp (bar)', 'Rsp (%)',
]
if add_condition_enabled:
    custom_col = st.sidebar.selectbox("Column", _expected_cols)
    custom_op = st.sidebar.selectbox("Operator", [">", ">=", "<", "<=", "==", "!="])
    custom_val = st.sidebar.number_input("Threshold value", value=0.0)

st.sidebar.header("4️⃣ Plot Y-Axis Ranges")
with st.sidebar.expander("Adjust plot ranges"):
    ylim_sspn = st.slider("SSPn (%)", 0.0, 20.0, (0.0, 5.0))
    ylim_qspn = st.slider("QSPn (gpm)", 0.0, 20.0, (0.0, 7.0))
    ylim_dpn = st.slider("DPn (psi)", 0.0, 100.0, (0.0, 35.0))
    ylim_wtcn = st.slider("WTCn (m/s-kPa) x1e-8", 0.0, 5.0, (0.0, 1.2))
    ylim_stcn = st.slider("STCn (m/s) x1e-7", 0.0, 10.0, (0.0, 2.5))
    ylim_aspn = st.slider("ASPn (%)", 0.0, 20.0, (0.0, 5.0))

if uploaded_file is None:
    st.info("👈 Upload an Excel file in the sidebar to begin.")
    st.stop()

# ==============================================================================
# STEP 1 — LOAD & CLEAN RAW DATA
# ==============================================================================
try:
    raw_df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Could not read Excel file: {e}")
    st.stop()

final_rename_map = {
    'Timestamp (yyyy-MM-dd HH:mm:ss)': 'Date',
    'RO-Mixed feed Temperature-°F': 'T (C)',
    'RO-Permeate flowrate-m³/hr': 'Qp (m3/h)',
    'RO-Concentrate Recircualtion Flow-m³/hr': 'Qc (m3/h)',
    'RO-Concentrate pressure-bar': 'Pc (bar)',
    'RO-Element feed Pressure-bar': 'Pf (bar)',
    'RO-Permeate line pressure-bar': 'Pp (bar)',
    'RO-Recovery Actual Setpoint-%': 'Rsp (%)',
    'RO-Raw feed pH-pH': 'pH',
    'RO-Raw feed Conductivity-mS/cm': 'Cf (uS/cm)',
    'RO-Permeate Conductivity-mS/cm': 'Cp (uS/cm)',
}
raw_df.rename(columns=final_rename_map, inplace=True)

expected_final_columns_for_analysis = [
    'Date', 'T (C)', 'pH', 'Cf (uS/cm)', 'Cp (uS/cm)', 'Qp (m3/h)',
    'Qc (m3/h)', 'Pf (bar)', 'Pc (bar)', 'Pp (bar)', 'Rsp (%)',
]
missing = [c for c in expected_final_columns_for_analysis if c not in raw_df.columns]
if missing:
    st.warning(f"⚠️ Missing expected columns (will be NaN): {missing}")

raw_df = raw_df.reindex(columns=expected_final_columns_for_analysis)
raw_df['Cf (uS/cm)'] *= 1000
raw_df['Date'] = pd.to_datetime(raw_df['Date'], errors='coerce')

st.header("Step 1 — Raw Data")
with st.expander("Show raw (renamed) data", expanded=False):
    st.dataframe(raw_df, use_container_width=True)
    st.caption(f"Rows: {len(raw_df)}")

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

if len(org_df) == 0:
    st.error("No data remains after cleaning. Check your file / units.")
    st.stop()

# ==============================================================================
# STEP 2 — FILTER FOR OPERATIONAL CONDITIONS
# ==============================================================================
st.header("Step 2 — Operational Filtering")

condition1 = (org_df['Pf (bar)'] > 3 * org_df['Pf (bar)'].shift(-1)) & (org_df['Pf (bar)'] > 6)
condition2 = (org_df['Pp (bar)'] > 0)
condition3 = (org_df['Qc (m3/h)'] >= 1)

combined_condition = pd.Series(True, index=org_df.index)
if apply_cond1:
    combined_condition &= condition1
if apply_cond2:
    combined_condition &= condition2
if apply_cond3:
    combined_condition &= condition3

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
# STEP 3 — CALCULATED METRICS
# ==============================================================================
st.header("Step 3 — Calculated Performance Metrics")

calc_df = fil_df[['Date', 'T (C)']].copy()

calc_df['Rec (%)'] = fil_df['Qp (m3/h)'] / (fil_df['Qp (m3/h)'] + fil_df['Qc (m3/h)'])
calc_df['DP (bar)'] = fil_df['Pf (bar)'] - fil_df['Pc (bar)']

calc_df['Cfo (ppm)'] = fil_df['Cf (uS/cm)'] / CONC_PPM_CONV
calc_df['Cpo (ppm)'] = fil_df['Cp (uS/cm)'] / CONC_PPM_CONV

calc_df['Qco (gpm)'] = fil_df['Qc (m3/h)'] * M3H_TO_GPM
calc_df['Qpo (gpm)'] = fil_df['Qp (m3/h)'] * M3H_TO_GPM

calc_df['Pfo (psi)'] = fil_df['Pf (bar)'] * BAR_TO_PSI
calc_df['Pco (psi)'] = fil_df['Pc (bar)'] * BAR_TO_PSI
calc_df['Ppo (psi)'] = fil_df['Pp (bar)'] * BAR_TO_PSI

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
# STEP 4 — REFERENCE POINTS & NORMALIZED METRICS
# ==============================================================================
st.header("Step 4 — Normalized Metrics")

Qpr = fil_df['Qp (m3/h)'].iloc[0]
TCFr = calc_df['TCF (no units)'].iloc[0]
NDPr = calc_df['NDP (psi)'].iloc[0]
Qcr = fil_df['Qc (m3/h)'].iloc[0]

st.caption(
    f"Reference point (first operational row): "
    f"Qpr = {Qpr:.3f} m3/h, TCFr = {TCFr:.4f}, NDPr = {NDPr:.2f} psi, Qcr = {Qcr:.3f} m3/h"
)

norm_df = pd.DataFrame({'Date': fil_df['Date']}).copy()

norm_df['SSPn (%)'] = (
    calc_df['SSP (%)'] * (fil_df['Qp (m3/h)'] / Qpr) * (TCFr / calc_df['TCF (no units)'])
)

norm_df['QSPn (gpm)'] = (
    fil_df['Qp (m3/h)'] * NDPr / calc_df['NDP (psi)'] * (TCFr / calc_df['TCF (no units)'])
) * M3H_TO_GPM

norm_df['DPn (psi)'] = (
    (
        calc_df['DP (bar)']
        * (Qpr / 2 + Qcr) ** 1.4
        / ((fil_df['Qp (m3/h)'] / 2 + fil_df['Qc (m3/h)'])) ** 1.4
    )
    * (1 + 0.01 * (calc_df['T (C)'] - 25))
) * BAR_TO_PSI

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
# STEP 5 — PLOTS
# ==============================================================================
st.header("Step 5 — Diagnostic Plots")


def plot_normalized_data(data_df, fil_df, norm_col, norm_label, title, norm_ylim):
    fig, ax1 = plt.subplots(figsize=(11, 5.5))

    ax1.scatter(data_df['Date'], data_df[norm_col], marker='D', label=norm_label, color='darkblue')
    ax1.set_xlabel('Date')
    ax1.set_ylabel(norm_label)
    ax1.set_title(title)
    ax1.set_ylim(norm_ylim)
    ax1.grid()

    min_date, max_date = data_df['Date'].min(), data_df['Date'].max()
    time_span = max_date - min_date

    if time_span < datetime.timedelta(days=90):
        locator = mdates.DayLocator(interval=7)
        formatter = mdates.DateFormatter('%Y-%m-%d')
    elif time_span < datetime.timedelta(days=365 * 2):
        locator = mdates.MonthLocator(interval=1)
        formatter = mdates.DateFormatter('%Y-%m')
    elif time_span < datetime.timedelta(days=365 * 5):
        locator = mdates.MonthLocator(interval=6)
        formatter = mdates.DateFormatter('%Y-%m')
    else:
        locator = mdates.YearLocator(interval=1)
        formatter = mdates.DateFormatter('%Y')

    ax1.xaxis.set_major_locator(locator)
    ax1.xaxis.set_major_formatter(formatter)
    ax1.tick_params(axis='x', rotation=45)
    ax1.minorticks_on()

    ax2 = ax1.twinx()
    ax2.scatter(fil_df['Date'], fil_df['Rsp (%)'], color='red', marker='s', label='Rsp (%)')
    ax2.set_ylabel('R_sp (%)')
    ax2.set_ylim(0, 100)
    ax2.minorticks_on()

    lines, labels = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax2.legend(lines + lines2, labels + labels2, loc='upper left')

    fig.tight_layout()
    return fig


def show_plot_with_download(fig, filename):
    st.pyplot(fig)
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    st.download_button(
        f"⬇️ Download '{filename}'",
        data=buf.getvalue(),
        file_name=filename,
        mime="image/png",
    )
    plt.close(fig)


plots_config = [
    (norm_df, 'SSPn (%)', 'SSPn (%)',
     'Normalized % Salt Passage (SSPn) vs Recovery Setpoint',
     'SSPn_and_R_sp_Over_Time.png', ylim_sspn),
    (norm_df, 'QSPn (gpm)', 'Permeate Flow Rate (gpm)',
     'Permeate Flow Rate (QSPn) vs Recovery Setpoint',
     'QSPn_and_R_sp_Over_Time.png', ylim_qspn),
    (norm_df, 'DPn (psi)', 'DPn (psi)',
     'Normalized Differential Pressure (DPn) vs Recovery Setpoint',
     'DPn_and_R_sp_Over_Time.png', ylim_dpn),
    (norm_df, 'WTCn (m/s-kPa)', 'WTCn (m/s-kPa)',
     'Water Transport Coefficient (WTCn) vs Recovery Setpoint',
     'WTCn_and_R_sp_Over_Time.png', tuple(v * 1e-8 for v in ylim_wtcn)),
    (norm_df, 'STCn (m/s)', 'STCn (m/s)',
     'Salt Transport Coefficient (STCn) vs Recovery Setpoint',
     'STCn_and_R_sp_Over_Time.png', tuple(v * 1e-7 for v in ylim_stcn)),
    (calc_df, 'ASPn (%)', 'ASPn (%)',
     '% Salt Passage (ASPn) vs Recovery Setpoint',
     'ASPn_and_R_sp_Over_Time.png', ylim_aspn),
]

tab_labels = ["SSPn", "QSPn", "DPn", "WTCn", "STCn", "ASPn"]
tabs = st.tabs(tab_labels)

for tab, (df_source, col, label, title, fname, ylim) in zip(tabs, plots_config):
    with tab:
        fig = plot_normalized_data(df_source, fil_df, col, label, title, ylim)
        show_plot_with_download(fig, fname)

st.success("✅ Analysis complete. Adjust settings in the sidebar and the dashboard updates automatically.")

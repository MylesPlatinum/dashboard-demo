import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from datetime import datetime
import yaml

# === LOAD CONFIG ===
@st.cache_resource
def load_config():
    try:
        with open('config.yaml', 'r') as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        st.error("config.yaml not found!")
        st.stop()

config = load_config()

st.set_page_config(page_title=config['dashboard']['title'], layout="wide", initial_sidebar_state="expanded")

# === PASSWORD ===
if "auth" not in st.session_state:
    st.session_state.auth = False

PASSWORD = config['dashboard']['password']
if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        try:
            from PIL import Image
            if Path(config['branding']['logo_file']).exists():
                st.image(config['branding']['logo_file'], width=150)
        except:
            st.markdown(
                f"<div style='text-align: center; font-size: 60px; margin: 20px; color: {config['branding']['primary_color']};'>Office Building</div>",
                unsafe_allow_html=True
            )
        st.markdown(f"<h1 style='text-align: center;'>{config['client']['name']}</h1>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #888;'>{config['dashboard']['subtitle']}</p>", unsafe_allow_html=True)
        pwd = st.text_input("Password", type="password")
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        elif pwd:
            st.error("Wrong password")
    st.stop()

# === DEBUG MODE ===
debug_mode = st.sidebar.checkbox("Debug Mode", value=False)

# === FIND LATEST FILES ===
@st.cache_data(ttl=60)
def get_latest_files():
    data_dir = Path("data")
    if not data_dir.exists():
        st.error("No 'data' folder found!")
        st.stop()

    rev_files = list(data_dir.glob(config['data']['revenue_file_pattern']))
    cost_files = list(data_dir.glob(config['data']['costs_file_pattern']))

    if not rev_files: st.error("No revenue file found!"); st.stop()
    if not cost_files: st.error("No costs file found!"); st.stop()

    return max(rev_files, key=lambda x: x.stat().st_mtime), max(cost_files, key=lambda x: x.stat().st_mtime)

# === LOAD DATA (NOW BULLETPROOF) ===
@st.cache_data(ttl=60)
def load_data():
    revenue_path, costs_path = get_latest_files()
    branches = config['data']['branches']

    if debug_mode:
        st.sidebar.write(f"Revenue: {revenue_path.name}")
        st.sidebar.write(f"Costs: {costs_path.name}")

    # AUTO-DETECT SHEET NAME (this fixes the "Sheet1 not found" error forever)
    with pd.ExcelFile(revenue_path) as xls:
        sheet_name = xls.sheet_names[0]  # always use first sheet
        revenue_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    # Your existing row numbers from config
    rev_start = config['data']['revenue_start_row']
    rev_end   = config['data']['revenue_end_row']
    hrs_start = config['data']['hours_start_row']
    hrs_end   = config['data']['hours_end_row']

    revenue_section = revenue_raw.iloc[rev_start:rev_end+1].copy()
    hours_section   = revenue_raw.iloc[hrs_start:hrs_end+1].copy()

    # Build data
    data_list = []
    for i, period in enumerate(revenue_section.iloc[:, 1].astype(str)):
        if period.strip() and period != 'nan':
            for j, branch in enumerate(branches):
                col_idx = j + 3
                rev = pd.to_numeric(revenue_section.iloc[i, col_idx], errors='coerce') or 0
                hrs = pd.to_numeric(hours_section.iloc[i, col_idx], errors='coerce') or 0
                if rev > 0:
                    data_list.append({
                        'Period': period.strip(),
                        'Branch': branch,
                        'Revenue': rev,
                        'Hours': hrs
                    })

    df = pd.DataFrame(data_list)

    # COSTS — also auto-detect sheet + skip junk row
    costs_raw = pd.read_excel(costs_path, header=1)  # skips Description row
    costs_raw.columns = ['Period'] + branches + ['Total']
    costs_clean = costs_raw[costs_raw['Period'].notna()].copy()
    costs_clean['Period'] = costs_clean['Period'].astype(str).str.extract('(\d+)')[0]

    cost_melt = costs_clean.melt(id_vars='Period', value_vars=branches, var_name='Branch', value_name='Cost')
    cost_melt['Cost'] = pd.to_numeric(cost_melt['Cost'], errors='coerce').fillna(0)

    df = df.merge(cost_melt, on=['Period', 'Branch'], how='left').fillna({'Cost': 0})

    # Calculations
    df['Gross Profit'] = df['Revenue'] - df['Cost']
    df['Margin %'] = (df['Gross Profit'] / df['Revenue'] * 100).round(1).fillna(0)
    df['Rev per Hour'] = (df['Revenue'] / df['Hours'].replace(0, 1)).round(2)
    df['Period_Int'] = df['Period'].astype(int)
    df = df.sort_values(['Period_Int', 'Branch'])

    return df, branches, revenue_path.name, costs_path.name

df, branches, rev_file, cost_file = load_data()

# === UI ===
col1, col2 = st.columns([1, 5])
with col1:
    try:
        from PIL import Image
        if Path(config['branding']['logo_file']).exists():
            st.image(config['branding']['logo_file'], width=80)
    except: pass
with col2:
    st.title(f"{config['client']['name']} – Dashboard")
    st.markdown(f"**Files:** `{rev_file}` | `{cost_file}` • {datetime.now():%d %b %Y, %H:%M}")

# Filters
all_periods = sorted(df['Period'].unique(), key=int)
sel_periods = st.sidebar.multiselect("Periods", all_periods, default=all_periods[-3:])
sel_branches = st.sidebar.multiselect("Branches", branches, default=branches)

filtered = df[df['Period'].isin(sel_periods) & df['Branch'].isin(sel_branches)]

# KPIs
totals = filtered.agg({'Revenue': 'sum', 'Hours': 'sum', 'Cost': 'sum', 'Gross Profit': 'sum'})
avg_margin = filtered['Margin %'].mean()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Revenue", f"£{totals['Revenue']:,.0f}")
c2.metric("Hours", f"{totals['Hours']:,.0f}")
c3.metric("Costs", f"£{totals['Cost']:,.0f}")
c4.metric("Profit", f"£{totals['Gross Profit']:,.0f}")
c5.metric("Margin", f"{avg_margin:.1f}%")

# Charts
st.plotly_chart(px.line(filtered, x='Period', y='Revenue', color='Branch', title="Revenue Trend"), use_container_width=True)
st.plotly_chart(px.bar(filtered.groupby('Branch').agg({'Gross Profit': 'sum'}).reset_index(), x='Branch', y='Gross Profit', title="Profit by Branch"), use_container_width=True)

st.success("Dashboard loaded perfectly – all sheet name issues fixed!")
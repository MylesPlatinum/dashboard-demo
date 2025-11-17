import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib import colors as pdf_colors
from io import BytesIO
from datetime import datetime
import yaml

# === LOAD CLIENT CONFIGURATION ===
@st.cache_resource
def load_config():
    try:
        with open('config.yaml', 'r') as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        st.error("Configuration file not found! Please ensure config.yaml exists.")
        st.stop()
    except yaml.YAMLError as e:
        st.error(f"Error reading configuration: {e}")
        st.stop()

config = load_config()

# === PAGE CONFIG ===
st.set_page_config(
    page_title=config['dashboard']['title'],
    layout="wide",
    initial_sidebar_state="expanded"
)

# === PASSWORD PROTECTION ===
PASSWORD = config['dashboard']['password']
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        logo_loaded = False
        try:
            from PIL import Image
            logo_path = config['branding']['logo_file']
            if Path(logo_path).exists():
                logo = Image.open(logo_path)
                col_a, col_b, col_c = st.columns([1, 2, 1])
                with col_b:
                    st.image(logo, width=150)
                logo_loaded = True
        except:
            pass
       
        if not logo_loaded:
            st.markdown(
                f"<div style='text-align: center; font-size: 60px; margin: 20px; color: {config['branding']['primary_color']};'>Office Building</div>",
                unsafe_allow_html=True
            )
       
        st.markdown(f"<h1 style='text-align: center;'>{config['client']['name']}</h1>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #888;'>{config['dashboard']['subtitle']}</p>", unsafe_allow_html=True)
       
        pwd = st.text_input("Enter Password", type="password")
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        elif pwd:
            st.error("Wrong password")
       
        if config['company_values']['show_on_login']:
            st.divider()
            with st.expander("Our Values & Mission", expanded=False):
                st.markdown("### Our Values")
                for value in config['company_values']['values']:
                    st.markdown(f"**{value['title']}** - *{value['tagline']}*")
                    st.markdown(value['description'])
                    st.markdown("")
                st.markdown("---")
                st.markdown("### Mission Statement")
                st.markdown(f"#### *\"{config['company_values']['mission']['title']}\"*")
                st.markdown(config['company_values']['mission']['description'])
                for point in config['company_values']['mission']['points']:
                    st.markdown(f"**{point['title']}** - {point['text']}")
    st.stop()

# === DEBUG MODE ===
debug_mode = st.sidebar.checkbox("Debug Mode", value=config['features']['debug_mode'], help="Show detailed data loading info")

# === AUTO-DETECT LATEST FILES ===
@st.cache_data(ttl=60)
def get_latest_files():
    data_dir = Path("data")
    if not data_dir.exists():
        st.error("'data' folder not found! Please create it and add your Excel files.")
        st.stop()
   
    revenue_pattern = config['data']['revenue_file_pattern']
    costs_pattern = config['data']['costs_file_pattern']
   
    revenue_files = list(data_dir.glob(revenue_pattern))
    costs_files = list(data_dir.glob(costs_pattern))
   
    if not revenue_files:
        st.error(f"No revenue files found matching pattern: {revenue_pattern}")
        st.stop()
    if not costs_files:
        st.error(f"No costs files found matching pattern: {costs_pattern}")
        st.stop()
   
    latest_revenue = max(revenue_files, key=lambda f: f.stat().st_mtime)
    latest_costs = max(costs_files, key=lambda f: f.stat().st_mtime)
   
    return latest_revenue, latest_costs

# === LOAD DATA WITH CORRECT STRUCTURE ===
@st.cache_data(ttl=60)
def load_data():
    revenue_path, costs_path = get_latest_files()
   
    if debug_mode:
        st.sidebar.info(f"Loading:\n- {revenue_path.name}\n- {costs_path.name}")
   
    try:
        branches = config['data']['branches']
        revenue_start_row = config['data']['revenue_start_row']
        revenue_end_row = config['data']['revenue_end_row']
        hours_start_row = config['data']['hours_start_row']
        hours_end_row = config['data']['hours_end_row']
       
        revenue_raw = pd.read_excel(revenue_path, sheet_name=config['data']['revenue_sheet'], header=None)
       
        revenue_section = revenue_raw.iloc[revenue_start_row:revenue_end_row+1].copy()
        hours_section = revenue_raw.iloc[hours_start_row:hours_end_row+1].copy()
       
        col_mapping = {branch: idx + 3 for idx, branch in enumerate(branches)}
       
        periods = revenue_section.iloc[:, 1].astype(str).tolist()
        dates = revenue_section.iloc[:, 2].tolist()
       
        data_list = []
        for branch, col_idx in col_mapping.items():
            for i, period in enumerate(periods):
                if period and period.strip() and period != 'None':
                    rev_val = revenue_section.iloc[i, col_idx] if col_idx < len(revenue_section.columns) else 0
                    hrs_val = hours_section.iloc[i, col_idx] if i < len(hours_section) and col_idx < len(hours_section.columns) else 0
                   
                    rev_numeric = pd.to_numeric(rev_val, errors='coerce')
                    hrs_numeric = pd.to_numeric(hrs_val, errors='coerce')
                   
                    if pd.notna(rev_numeric) and rev_numeric > 0:
                        data_list.append({
                            'Period': period.strip(),
                            'Date Range': dates[i] if i < len(dates) else '',
                            'Branch': branch,
                            'Revenue': rev_numeric,
                            'Hours': hrs_numeric if pd.notna(hrs_numeric) else 0
                        })
       
        df = pd.DataFrame(data_list)
       
        # FIXED: header=1 to skip Description/Date Range row
        costs_raw = pd.read_excel(costs_path, header=1)
       
        costs_columns = ['Period'] + branches + ['Total']
        costs_raw.columns = costs_columns
       
        costs_clean = costs_raw[costs_raw['Period'].notna()].copy()
        costs_clean['Period'] = costs_clean['Period'].astype(str).str.extract('(\d+)', expand=False)
        costs_clean = costs_clean[costs_clean['Period'].notna()]
       
        cost_melt = costs_clean.melt(id_vars='Period', value_vars=branches, var_name='Branch', value_name='Cost')
        cost_melt['Cost'] = pd.to_numeric(cost_melt['Cost'], errors='coerce').fillna(0)
       
        df = df.merge(cost_melt, on=['Period', 'Branch'], how='left')
        df['Cost'] = df['Cost'].fillna(0)
       
        df['Gross Profit'] = df['Revenue'] - df['Cost']
        df['Margin %'] = df.apply(lambda r: round(r['Gross Profit'] / r['Revenue'] * 100, 1) if r['Revenue'] > 0 else 0, axis=1)
        df['Rev per Hour'] = df.apply(lambda r: round(r['Revenue'] / r['Hours'], 2) if r['Hours'] > 0 else 0, axis=1)
        df['Period_Int'] = df['Period'].astype(int)
        df = df.sort_values(['Period_Int', 'Branch'])
       
        care_f = pd.DataFrame()
        care_hours = pd.DataFrame()
       
        return df, care_f, care_hours, branches, revenue_path.name, costs_path.name
       
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        st.stop()

df, care_f, care_hours, branches, rev_file, cost_file = load_data()

# === HEADER ===
col1, col2 = st.columns([1, 5])
with col1:
    try:
        from PIL import Image
        logo_path = config['branding']['logo_file']
        if Path(logo_path).exists():
            logo = Image.open(logo_path)
            st.image(logo, width=80)
    except:
        pass
with col2:
    st.title(f"{config['client']['name']} – Interactive Dashboard")
    st.markdown(f"**Data Sources:** `{rev_file}` | `{cost_file}` | **Last Updated:** {datetime.now():%d %b %Y, %H:%M}")
st.divider()

# === SIDEBAR FILTERS ===
st.sidebar.header("Filters & Controls")
if st.sidebar.button("Refresh Data Now"):
    st.cache_data.clear()
    st.rerun()

all_periods = sorted(df["Period"].unique(), key=lambda x: int(x))
period_option = st.sidebar.radio("Period Selection", ["All Periods", "Latest Only", "Latest 3"])
if period_option == "All Periods":
    sel_periods = all_periods
elif period_option == "Latest Only":
    sel_periods = [all_periods[-1]]
else:
    sel_periods = all_periods[-3:]

select_all_branches = st.sidebar.checkbox("Select All Branches", value=True)
sel_branches = branches if select_all_branches else st.sidebar.multiselect("Choose Branches", branches, default=branches)

chart_height = st.sidebar.slider("Chart Height", 300, 800, 450)
show_markers = st.sidebar.checkbox("Show Data Points on Lines", value=True)

filtered_df = df[df['Period'].isin(sel_periods) & df['Branch'].isin(sel_branches)].copy()

branch_totals = filtered_df.groupby('Branch').agg({
    'Revenue': 'sum', 'Hours': 'sum', 'Cost': 'sum', 'Gross Profit': 'sum'
}).reset_index()
branch_totals['Margin %'] = (branch_totals['Gross Profit'] / branch_totals['Revenue'] * 100).round(1)

# === KPIs ===
st.header("Key Performance Indicators")
c1, c2, c3, c4, c5 = st.columns(5)
total_revenue = filtered_df['Revenue'].sum()
total_hours = filtered_df['Hours'].sum()
total_cost = filtered_df['Cost'].sum()
total_profit = filtered_df['Gross Profit'].sum()
avg_margin = filtered_df['Margin %'].mean()
c1.metric("Total Revenue", f"£{total_revenue:,.0f}")
c2.metric("Total Hours", f"{total_hours:,.0f}")
c3.metric("Total Costs", f"£{total_cost:,.0f}")
c4.metric("Gross Profit", f"£{total_profit:,.0f}")
c5.metric("Avg Margin", f"{avg_margin:.1f}%")

st.success("Dashboard fully loaded – all fixes applied!")
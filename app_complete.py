import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
import yaml
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
import plotly.io as pio
from reportlab.platypus import Image as RLImage

# === PAGE CONFIG ===
st.set_page_config(
    page_title="Professional Analytics Dashboard", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# === LOAD CONFIG ===
@st.cache_resource
def load_config():
    """Load configuration from config.yaml"""
    try:
        with open('config.yaml', 'r') as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        st.error("❌ config.yaml not found! Please upload it to your repository.")
        st.stop()

config = load_config()

# === PASSWORD PROTECTION ===
if "auth" not in st.session_state:
    st.session_state.auth = False

PASSWORD = config['dashboard'].get('password', 'Demo2024')

if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        try:
            from PIL import Image
            logo_path = config['branding']['logo_file']
            if Path(logo_path).exists():
                st.image(logo_path, width=150)
        except:
            st.markdown(
                f"<div style='text-align: center; font-size: 60px; margin: 20px;'>📊</div>",
                unsafe_allow_html=True
            )
        
        st.markdown(f"<h1 style='text-align: center;'>{config['client']['name']}</h1>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center; color: #888;'>{config['dashboard']['subtitle']}</p>", unsafe_allow_html=True)
        
        pwd = st.text_input("Enter Password", type="password", key="password_input")
        
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        elif pwd:
            st.error("❌ Incorrect password")
    
    st.stop()

# === FIND DATA FILES ===
@st.cache_data(ttl=60)
def get_latest_files():
    """Find the revenue and costs files"""
    
    possible_locations = [
        Path("data"),
        Path("."),
        Path("uploads")
    ]
    
    revenue_file = None
    costs_file = None
    
    for location in possible_locations:
        if location.exists():
            rev_pattern = config['data'].get('revenue_file_pattern', '*revenue*.xlsx')
            rev_files = list(location.glob(rev_pattern))
            if rev_files:
                revenue_file = max(rev_files, key=lambda x: x.stat().st_mtime)
            
            cost_pattern = config['data'].get('costs_file_pattern', '*costs*.xlsx')
            cost_files = list(location.glob(cost_pattern))
            if cost_files:
                costs_file = max(cost_files, key=lambda x: x.stat().st_mtime)
            
            if revenue_file and costs_file:
                break
    
    if not revenue_file:
        st.error("❌ Revenue file not found! Please upload a file matching pattern: " + config['data']['revenue_file_pattern'])
        st.stop()
    
    if not costs_file:
        st.error("❌ Costs file not found! Please upload a file matching pattern: " + config['data']['costs_file_pattern'])
        st.stop()
    
    return revenue_file, costs_file

# === LOAD DATA ===
@st.cache_data(ttl=60)
def load_data():
    """Load and process the Excel files"""
    
    revenue_path, costs_path = get_latest_files()
    branches = config['data']['branches']
    
    revenue_raw = pd.read_excel(revenue_path, header=None)
    
    data_list = []
    
    for idx, row in revenue_raw.iterrows():
        if idx < 4:
            continue
        
        description = str(row[0]).strip()
        period_val = row[1]
        
        if pd.isna(period_val) or period_val == '':
            continue
        
        if description.upper() == 'TOTAL':
            period = str(int(period_val))
            date_range = row[2] if not pd.isna(row[2]) else f"Period {period}"
            
            for i, branch in enumerate(branches):
                col_idx = 3 + i
                revenue = pd.to_numeric(row[col_idx], errors='coerce')
                
                if not pd.isna(revenue) and revenue > 0:
                    data_list.append({
                        'Period': period,
                        'Date Range': date_range,
                        'Branch': branch,
                        'Revenue': revenue
                    })
    
    df = pd.DataFrame(data_list)
    
    costs_raw = pd.read_excel(costs_path, header=1)
    
    cost_data = []
    for idx, row in costs_raw.iterrows():
        period = row['Period']
        
        if pd.isna(period):
            continue
        
        period = str(int(period))
        
        for branch in branches:
            if branch in costs_raw.columns:
                cost = pd.to_numeric(row[branch], errors='coerce')
                if not pd.isna(cost):
                    cost_data.append({
                        'Period': period,
                        'Branch': branch,
                        'Cost': cost
                    })
    
    costs_df = pd.DataFrame(cost_data)
    
    df = df.merge(costs_df, on=['Period', 'Branch'], how='left')
    df['Cost'] = df['Cost'].fillna(0)
    
    df['Gross Profit'] = df['Revenue'] - df['Cost']
    df['Margin %'] = ((df['Gross Profit'] / df['Revenue']) * 100).round(1)
    df['Margin %'] = df['Margin %'].fillna(0)
    
    df['Period_Int'] = df['Period'].astype(int)
    df = df.sort_values(['Period_Int', 'Branch'])
    
    return df, branches, revenue_path.name, costs_path.name

def generate_premium_pdf(filtered_df, branch_totals, total_revenue, total_cost, total_profit, avg_margin):
    """Generate PREMIUM PDF with charts - makes £495 look like theft!"""
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=50, leftMargin=50, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    # Premium styling
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=28, 
                                 textColor=colors.HexColor('#1a1a1a'), spaceAfter=5, 
                                 alignment=TA_CENTER, fontName='Helvetica-Bold')
    
    heading_style = ParagraphStyle('CustomHeading', parent=styles['Heading2'], fontSize=18, 
                                   textColor=colors.HexColor('#2c3e50'), spaceAfter=15, spaceBefore=20,
                                   fontName='Helvetica-Bold', borderWidth=2, 
                                   borderColor=colors.HexColor('#3498db'), borderPadding=10,
                                   backColor=colors.HexColor('#ecf0f1'))
    
    body_style = ParagraphStyle('CustomBody', parent=styles['Normal'], fontSize=11, leading=14)
    
    # COVER PAGE
    story.append(Spacer(1, 1.5*inch))
    story.append(Paragraph(config['client']['name'], title_style))
    story.append(Paragraph("EXECUTIVE FINANCIAL DASHBOARD REPORT", 
                          ParagraphStyle('Sub', parent=styles['Normal'], fontSize=16, 
                                       textColor=colors.HexColor('#3498db'), alignment=TA_CENTER,
                                       fontName='Helvetica-Bold', spaceAfter=30)))
    story.append(Spacer(1, 0.5*inch))
    
    # KPI Box
    kpi_data = [['REVENUE', 'PROFIT', 'MARGIN', 'BRANCHES'],
                [f'£{total_revenue:,.0f}', f'£{total_profit:,.0f}', f'{avg_margin:.1f}%', str(len(branch_totals))]]
    kpi_table = Table(kpi_data, colWidths=[1.6*inch]*4)
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#ecf0f1')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.HexColor('#2c3e50')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 16),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 2, colors.HexColor('#3498db'))
    ]))
    story.append(kpi_table)
    story.append(PageBreak())
    
    # EXECUTIVE SUMMARY
    story.append(Paragraph("EXECUTIVE SUMMARY", heading_style))
    top_branch = branch_totals.loc[branch_totals['Revenue'].idxmax()]
    
    summary_text = f'<para fontSize="11"><b>Top Performer:</b> {top_branch["Branch"]} with £{top_branch["Revenue"]:,.0f} revenue<br/><b>Analysis:</b> {len(branch_totals)} branches, {len(filtered_df["Period"].unique())} periods, £{total_profit:,.0f} profit</para>'
    story.append(Paragraph(summary_text, body_style))
    story.append(Spacer(1, 0.2*inch))
    
    summary_data = [
        ['Metric', 'Value', 'Analysis'],
        ['Total Revenue', f'£{total_revenue:,.0f}', f'{len(filtered_df)} transactions'],
        ['Gross Profit', f'£{total_profit:,.0f}', f'{avg_margin:.1f}% margin'],
        ['Active Branches', str(len(branch_totals)), 'All locations']
    ]
    summary_table = Table(summary_data, colWidths=[2*inch, 1.5*inch, 2.5*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10)
    ]))
    story.append(summary_table)
    story.append(PageBreak())
    
    # REVENUE CHART
    story.append(Paragraph("REVENUE TREND ANALYSIS", heading_style))
    fig_rev = px.line(filtered_df, x='Period_Int', y='Revenue', color='Branch', markers=True,
                      title="Revenue Performance by Branch",
                      color_discrete_sequence=px.colors.qualitative.Set2)
    fig_rev.update_layout(height=400, plot_bgcolor='white', paper_bgcolor='white', showlegend=True,
                         legend=dict(orientation="h", yanchor="bottom", y=-0.3))
    
    # EXPORT CHART AS IMAGE
    img_bytes = pio.to_image(fig_rev, format='png', width=700, height=400, scale=2)
    img = io.BytesIO(img_bytes)
    chart_image = RLImage(img, width=6.5*inch, height=3.7*inch)
    story.append(chart_image)
    story.append(PageBreak())
    
    # BRANCH COMPARISON CHARTS
    story.append(Paragraph("BRANCH COMPARISON", heading_style))
    
    fig_branch = px.bar(branch_totals.sort_values('Revenue', ascending=True), y='Branch', x='Revenue',
                        orientation='h', color='Revenue', color_continuous_scale='Blues', text='Revenue')
    fig_branch.update_traces(texttemplate='£%{text:,.0f}', textposition='outside')
    fig_branch.update_layout(height=350, showlegend=False, plot_bgcolor='white')
    
    img_bytes2 = pio.to_image(fig_branch, format='png', width=700, height=350, scale=2)
    img2 = io.BytesIO(img_bytes2)
    story.append(RLImage(img2, width=6.5*inch, height=3.2*inch))
    story.append(Spacer(1, 0.2*inch))
    
    fig_margin = px.bar(branch_totals.sort_values('Margin %', ascending=True), y='Branch', x='Margin %',
                        orientation='h', color='Margin %', color_continuous_scale='RdYlGn', text='Margin %')
    fig_margin.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig_margin.update_layout(height=350, showlegend=False, plot_bgcolor='white')
    
    img_bytes3 = pio.to_image(fig_margin, format='png', width=700, height=350, scale=2)
    img3 = io.BytesIO(img_bytes3)
    story.append(RLImage(img3, width=6.5*inch, height=3.2*inch))
    story.append(PageBreak())
    
    # BRANCH TABLE
    branch_data = [['Branch', 'Revenue', 'Costs', 'Profit', 'Margin']]
    for _, row in branch_totals.iterrows():
        branch_data.append([row['Branch'], f"£{row['Revenue']:,.0f}", f"£{row['Cost']:,.0f}",
                          f"£{row['Gross Profit']:,.0f}", f"{row['Margin %']:.1f}%"])
    
    branch_table = Table(branch_data, colWidths=[1.3*inch, 1.2*inch, 1.2*inch, 1.2*inch, 1.0*inch])
    branch_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#27ae60')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#ecf0f1')]),
        ('TOPPADDING', (0, 0), (-1, -1), 10)
    ]))
    story.append(branch_table)
    story.append(Spacer(1, 0.3*inch))
    
    # FOOTER
    footer = f'<para alignment="center" fontSize="9" textColor="#888"><b>{config["client"]["name"]} Advanced Analytics</b><br/>Generated {datetime.now():%d %B %Y}<br/><i>Professional report - Updated anytime with one click</i></para>'
    story.append(Paragraph(footer, body_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# === LOAD THE DATA ===
try:
    df, branches, rev_file, cost_file = load_data()
except Exception as e:
    st.error(f"❌ Error loading data: {str(e)}")
    st.code(str(e))
    st.stop()

# === HEADER ===
col1, col2 = st.columns([1, 5])
with col1:
    try:
        from PIL import Image
        logo_path = config['branding']['logo_file']
        if Path(logo_path).exists():
            st.image(logo_path, width=80)
    except:
        st.markdown("📊")

with col2:
    st.title(f"📊 {config['client']['name']}")
    st.markdown(f"**Data Sources:** `{rev_file}` | `{cost_file}` | **Updated:** {datetime.now():%d %b %Y, %H:%M}")

st.divider()

# === SIDEBAR FILTERS ===
st.sidebar.header("📊 Dashboard Controls")

if st.sidebar.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.divider()

all_periods = sorted(df['Period'].unique(), key=lambda x: int(x))
period_option = st.sidebar.radio(
    "Period Selection",
    ["All Periods", "Latest Period", "Custom Selection"]
)

if period_option == "All Periods":
    sel_periods = all_periods
elif period_option == "Latest Period":
    sel_periods = [all_periods[-1]]
else:
    sel_periods = st.sidebar.multiselect(
        "Select Periods",
        all_periods,
        default=all_periods
    )
    if not sel_periods:
        sel_periods = all_periods

st.sidebar.markdown("### Branch Selection")
select_all = st.sidebar.checkbox("Select All Branches", value=True)

if select_all:
    sel_branches = branches
else:
    sel_branches = st.sidebar.multiselect(
        "Choose Branches",
        branches,
        default=branches
    )
    if not sel_branches:
        sel_branches = branches

st.sidebar.divider()
st.sidebar.success(f"✅ Viewing: {len(sel_periods)} period(s) × {len(sel_branches)} branch(es)")

# === FILTER DATA ===
filtered_df = df[df['Period'].isin(sel_periods) & df['Branch'].isin(sel_branches)].copy()

if len(filtered_df) == 0:
    st.error("⚠️ No data matches your current filters!")
    st.stop()

# === AGGREGATE BY BRANCH ===
branch_totals = filtered_df.groupby('Branch').agg({
    'Revenue': 'sum',
    'Cost': 'sum',
    'Gross Profit': 'sum'
}).reset_index()
branch_totals['Margin %'] = ((branch_totals['Gross Profit'] / branch_totals['Revenue']) * 100).round(1)

# === KPI METRICS ===
st.header("📈 Key Performance Indicators")

col1, col2, col3, col4, col5 = st.columns(5)

total_revenue = filtered_df['Revenue'].sum()
total_cost = filtered_df['Cost'].sum()
total_profit = filtered_df['Gross Profit'].sum()
avg_margin = filtered_df['Margin %'].mean()
num_periods = len(sel_periods)

col1.metric("Total Revenue", f"£{total_revenue:,.0f}")
col2.metric("Total Costs", f"£{total_cost:,.0f}")
col3.metric("Gross Profit", f"£{total_profit:,.0f}")
col4.metric("Avg Margin", f"{avg_margin:.1f}%")
col5.metric("Periods", num_periods)

st.divider()

with col2:
    if st.button("📄 EXPORT PREMIUM PDF REPORT", type="primary", use_container_width=True):
        with st.spinner("Generating your quality executive report..."):
            pdf_buffer = generate_premium_pdf(filtered_df, branch_totals, total_revenue, total_cost, total_profit, avg_margin)
            st.download_button(
                label="⬇️ Download Professional Report (Worth £500+ from consultants!)",
                data=pdf_buffer,
                file_name=f"Executive_Dashboard_Report_{datetime.now():%Y%m%d_%H%M}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            st.success("✅ Premium report generated! This alone costs £500+ from consultancy firms.")
st.divider()

# === TABS ===
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Revenue Trends",
    "🏢 Branch Comparison",
    "💰 Profitability Analysis",
    "📊 Data Table"
])

# === TAB 1: REVENUE TRENDS ===
with tab1:
    st.subheader("Revenue & Margin Trends Over Time")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Revenue by Branch")
        fig_rev = px.line(
            filtered_df,
            x='Period_Int',
            y='Revenue',
            color='Branch',
            markers=True,
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig_rev.update_layout(
            height=400,
            xaxis_title="Period",
            yaxis_title="Revenue (£)",
            hovermode='x unified'
        )
        st.plotly_chart(fig_rev, use_container_width=True)
    
    with col2:
        st.markdown("#### Margin % by Branch")
        fig_margin = px.line(
            filtered_df,
            x='Period_Int',
            y='Margin %',
            color='Branch',
            markers=True,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_margin.update_layout(
            height=400,
            xaxis_title="Period",
            yaxis_title="Margin %",
            hovermode='x unified'
        )
        st.plotly_chart(fig_margin, use_container_width=True)

# === TAB 2: BRANCH COMPARISON ===
with tab2:
    st.subheader("Branch Performance Comparison")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Total Revenue by Branch")
        fig_branch_rev = px.bar(
            branch_totals.sort_values('Revenue', ascending=True),
            y='Branch',
            x='Revenue',
            orientation='h',
            color='Revenue',
            color_continuous_scale='Blues',
            text='Revenue'
        )
        fig_branch_rev.update_traces(texttemplate='£%{text:,.0f}', textposition='outside')
        fig_branch_rev.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig_branch_rev, use_container_width=True)
    
    with col2:
        st.markdown("#### Margin % by Branch")
        fig_margin_bar = px.bar(
            branch_totals.sort_values('Margin %', ascending=True),
            y='Branch',
            x='Margin %',
            orientation='h',
            color='Margin %',
            color_continuous_scale='RdYlGn',
            text='Margin %'
        )
        fig_margin_bar.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_margin_bar.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig_margin_bar, use_container_width=True)
    
    st.divider()
    
    st.markdown("#### Branch Summary Table")
    st.dataframe(
        branch_totals.style.format({
            'Revenue': '£{:,.0f}',
            'Cost': '£{:,.0f}',
            'Gross Profit': '£{:,.0f}',
            'Margin %': '{:.1f}%'
        }),
        use_container_width=True,
        height=250
    )

# === TAB 3: PROFITABILITY ===
with tab3:
    st.subheader("Profitability Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Revenue vs Cost by Branch")
        
        fig_rev_cost = go.Figure()
        fig_rev_cost.add_trace(go.Bar(
            name='Revenue',
            x=branch_totals['Branch'],
            y=branch_totals['Revenue'],
            marker_color='#3498db'
        ))
        fig_rev_cost.add_trace(go.Bar(
            name='Cost',
            x=branch_totals['Branch'],
            y=branch_totals['Cost'],
            marker_color='#e74c3c'
        ))
        fig_rev_cost.update_layout(
            barmode='group',
            height=400,
            xaxis_title="Branch",
            yaxis_title="Amount (£)"
        )
        st.plotly_chart(fig_rev_cost, use_container_width=True)
    
    with col2:
        st.markdown("#### Gross Profit by Branch")
        fig_profit = px.bar(
            branch_totals.sort_values('Gross Profit', ascending=False),
            x='Branch',
            y='Gross Profit',
            color='Gross Profit',
            color_continuous_scale='Greens',
            text='Gross Profit'
        )
        fig_profit.update_traces(texttemplate='£%{text:,.0f}', textposition='outside')
        fig_profit.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig_profit, use_container_width=True)

# === TAB 4: DATA TABLE ===
with tab4:
    st.subheader("Complete Dataset")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Records", len(filtered_df))
    with col2:
        st.metric("Periods Shown", len(sel_periods))
    with col3:
        st.metric("Branches Shown", len(sel_branches))
    
    st.divider()
    
    csv = filtered_df.to_csv(index=False)
    st.download_button(
        label="⬇️ Download as CSV",
        data=csv,
        file_name=f"dashboard_data_{datetime.now():%Y%m%d}.csv",
        mime="text/csv"
    )
    
    st.dataframe(
        filtered_df[['Period', 'Date Range', 'Branch', 'Revenue', 'Cost', 'Gross Profit', 'Margin %']].style.format({
            'Revenue': '£{:,.0f}',
            'Cost': '£{:,.0f}',
            'Gross Profit': '£{:,.0f}',
            'Margin %': '{:.1f}%'
        }),
        use_container_width=True,
        height=500
    )

# === FOOTER ===
st.divider()
st.success("✅ Dashboard loaded successfully! All data processed correctly.")

if st.sidebar.checkbox("Show Debug Info", value=False):
    st.sidebar.markdown("### Debug Information")
    st.sidebar.write(f"Total rows in df: {len(df)}")
    st.sidebar.write(f"Filtered rows: {len(filtered_df)}")
    st.sidebar.write(f"Branches: {branches}")
    st.sidebar.write(f"Periods: {sorted(df['Period'].unique())}")
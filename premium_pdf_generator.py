"""
Premium Investor-Grade PDF Report Generator
Produces executive-level financial analysis reports comparable to Big 4 consulting firms
"""

import io
import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, 
    PageBreak, Image as RLImage, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfgen import canvas
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.io as pio

# Configure Plotly for PDF export
pio.kaleido.scope.mathjax = None


class InvestorGradeReport:
    """Generate premium investor-ready financial reports"""
    
    def __init__(self, client_name, report_title="FINANCIAL PERFORMANCE ANALYSIS"):
        self.client_name = client_name
        self.report_title = report_title
        self.colors = {
            'primary': colors.HexColor('#1a3a52'),      # Deep navy
            'secondary': colors.HexColor('#2c5f7f'),    # Medium blue
            'accent': colors.HexColor('#e8953c'),       # Gold
            'success': colors.HexColor('#2d7a4e'),      # Green
            'warning': colors.HexColor('#d97642'),      # Orange
            'danger': colors.HexColor('#c1403d'),       # Red
            'neutral': colors.HexColor('#5a6c7d'),      # Grey
            'light': colors.HexColor('#f4f6f8'),        # Light grey
        }
        
        self.styles = self._create_styles()
        
    def _create_styles(self):
        """Create custom styles for the report"""
        styles = getSampleStyleSheet()
        
        # Cover page title
        styles.add(ParagraphStyle(
            name='CoverTitle',
            parent=styles['Heading1'],
            fontSize=32,
            textColor=self.colors['primary'],
            spaceAfter=12,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            leading=38
        ))
        
        # Cover subtitle
        styles.add(ParagraphStyle(
            name='CoverSubtitle',
            parent=styles['Normal'],
            fontSize=14,
            textColor=self.colors['neutral'],
            alignment=TA_CENTER,
            fontName='Helvetica',
            spaceAfter=6
        ))
        
        # Section headers
        styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=styles['Heading1'],
            fontSize=18,
            textColor=self.colors['primary'],
            spaceAfter=16,
            spaceBefore=24,
            fontName='Helvetica-Bold',
            borderWidth=0,
            borderColor=self.colors['accent'],
            borderPadding=0,
            leftIndent=0,
            borderRadius=0,
            backColor=None
        ))
        
        # Subsection headers
        styles.add(ParagraphStyle(
            name='SubsectionHeader',
            parent=styles['Heading2'],
            fontSize=14,
            textColor=self.colors['secondary'],
            spaceAfter=12,
            spaceBefore=16,
            fontName='Helvetica-Bold'
        ))
        
        # Executive summary style
        styles.add(ParagraphStyle(
            name='ExecutiveSummary',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.black,
            alignment=TA_JUSTIFY,
            leading=16,
            spaceAfter=8,
            fontName='Helvetica'
        ))
        
        # Insight box
        styles.add(ParagraphStyle(
            name='InsightBox',
            parent=styles['Normal'],
            fontSize=10,
            textColor=self.colors['primary'],
            alignment=TA_LEFT,
            leading=14,
            leftIndent=12,
            rightIndent=12,
            fontName='Helvetica',
            spaceAfter=4
        ))
        
        # Key finding
        styles.add(ParagraphStyle(
            name='KeyFinding',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.black,
            alignment=TA_LEFT,
            leading=15,
            leftIndent=20,
            fontName='Helvetica',
            spaceAfter=8,
            bulletIndent=10,
            bulletFontSize=11
        ))
        
        # Footer style
        styles.add(ParagraphStyle(
            name='Footer',
            parent=styles['Normal'],
            fontSize=8,
            textColor=self.colors['neutral'],
            alignment=TA_CENTER,
            fontName='Helvetica'
        ))
        
        return styles
    
    def _add_page_number(self, canvas, doc):
        """Add page numbers and footer to each page"""
        page_num = canvas.getPageNumber()
        text = f"Page {page_num}"
        canvas.saveState()
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(self.colors['neutral'])
        canvas.drawCentredString(letter[0]/2, 0.5*inch, text)
        
        # Add confidentiality notice
        canvas.setFont('Helvetica', 7)
        canvas.drawCentredString(
            letter[0]/2, 
            0.35*inch, 
            "CONFIDENTIAL - For Internal Use Only"
        )
        canvas.restoreState()
    
    def generate_report(self, df, branches, config):
        """Generate the complete investor-grade report"""
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=letter,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=0.75*inch,
            bottomMargin=1*inch
        )
        
        story = []
        
        # 1. COVER PAGE
        story.extend(self._create_cover_page(df))
        story.append(PageBreak())
        
        # 2. EXECUTIVE SUMMARY
        story.extend(self._create_executive_summary(df, branches))
        story.append(PageBreak())
        
        # 3. TABLE OF CONTENTS
        story.extend(self._create_table_of_contents())
        story.append(PageBreak())
        
        # 4. FINANCIAL OVERVIEW
        story.extend(self._create_financial_overview(df, branches))
        story.append(PageBreak())
        
        # 5. BRANCH PERFORMANCE ANALYSIS
        story.extend(self._create_branch_analysis(df, branches))
        story.append(PageBreak())
        
        # 6. TREND ANALYSIS
        story.extend(self._create_trend_analysis(df, branches))
        story.append(PageBreak())
        
        # 7. RISK & OPPORTUNITY ANALYSIS
        story.extend(self._create_risk_opportunity_analysis(df, branches))
        story.append(PageBreak())
        
        # 8. STRATEGIC RECOMMENDATIONS
        story.extend(self._create_recommendations(df, branches))
        story.append(PageBreak())
        
        # 9. APPENDIX
        story.extend(self._create_appendix(df, branches))
        
        # Build the PDF
        doc.build(story, onFirstPage=self._add_page_number, onLaterPages=self._add_page_number)
        
        buffer.seek(0)
        return buffer
    
    def _create_cover_page(self, df):
        """Create professional cover page"""
        elements = []
        
        # Add spacing from top
        elements.append(Spacer(1, 1.5*inch))
        
        # Company name
        elements.append(Paragraph(self.client_name.upper(), self.styles['CoverTitle']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Report title
        elements.append(Paragraph(self.report_title, self.styles['CoverSubtitle']))
        elements.append(Spacer(1, 0.5*inch))
        
        # Decorative line
        line_table = Table([['']], colWidths=[5*inch])
        line_table.setStyle(TableStyle([
            ('LINEABOVE', (0, 0), (-1, 0), 2, self.colors['accent']),
            ('LINEBELOW', (0, 0), (-1, 0), 2, self.colors['accent']),
        ]))
        elements.append(line_table)
        elements.append(Spacer(1, 0.5*inch))
        
        # Period coverage
        periods = sorted(df['Period'].unique())
        period_text = f"Period {periods[0]} - Period {periods[-1]}"
        elements.append(Paragraph(period_text, self.styles['CoverSubtitle']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Report date
        report_date = datetime.now().strftime("%B %d, %Y")
        elements.append(Paragraph(f"Generated: {report_date}", self.styles['CoverSubtitle']))
        elements.append(Spacer(1, 1.5*inch))
        
        # Key metrics box
        total_revenue = df['Revenue'].sum()
        total_profit = df['Gross Profit'].sum()
        avg_margin = df['Margin %'].mean()
        
        metrics_data = [
            ['EXECUTIVE DASHBOARD METRICS'],
            ['Total Revenue', f'£{total_revenue:,.0f}'],
            ['Gross Profit', f'£{total_profit:,.0f}'],
            ['Average Margin', f'{avg_margin:.1f}%'],
            ['Reporting Periods', str(len(periods))],
            ['Business Units', str(len(df['Branch'].unique()))]
        ]
        
        metrics_table = Table(metrics_data, colWidths=[2.5*inch, 2.5*inch])
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('SPAN', (0, 0), (-1, 0)),
            
            ('BACKGROUND', (0, 1), (-1, -1), self.colors['light']),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, 1), (0, -1), 'Helvetica'),
            ('FONTNAME', (1, 1), (1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, -1), 11),
            
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('LEFTPADDING', (0, 0), (-1, -1), 15),
            ('GRID', (0, 0), (-1, -1), 0.5, self.colors['secondary'])
        ]))
        
        elements.append(metrics_table)
        elements.append(Spacer(1, 1*inch))
        
        # Confidentiality notice
        conf_text = "CONFIDENTIAL AND PROPRIETARY<br/>This report contains sensitive financial information"
        elements.append(Paragraph(conf_text, self.styles['Footer']))
        
        return elements
    
    def _create_executive_summary(self, df, branches):
        """Create executive summary with key insights"""
        elements = []
        
        elements.append(Paragraph("EXECUTIVE SUMMARY", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Calculate key metrics
        total_revenue = df['Revenue'].sum()
        total_cost = df['Cost'].sum()
        total_profit = df['Gross Profit'].sum()
        overall_margin = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
        
        # Period-over-period analysis
        df_sorted = df.sort_values('Period_Int')
        periods = df_sorted['Period'].unique()
        
        if len(periods) >= 2:
            recent_period = df_sorted[df_sorted['Period'] == periods[-1]]
            previous_period = df_sorted[df_sorted['Period'] == periods[-2]]
            
            recent_rev = recent_period['Revenue'].sum()
            prev_rev = previous_period['Revenue'].sum()
            revenue_growth = ((recent_rev - prev_rev) / prev_rev * 100) if prev_rev > 0 else 0
        else:
            revenue_growth = 0
        
        # Branch performance
        branch_totals = df.groupby('Branch').agg({
            'Revenue': 'sum',
            'Gross Profit': 'sum',
            'Margin %': 'mean'
        }).reset_index()
        
        best_performer = branch_totals.loc[branch_totals['Revenue'].idxmax()]
        worst_performer = branch_totals.loc[branch_totals['Revenue'].idxmin()]
        
        # Overview paragraph
        overview = f"""
        This report presents a comprehensive financial analysis of {self.client_name}'s operations 
        across {len(branches)} business units over {len(periods)} reporting periods. The analysis 
        reveals total revenue of <b>£{total_revenue:,.0f}</b> with an overall gross profit margin 
        of <b>{overall_margin:.1f}%</b>. Period-over-period revenue growth stands at 
        <b>{revenue_growth:+.1f}%</b>, indicating {'strong momentum' if revenue_growth > 5 else 'stable performance' if revenue_growth > 0 else 'areas requiring attention'}.
        """
        
        elements.append(Paragraph(overview, self.styles['ExecutiveSummary']))
        elements.append(Spacer(1, 0.3*inch))
        
        # Key Findings section
        elements.append(Paragraph("KEY FINDINGS", self.styles['SubsectionHeader']))
        
        findings = [
            f"<b>Revenue Performance:</b> Total revenue of £{total_revenue:,.0f} across all business units with a {revenue_growth:+.1f}% period-over-period growth trajectory",
            
            f"<b>Profitability Analysis:</b> Overall gross profit margin of {overall_margin:.1f}% demonstrates {'strong' if overall_margin > 30 else 'moderate' if overall_margin > 20 else 'challenged'} operational efficiency with total gross profit of £{total_profit:,.0f}",
            
            f"<b>Branch Performance:</b> {best_performer['Branch']} leads with £{best_performer['Revenue']:,.0f} in revenue, while {worst_performer['Branch']} represents an opportunity for strategic intervention",
            
            f"<b>Cost Management:</b> Total operational costs of £{total_cost:,.0f} require ongoing monitoring to maintain competitive margin profiles",
            
            f"<b>Margin Sustainability:</b> Branch-level margin variance suggests potential for operational standardization and best practice sharing across the organization"
        ]
        
        for finding in findings:
            elements.append(Paragraph(f"• {finding}", self.styles['KeyFinding']))
        
        elements.append(Spacer(1, 0.3*inch))
        
        # Strategic outlook box
        outlook_data = [[Paragraph('<b>STRATEGIC OUTLOOK</b>', self.styles['InsightBox'])]]
        
        if overall_margin > 25 and revenue_growth > 3:
            outlook_text = "Strong financial position with positive momentum across key metrics. Focus should shift to market expansion and operational scaling."
        elif overall_margin > 20 or revenue_growth > 0:
            outlook_text = "Solid foundation with selective opportunities for margin enhancement and revenue acceleration through targeted initiatives."
        else:
            outlook_text = "Performance indicates need for strategic review of cost structures and revenue generation capabilities. Immediate attention to underperforming units recommended."
        
        outlook_data.append([Paragraph(outlook_text, self.styles['InsightBox'])])
        
        outlook_table = Table(outlook_data, colWidths=[6.5*inch])
        outlook_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('BACKGROUND', (0, 1), (-1, -1), self.colors['light']),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('BOX', (0, 0), (-1, -1), 1, self.colors['accent'])
        ]))
        
        elements.append(outlook_table)
        
        return elements
    
    def _create_table_of_contents(self):
        """Create table of contents"""
        elements = []
        
        elements.append(Paragraph("TABLE OF CONTENTS", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.3*inch))
        
        toc_data = [
            ['Section', 'Page'],
            ['Executive Summary', '2'],
            ['Financial Overview', '4'],
            ['Branch Performance Analysis', '5'],
            ['Trend Analysis & Insights', '6'],
            ['Risk & Opportunity Assessment', '7'],
            ['Strategic Recommendations', '8'],
            ['Appendix: Detailed Data Tables', '9']
        ]
        
        toc_table = Table(toc_data, colWidths=[5*inch, 1.5*inch])
        toc_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            
            ('FONTNAME', (0, 1), (0, -1), 'Helvetica'),
            ('FONTNAME', (1, 1), (1, -1), 'Helvetica-Bold'),
            ('ALIGN', (1, 1), (1, -1), 'CENTER'),
            
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            
            ('LINEBELOW', (0, 0), (-1, 0), 1, self.colors['accent']),
            ('LINEBELOW', (0, 1), (-1, -2), 0.5, colors.grey),
            ('LINEBELOW', (0, -1), (-1, -1), 1, self.colors['primary'])
        ]))
        
        elements.append(toc_table)
        
        return elements
    
    def _create_financial_overview(self, df, branches):
        """Create detailed financial overview section"""
        elements = []
        
        elements.append(Paragraph("FINANCIAL OVERVIEW", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Period-by-period summary
        period_summary = df.groupby('Period').agg({
            'Revenue': 'sum',
            'Cost': 'sum',
            'Gross Profit': 'sum'
        }).reset_index()
        
        period_summary['Margin %'] = (period_summary['Gross Profit'] / period_summary['Revenue'] * 100).round(1)
        period_summary['Growth %'] = period_summary['Revenue'].pct_change() * 100
        
        elements.append(Paragraph("Period-by-Period Performance", self.styles['SubsectionHeader']))
        
        # Create performance table
        table_data = [['Period', 'Revenue', 'Costs', 'Gross Profit', 'Margin %', 'Growth %']]
        
        for _, row in period_summary.iterrows():
            growth_display = f"{row['Growth %']:+.1f}%" if not pd.isna(row['Growth %']) else "—"
            table_data.append([
                str(row['Period']),
                f"£{row['Revenue']:,.0f}",
                f"£{row['Cost']:,.0f}",
                f"£{row['Gross Profit']:,.0f}",
                f"{row['Margin %']:.1f}%",
                growth_display
            ])
        
        # Add totals row
        total_row = ['TOTAL', 
                    f"£{period_summary['Revenue'].sum():,.0f}",
                    f"£{period_summary['Cost'].sum():,.0f}",
                    f"£{period_summary['Gross Profit'].sum():,.0f}",
                    f"{(period_summary['Gross Profit'].sum() / period_summary['Revenue'].sum() * 100):.1f}%",
                    "—"]
        table_data.append(total_row)
        
        perf_table = Table(table_data, colWidths=[0.8*inch, 1.3*inch, 1.3*inch, 1.3*inch, 1*inch, 1*inch])
        perf_table.setStyle(TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -2), 9),
            ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),
            
            # Total row
            ('BACKGROUND', (0, -1), (-1, -1), self.colors['light']),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('LINEABOVE', (0, -1), (-1, -1), 2, self.colors['primary']),
            
            # Alternating rows
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f9f9f9')]),
            
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        
        elements.append(perf_table)
        elements.append(Spacer(1, 0.4*inch))
        
        # Key financial ratios
        elements.append(Paragraph("Key Financial Ratios", self.styles['SubsectionHeader']))
        
        avg_revenue_per_period = period_summary['Revenue'].mean()
        revenue_volatility = period_summary['Revenue'].std() / avg_revenue_per_period * 100 if avg_revenue_per_period > 0 else 0
        margin_consistency = period_summary['Margin %'].std()
        
        ratios_data = [
            ['Metric', 'Value', 'Interpretation'],
            ['Average Revenue per Period', f'£{avg_revenue_per_period:,.0f}', 'Baseline performance level'],
            ['Revenue Volatility', f'{revenue_volatility:.1f}%', 'Indicates revenue stability'],
            ['Margin Consistency (σ)', f'{margin_consistency:.1f}%', 'Lower is better'],
            ['Cost-to-Revenue Ratio', f'{(period_summary["Cost"].sum() / period_summary["Revenue"].sum() * 100):.1f}%', 'Operating efficiency measure']
        ]
        
        ratios_table = Table(ratios_data, colWidths=[2.3*inch, 1.5*inch, 2.7*inch])
        ratios_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['secondary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('BACKGROUND', (0, 1), (-1, -1), self.colors['light']),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        
        elements.append(ratios_table)
        
        return elements
    
    def _create_branch_analysis(self, df, branches):
        """Deep-dive analysis of each branch"""
        elements = []
        
        elements.append(Paragraph("BRANCH PERFORMANCE ANALYSIS", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Calculate branch metrics
        branch_totals = df.groupby('Branch').agg({
            'Revenue': 'sum',
            'Cost': 'sum',
            'Gross Profit': 'sum'
        }).reset_index()
        
        branch_totals['Margin %'] = (branch_totals['Gross Profit'] / branch_totals['Revenue'] * 100).round(1)
        branch_totals['Revenue Share %'] = (branch_totals['Revenue'] / branch_totals['Revenue'].sum() * 100).round(1)
        branch_totals = branch_totals.sort_values('Revenue', ascending=False)
        
        # Ranking table
        elements.append(Paragraph("Branch Performance Rankings", self.styles['SubsectionHeader']))
        
        rank_data = [['Rank', 'Branch', 'Revenue', 'Margin %', 'Share %', 'Status']]
        
        for idx, (_, row) in enumerate(branch_totals.iterrows(), 1):
            if row['Margin %'] >= 25:
                status = "⭐ Excellent"
            elif row['Margin %'] >= 20:
                status = "✓ Good"
            elif row['Margin %'] >= 15:
                status = "→ Fair"
            else:
                status = "⚠ Review"
                
            rank_data.append([
                str(idx),
                row['Branch'],
                f"£{row['Revenue']:,.0f}",
                f"{row['Margin %']:.1f}%",
                f"{row['Revenue Share %']:.1f}%",
                status
            ])
        
        rank_table = Table(rank_data, colWidths=[0.6*inch, 1.8*inch, 1.4*inch, 1*inch, 1*inch, 1*inch])
        rank_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f9f9f9')]),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        
        elements.append(rank_table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Individual branch insights
        elements.append(Paragraph("Branch-Level Insights", self.styles['SubsectionHeader']))
        
        for _, row in branch_totals.head(3).iterrows():  # Top 3 branches
            branch_df = df[df['Branch'] == row['Branch']].sort_values('Period_Int')
            
            if len(branch_df) > 1:
                revenue_trend = branch_df['Revenue'].iloc[-1] - branch_df['Revenue'].iloc[0]
                trend_direction = "upward" if revenue_trend > 0 else "downward"
            else:
                trend_direction = "stable"
            
            insight = f"<b>{row['Branch']}</b>: Contributes {row['Revenue Share %']:.1f}% of total revenue with a {row['Margin %']:.1f}% margin. "
            insight += f"Revenue trend is {trend_direction} across the reporting period. "
            
            if row['Margin %'] >= 25:
                insight += "Demonstrates strong operational efficiency and cost management."
            elif row['Margin %'] < 20:
                insight += "Margin improvement initiatives recommended."
            
            elements.append(Paragraph(f"• {insight}", self.styles['KeyFinding']))
        
        return elements
    
    def _create_trend_analysis(self, df, branches):
        """Trend analysis with embedded charts"""
        elements = []
        
        elements.append(Paragraph("TREND ANALYSIS & INSIGHTS", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Revenue trend chart
        fig = make_subplots(
            rows=2, cols=1,
            subplot_titles=('Revenue Trends by Branch', 'Margin Performance Over Time'),
            vertical_spacing=0.15
        )
        
        for branch in branches:
            branch_df = df[df['Branch'] == branch].sort_values('Period_Int')
            fig.add_trace(
                go.Scatter(
                    x=branch_df['Period_Int'],
                    y=branch_df['Revenue'],
                    name=branch,
                    mode='lines+markers',
                    line=dict(width=2)
                ),
                row=1, col=1
            )
            
            fig.add_trace(
                go.Scatter(
                    x=branch_df['Period_Int'],
                    y=branch_df['Margin %'],
                    name=branch,
                    mode='lines+markers',
                    line=dict(width=2),
                    showlegend=False
                ),
                row=2, col=1
            )
        
        fig.update_xaxes(title_text="Period", row=2, col=1)
        fig.update_yaxes(title_text="Revenue (£)", row=1, col=1)
        fig.update_yaxes(title_text="Margin %", row=2, col=1)
        
        fig.update_layout(
            height=600,
            showlegend=True,
            legend=dict(x=1.05, y=1),
            font=dict(size=10),
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        # Save chart as image
        img_bytes = pio.to_image(fig, format='png', width=800, height=600, scale=2)
        img_buffer = io.BytesIO(img_bytes)
        
        img = RLImage(img_buffer, width=6.5*inch, height=4.875*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.3*inch))
        
        # Trend commentary
        elements.append(Paragraph("Trend Commentary", self.styles['SubsectionHeader']))
        
        period_totals = df.groupby('Period_Int')['Revenue'].sum()
        if len(period_totals) > 1:
            overall_trend = "increasing" if period_totals.iloc[-1] > period_totals.iloc[0] else "decreasing"
            pct_change = ((period_totals.iloc[-1] - period_totals.iloc[0]) / period_totals.iloc[0] * 100)
            
            commentary = f"""
            Revenue demonstrates an {overall_trend} trend over the reporting period, with a 
            {abs(pct_change):.1f}% {'increase' if pct_change > 0 else 'decrease'} from first to last period. 
            Margin performance shows {'consistency' if df['Margin %'].std() < 3 else 'variability'} across branches, 
            suggesting {'strong' if df['Margin %'].std() < 3 else 'inconsistent'} operational standardization.
            """
        else:
            commentary = "Insufficient period data for trend analysis. Recommend tracking over additional periods."
        
        elements.append(Paragraph(commentary, self.styles['ExecutiveSummary']))
        
        return elements
    
    def _create_risk_opportunity_analysis(self, df, branches):
        """Risk and opportunity assessment"""
        elements = []
        
        elements.append(Paragraph("RISK & OPPORTUNITY ASSESSMENT", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Calculate risk indicators
        branch_totals = df.groupby('Branch').agg({
            'Revenue': 'sum',
            'Margin %': 'mean',
            'Cost': 'sum'
        }).reset_index()
        
        # Identify risks
        elements.append(Paragraph("Identified Risks", self.styles['SubsectionHeader']))
        
        risks = []
        
        # Low margin branches
        low_margin_branches = branch_totals[branch_totals['Margin %'] < 20]
        if len(low_margin_branches) > 0:
            risk_text = f"<b>Margin Pressure:</b> {len(low_margin_branches)} branch(es) operating below 20% margin threshold - "
            risk_text += ", ".join(low_margin_branches['Branch'].tolist())
            risk_text += ". Immediate cost structure review recommended."
            risks.append(risk_text)
        
        # Revenue concentration
        top_branch_share = branch_totals['Revenue'].max() / branch_totals['Revenue'].sum() * 100
        if top_branch_share > 40:
            risk_text = f"<b>Revenue Concentration:</b> Single branch represents {top_branch_share:.1f}% of total revenue, creating dependency risk."
            risks.append(risk_text)
        
        # Margin volatility
        margin_std = df.groupby('Branch')['Margin %'].std()
        volatile_branches = margin_std[margin_std > 5]
        if len(volatile_branches) > 0:
            risk_text = f"<b>Margin Volatility:</b> {len(volatile_branches)} branch(es) showing inconsistent margin performance, indicating operational instability."
            risks.append(risk_text)
        
        if not risks:
            risks.append("<b>Low Risk Profile:</b> Current operations demonstrate stable performance across key risk metrics.")
        
        for risk in risks:
            elements.append(Paragraph(f"• {risk}", self.styles['KeyFinding']))
        
        elements.append(Spacer(1, 0.3*inch))
        
        # Identify opportunities
        elements.append(Paragraph("Strategic Opportunities", self.styles['SubsectionHeader']))
        
        opportunities = []
        
        # High margin branches
        high_margin_branches = branch_totals[branch_totals['Margin %'] > 25]
        if len(high_margin_branches) > 0:
            opp_text = f"<b>Best Practice Replication:</b> {', '.join(high_margin_branches['Branch'].tolist())} demonstrate(s) superior margin performance. "
            opp_text += "Analyze and replicate operational practices across lower-performing units."
            opportunities.append(opp_text)
        
        # Growth potential
        avg_margin = branch_totals['Margin %'].mean()
        improvement_potential = branch_totals[branch_totals['Margin %'] < avg_margin]
        if len(improvement_potential) > 0:
            potential_value = improvement_potential.apply(
                lambda x: x['Revenue'] * (avg_margin - x['Margin %']) / 100, 
                axis=1
            ).sum()
            opp_text = f"<b>Margin Enhancement:</b> Bringing underperforming branches to average margin could generate additional £{potential_value:,.0f} in gross profit."
            opportunities.append(opp_text)
        
        # Scale advantages
        total_revenue = branch_totals['Revenue'].sum()
        opp_text = f"<b>Scale Optimization:</b> Current revenue base of £{total_revenue:,.0f} provides foundation for procurement leverage and shared services expansion."
        opportunities.append(opp_text)
        
        for opp in opportunities:
            elements.append(Paragraph(f"• {opp}", self.styles['KeyFinding']))
        
        return elements
    
    def _create_recommendations(self, df, branches):
        """Strategic recommendations section"""
        elements = []
        
        elements.append(Paragraph("STRATEGIC RECOMMENDATIONS", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Calculate metrics for recommendations
        branch_totals = df.groupby('Branch').agg({
            'Revenue': 'sum',
            'Margin %': 'mean',
            'Cost': 'sum'
        }).reset_index()
        
        avg_margin = branch_totals['Margin %'].mean()
        total_revenue = branch_totals['Revenue'].sum()
        
        recommendations = []
        
        # Priority 1: Quick wins
        elements.append(Paragraph("Priority 1: Immediate Actions (0-3 Months)", self.styles['SubsectionHeader']))
        
        low_performers = branch_totals[branch_totals['Margin %'] < avg_margin]
        if len(low_performers) > 0:
            rec = f"<b>Margin Improvement Initiative:</b> Conduct deep-dive cost analysis for {', '.join(low_performers['Branch'].tolist())}. "
            rec += "Target: 2-3 percentage point margin improvement through operational efficiency. "
            rec += f"Projected impact: £{(low_performers['Revenue'].sum() * 0.025):,.0f} additional annual profit."
            recommendations.append(rec)
        
        rec = "<b>Performance Dashboard Implementation:</b> Deploy real-time financial monitoring to enable proactive management and early intervention for underperformance."
        recommendations.append(rec)
        
        for rec in recommendations:
            elements.append(Paragraph(f"• {rec}", self.styles['KeyFinding']))
        
        recommendations = []
        
        # Priority 2: Medium-term initiatives
        elements.append(Spacer(1, 0.2*inch))
        elements.append(Paragraph("Priority 2: Strategic Initiatives (3-12 Months)", self.styles['SubsectionHeader']))
        
        rec = "<b>Best Practice Standardization:</b> Document and roll out operational procedures from top-performing branches across all units. Include training program and quarterly audits."
        recommendations.append(rec)
        
        rec = f"<b>Procurement Optimization:</b> Leverage £{total_revenue:,.0f} revenue scale for enhanced vendor negotiations. Target: 3-5% cost reduction through volume purchasing and contract consolidation."
        recommendations.append(rec)
        
        rec = "<b>Revenue Diversification:</b> Reduce dependence on single revenue streams through service expansion analysis and market opportunity assessment."
        recommendations.append(rec)
        
        for rec in recommendations:
            elements.append(Paragraph(f"• {rec}", self.styles['KeyFinding']))
        
        recommendations = []
        
        # Priority 3: Long-term strategic
        elements.append(Spacer(1, 0.2*inch))
        elements.append(Paragraph("Priority 3: Long-Term Strategic (12+ Months)", self.styles['SubsectionHeader']))
        
        rec = "<b>Portfolio Optimization:</b> Evaluate branch portfolio for potential consolidation, expansion, or divestment opportunities based on ROI and strategic fit."
        recommendations.append(rec)
        
        rec = "<b>Technology Investment:</b> Assess automation and digital transformation opportunities to reduce cost-to-serve while maintaining quality standards."
        recommendations.append(rec)
        
        rec = "<b>Market Expansion Analysis:</b> With proven operational model, investigate geographic or service line expansion opportunities for sustainable growth."
        recommendations.append(rec)
        
        for rec in recommendations:
            elements.append(Paragraph(f"• {rec}", self.styles['KeyFinding']))
        
        return elements
    
    def _create_appendix(self, df, branches):
        """Detailed data appendix"""
        elements = []
        
        elements.append(Paragraph("APPENDIX: DETAILED DATA TABLES", self.styles['SectionHeader']))
        elements.append(Spacer(1, 0.2*inch))
        
        # Complete dataset
        elements.append(Paragraph("Complete Performance Dataset", self.styles['SubsectionHeader']))
        
        # Prepare data sorted by period and branch
        display_df = df[['Period', 'Branch', 'Revenue', 'Cost', 'Gross Profit', 'Margin %']].copy()
        display_df = display_df.sort_values(['Period', 'Branch'])
        
        # Create table data
        table_data = [['Period', 'Branch', 'Revenue', 'Cost', 'Gross Profit', 'Margin %']]
        
        for _, row in display_df.iterrows():
            table_data.append([
                str(row['Period']),
                row['Branch'],
                f"£{row['Revenue']:,.0f}",
                f"£{row['Cost']:,.0f}",
                f"£{row['Gross Profit']:,.0f}",
                f"{row['Margin %']:.1f}%"
            ])
        
        # Create table with appropriate styling
        data_table = Table(table_data, colWidths=[0.7*inch, 1.5*inch, 1.2*inch, 1.2*inch, 1.2*inch, 0.8*inch])
        data_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary']),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (1, 0), (1, -1), 'LEFT'),
            ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f9f9f9')]),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        
        elements.append(data_table)
        elements.append(Spacer(1, 0.3*inch))
        
        # Methodology notes
        elements.append(Paragraph("Methodology & Assumptions", self.styles['SubsectionHeader']))
        
        methodology = """
        <b>Data Sources:</b> Revenue and cost data derived from operational financial systems.<br/>
        <b>Calculations:</b> Gross Profit = Revenue - Cost; Margin % = (Gross Profit / Revenue) × 100<br/>
        <b>Period Definition:</b> Financial reporting periods as defined in source data<br/>
        <b>Scope:</b> Analysis covers all operational branches over complete reporting periods<br/>
        <b>Limitations:</b> Analysis based on gross profit only; net profit and EBITDA require additional data
        """
        
        elements.append(Paragraph(methodology, self.styles['ExecutiveSummary']))
        
        # Report metadata
        elements.append(Spacer(1, 0.3*inch))
        elements.append(Paragraph("Report Information", self.styles['SubsectionHeader']))
        
        metadata = f"""
        <b>Generated:</b> {datetime.now().strftime("%B %d, %Y at %H:%M")}<br/>
        <b>Report Version:</b> 1.0<br/>
        <b>Confidentiality:</b> This report contains proprietary financial information<br/>
        <b>Distribution:</b> Internal management use only
        """
        
        elements.append(Paragraph(metadata, self.styles['ExecutiveSummary']))
        
        return elements


def generate_investor_grade_pdf(df, branches, client_name, config):
    """
    Wrapper function to generate the complete investor-grade report
    
    Parameters:
    - df: DataFrame with columns: Period, Branch, Revenue, Cost, Gross Profit, Margin %, Period_Int
    - branches: List of branch names
    - client_name: Company name for branding
    - config: Configuration dictionary (optional, for future expansion)
    
    Returns:
    - BytesIO buffer containing the PDF
    """
    
    report = InvestorGradeReport(client_name=client_name)
    pdf_buffer = report.generate_report(df, branches, config)
    
    return pdf_buffer

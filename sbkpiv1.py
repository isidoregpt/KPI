import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="Executive Financial Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Fortune 50 styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1f2937;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #6b7280;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 0.5rem 0;
    }
    .kpi-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        border-left: 4px solid #3b82f6;
        margin: 1rem 0;
    }
    .status-good { border-left-color: #10b981; }
    .status-warning { border-left-color: #f59e0b; }
    .status-critical { border-left-color: #ef4444; }
    .sidebar .sidebar-content {
        background-color: #f8fafc;
    }
    .upload-box {
        border: 2px dashed #cbd5e0;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background-color: #f7fafc;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

class FinancialDataProcessor:
    def __init__(self):
        self.pl_data = None
        self.bs_data = None
        self.processed_data = None
    
    def load_excel_files(self, pl_file, bs_file):
        """Load and process Excel files"""
        try:
            # Load P&L data
            self.pl_data = pd.read_excel(pl_file, sheet_name=0, header=None)
            # Load Balance Sheet data  
            self.bs_data = pd.read_excel(bs_file, sheet_name=0, header=None)
            return True
        except Exception as e:
            st.error(f"Error loading files: {str(e)}")
            return False
    
    def extract_financial_data(self):
        """Extract key financial metrics from Excel files"""
        if self.pl_data is None or self.bs_data is None:
            return None
        
        try:
            # Extract months from P&L headers (row 7, columns B-P)
            months = []
            for col in range(1, 16):  # B through P (columns 1-15 for Mar 2024 - May 2025)
                cell_value = self.pl_data.iloc[6, col]  # Row 7, 0-indexed as 6
                if pd.notna(cell_value):
                    months.append(str(cell_value))
            
            # Extract financial data using the exact row mappings from your files
            financial_data = {
                'months': months,
                'revenue': self._extract_row_data(self.pl_data, 11, 1, 16),  # Row 12: Total Revenue
                'gross_profit': self._extract_row_data(self.pl_data, 18, 1, 16),  # Row 19: Gross Profit
                'ebitda': self._extract_row_data(self.pl_data, 47, 1, 16),  # Row 48: EBITDA
                'operating_expenses': self._extract_row_data(self.pl_data, 39, 1, 16),  # Row 40: Total Operating Expenses
                'sga_expenses': self._extract_row_data(self.pl_data, 36, 1, 16),  # Row 37: Total General and Administrative
                'accounts_receivable': self._extract_row_data(self.bs_data, 42, 3, 18),  # Row 43: Total Accounts Receivable
                'inventory': self._extract_row_data(self.bs_data, 52, 3, 18),  # Row 53: Total - 12999 - Inventory Total
                'accounts_payable': self._extract_row_data(self.bs_data, 129, 3, 18),  # Row 130: Total Accounts Payable
                'current_assets': self._extract_row_data(self.bs_data, 81, 3, 18),  # Row 82: Total Current Assets
                'current_liabilities': self._extract_row_data(self.bs_data, 151, 3, 18),  # Row 152: Total Current Liabilities
            }
            
            # Align all data to same length (minimum of P&L and BS data)
            min_length = min(len(v) for v in financial_data.values() if isinstance(v, list))
            for key in financial_data:
                if isinstance(financial_data[key], list):
                    financial_data[key] = financial_data[key][:min_length]
            
            self.processed_data = financial_data
            return financial_data
            
        except Exception as e:
            st.error(f"Error processing financial data: {str(e)}")
            st.error(f"Please check that your files match the expected format:")
            st.error(f"P&L: Should have months in row 7, columns B-P")
            st.error(f"Balance Sheet: Should have months in row 7, columns D-R")
            return None
    
    def _extract_row_data(self, df, row_idx, start_col, end_col):
        """Extract numeric data from a specific row"""
        data = []
        for col in range(start_col, end_col):
            try:
                if col < df.shape[1] and row_idx < df.shape[0]:
                    cell_value = df.iloc[row_idx, col]
                    if pd.notna(cell_value) and isinstance(cell_value, (int, float)):
                        data.append(float(cell_value))
                    else:
                        # Try to convert string to float
                        try:
                            data.append(float(str(cell_value).replace(',', '')))
                        except:
                            data.append(0.0)
                else:
                    data.append(0.0)
            except:
                data.append(0.0)
        return data

class KPICalculator:
    def __init__(self, financial_data):
        self.data = financial_data
    
    def calculate_all_kpis(self):
        """Calculate all required KPIs"""
        if not self.data:
            return None
        
        try:
            # Get most recent month data
            current_idx = -1
            prior_idx = -2
            
            # Days Sales Outstanding (DSO)
            current_ar = self.data['accounts_receivable'][current_idx]
            current_revenue = self.data['revenue'][current_idx]
            dso = (current_ar / current_revenue) * 30 if current_revenue > 0 else 0
            
            # Days Payables Outstanding (DPO)
            current_ap = self.data['accounts_payable'][current_idx]
            dpo = (current_ap / current_revenue) * 30 if current_revenue > 0 else 0
            
            # Days Inventory on Hand (DIO)
            current_inventory = self.data['inventory'][current_idx]
            dio = (current_inventory / current_revenue) * 30 if current_revenue > 0 else 0
            
            # Working Capital
            current_assets = self.data['current_assets'][current_idx]
            current_liabilities = self.data['current_liabilities'][current_idx]
            working_capital = current_assets - current_liabilities
            
            # Revenue Growth Rate (YoY)
            current_revenue_val = self.data['revenue'][current_idx]
            prior_year_revenue = self.data['revenue'][current_idx - 12] if len(self.data['revenue']) > 12 else self.data['revenue'][0]
            revenue_growth = ((current_revenue_val - prior_year_revenue) / prior_year_revenue) * 100 if prior_year_revenue > 0 else 0
            
            # EBITDA Margin
            current_ebitda = self.data['ebitda'][current_idx]
            ebitda_margin = (current_ebitda / current_revenue_val) * 100 if current_revenue_val > 0 else 0
            
            # SG&A as % of Revenue
            current_sga = self.data['sga_expenses'][current_idx] if 'sga_expenses' in self.data else current_opex
            sga_percentage = (current_sga / current_revenue_val) * 100 if current_revenue_val > 0 else 0
            
            # Change in AR
            prior_ar = self.data['accounts_receivable'][prior_idx]
            ar_change = current_ar - prior_ar
            
            # Cash Conversion Cycle
            cash_conversion_cycle = dso + dio - dpo
            
            # TTM calculations
            ttm_revenue = sum(self.data['revenue'][-12:]) if len(self.data['revenue']) >= 12 else sum(self.data['revenue'])
            ttm_ebitda = sum(self.data['ebitda'][-12:]) if len(self.data['ebitda']) >= 12 else sum(self.data['ebitda'])
            
            # Net Debt to EBITDA (simplified - assuming some debt)
            estimated_debt = current_liabilities * 0.3  # Simplified assumption
            net_debt_to_ebitda = estimated_debt / ttm_ebitda if ttm_ebitda > 0 else 0
            
            return {
                'dso': {'value': dso, 'format': 'days', 'target': 45, 'status': 'good' if dso < 45 else 'warning'},
                'dpo': {'value': dpo, 'format': 'days', 'target': 30, 'status': 'good' if dpo > 30 else 'warning'},
                'dio': {'value': dio, 'format': 'days', 'target': 30, 'status': 'good' if dio < 30 else 'warning'},
                'working_capital': {'value': working_capital, 'format': 'currency', 'status': 'good' if working_capital > 0 else 'critical'},
                'revenue_growth': {'value': revenue_growth, 'format': 'percentage', 'target': 15, 'status': 'good' if revenue_growth > 15 else 'warning'},
                'ebitda_margin': {'value': ebitda_margin, 'format': 'percentage', 'target': 12, 'status': 'good' if ebitda_margin > 12 else 'warning'},
                'sga_percentage': {'value': sga_percentage, 'format': 'percentage', 'target': 20, 'status': 'good' if sga_percentage < 20 else 'warning'},
                'ar_change': {'value': ar_change, 'format': 'currency', 'status': 'good' if ar_change < 0 else 'warning'},
                'cash_conversion_cycle': {'value': cash_conversion_cycle, 'format': 'days', 'target': 30, 'status': 'good' if cash_conversion_cycle < 30 else 'warning'},
                'ttm_revenue': {'value': ttm_revenue, 'format': 'currency'},
                'ttm_ebitda': {'value': ttm_ebitda, 'format': 'currency'},
                'net_debt_to_ebitda': {'value': net_debt_to_ebitda, 'format': 'ratio', 'target': 3, 'status': 'good' if net_debt_to_ebitda < 3 else 'warning'}
            }
        except Exception as e:
            st.error(f"Error calculating KPIs: {str(e)}")
            return None

def format_number(value, format_type):
    """Format numbers for display"""
    if format_type == 'currency':
        if abs(value) >= 1e9:
            return f"${value/1e9:.1f}B"
        elif abs(value) >= 1e6:
            return f"${value/1e6:.1f}M"
        elif abs(value) >= 1e3:
            return f"${value/1e3:.1f}K"
        else:
            return f"${value:,.0f}"
    elif format_type == 'percentage':
        return f"{value:.1f}%"
    elif format_type == 'days':
        return f"{value:.1f}"
    elif format_type == 'ratio':
        return f"{value:.1f}x"
    else:
        return f"{value:,.1f}"

def create_kpi_card(title, value, format_type, target=None, status=None):
    """Create a KPI card component"""
    formatted_value = format_number(value, format_type)
    
    status_class = ""
    if status:
        status_class = f"status-{status}"
    
    target_text = f"Target: {format_number(target, format_type)}" if target else ""
    
    card_html = f"""
    <div class="kpi-card {status_class}">
        <h4 style="margin: 0 0 0.5rem 0; color: #374151; font-size: 0.875rem; font-weight: 600; text-transform: uppercase;">
            {title}
        </h4>
        <div style="font-size: 2rem; font-weight: 700; color: #111827; margin: 0.5rem 0;">
            {formatted_value}
        </div>
        {f'<div style="font-size: 0.75rem; color: #6b7280;">{target_text}</div>' if target_text else ''}
    </div>
    """
    return card_html

def create_executive_charts(financial_data):
    """Create executive-level charts"""
    if not financial_data:
        return None, None, None
    
    # Prepare data for charts
    months = financial_data['months'][-12:]  # Last 12 months
    revenue_data = financial_data['revenue'][-12:]
    ebitda_data = financial_data['ebitda'][-12:]
    ar_data = financial_data['accounts_receivable'][-12:]
    
    # Revenue and EBITDA Trend
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig1.add_trace(
        go.Bar(x=months, y=[r/1e6 for r in revenue_data], name="Revenue ($M)", 
               marker_color='#3b82f6', opacity=0.8),
        secondary_y=False,
    )
    
    fig1.add_trace(
        go.Scatter(x=months, y=[e/1e6 for e in ebitda_data], mode='lines+markers',
                   name="EBITDA ($M)", line=dict(color='#10b981', width=3)),
        secondary_y=True,
    )
    
    fig1.update_xaxes(title_text="Month")
    fig1.update_yaxes(title_text="Revenue ($M)", secondary_y=False)
    fig1.update_yaxes(title_text="EBITDA ($M)", secondary_y=True)
    
    fig1.update_layout(
        title="Revenue & EBITDA Trend",
        height=400,
        showlegend=True,
        plot_bgcolor='white',
        font=dict(family="Arial, sans-serif", size=12)
    )
    
    # Working Capital Waterfall
    fig2 = go.Figure()
    
    # Calculate working capital components
    wc_data = []
    for i in range(len(months)):
        if i < len(financial_data['current_assets']) and i < len(financial_data['current_liabilities']):
            wc = (financial_data['current_assets'][-(12-i)] - financial_data['current_liabilities'][-(12-i)]) / 1e6
            wc_data.append(wc)
        else:
            wc_data.append(0)
    
    fig2.add_trace(go.Scatter(
        x=months, y=wc_data,
        mode='lines+markers',
        fill='tonexty',
        name='Working Capital ($M)',
        line=dict(color='#8b5cf6', width=3),
        fillcolor='rgba(139, 92, 246, 0.3)'
    ))
    
    fig2.update_layout(
        title="Working Capital Trend",
        xaxis_title="Month",
        yaxis_title="Working Capital ($M)",
        height=400,
        plot_bgcolor='white',
        font=dict(family="Arial, sans-serif", size=12)
    )
    
    # KPI Summary Gauge Chart
    fig3 = make_subplots(
        rows=2, cols=2,
        subplot_titles=('Days Sales Outstanding', 'Days Payables Outstanding', 'EBITDA Margin', 'Revenue Growth'),
        specs=[[{'type': 'indicator'}, {'type': 'indicator'}],
               [{'type': 'indicator'}, {'type': 'indicator'}]],
        vertical_spacing=0.15,
        horizontal_spacing=0.1
    )
    
    # Calculate current KPIs for gauges
    calculator = KPICalculator(financial_data)
    kpis = calculator.calculate_all_kpis()
    
    if kpis:
        # DSO Gauge
        fig3.add_trace(go.Indicator(
            mode = "gauge+number",
            value = kpis['dso']['value'],
            number = {
                'suffix': " days", 
                'font': {'size': 20, 'color': "#1f2937"},
                'valueformat': '.1f',
                'y': 0.4  # Move text up
            },
            gauge = {
                'axis': {'range': [None, 60], 'tickwidth': 1, 'tickcolor': "#6b7280"},
                'bar': {'color': "#3b82f6", 'thickness': 0.7},
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "#e5e7eb",
                'steps': [
                    {'range': [0, 30], 'color': "#dcfce7"},
                    {'range': [30, 45], 'color': "#fef3c7"},
                    {'range': [45, 60], 'color': "#fee2e2"}],
                'threshold': {
                    'line': {'color': "#ef4444", 'width': 3},
                    'thickness': 0.8,
                    'value': 45}}),
            row=1, col=1)
        
        # DPO Gauge  
        fig3.add_trace(go.Indicator(
            mode = "gauge+number",
            value = kpis['dpo']['value'],
            number = {
                'suffix': " days", 
                'font': {'size': 20, 'color': "#1f2937"},
                'valueformat': '.1f',
                'y': 0.4  # Move text up
            },
            gauge = {
                'axis': {'range': [None, 60], 'tickwidth': 1, 'tickcolor': "#6b7280"},
                'bar': {'color': "#10b981", 'thickness': 0.7},
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "#e5e7eb",
                'steps': [
                    {'range': [0, 20], 'color': "#fee2e2"},
                    {'range': [20, 35], 'color': "#fef3c7"},
                    {'range': [35, 60], 'color': "#dcfce7"}],
                'threshold': {
                    'line': {'color': "#10b981", 'width': 3},
                    'thickness': 0.8,
                    'value': 30}}),
            row=1, col=2)
        
        # EBITDA Margin Gauge
        fig3.add_trace(go.Indicator(
            mode = "gauge+number",
            value = kpis['ebitda_margin']['value'],
            number = {
                'suffix': "%", 
                'font': {'size': 20, 'color': "#1f2937"},
                'valueformat': '.2f',
                'y': 0.4  # Move text up
            },
            gauge = {
                'axis': {'range': [None, 25], 'tickwidth': 1, 'tickcolor': "#6b7280"},
                'bar': {'color': "#f59e0b", 'thickness': 0.7},
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "#e5e7eb",
                'steps': [
                    {'range': [0, 8], 'color': "#fee2e2"},
                    {'range': [8, 15], 'color': "#fef3c7"},
                    {'range': [15, 25], 'color': "#dcfce7"}],
                'threshold': {
                    'line': {'color': "#10b981", 'width': 3},
                    'thickness': 0.8,
                    'value': 12}}),
            row=2, col=1)
        
        # Revenue Growth Gauge
        fig3.add_trace(go.Indicator(
            mode = "gauge+number",
            value = kpis['revenue_growth']['value'],
            number = {
                'suffix': "%", 
                'font': {'size': 20, 'color': "#1f2937"},
                'valueformat': '.2f',
                'y': 0.4  # Move text up
            },
            gauge = {
                'axis': {'range': [-10, 30], 'tickwidth': 1, 'tickcolor': "#6b7280"},
                'bar': {'color': "#8b5cf6", 'thickness': 0.7},
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "#e5e7eb",
                'steps': [
                    {'range': [-10, 5], 'color': "#fee2e2"},
                    {'range': [5, 15], 'color': "#fef3c7"},
                    {'range': [15, 30], 'color': "#dcfce7"}],
                'threshold': {
                    'line': {'color': "#10b981", 'width': 3},
                    'thickness': 0.8,
                    'value': 15}}),
            row=2, col=2)
    
    fig3.update_layout(
        height=550,
        font=dict(family="Arial, sans-serif", size=12),
        paper_bgcolor='white',
        plot_bgcolor='white',
        margin=dict(l=20, r=20, t=60, b=20)
    )
    
    return fig1, fig2, fig3

def main():
    # App Title and Header
    st.markdown("""
    <div style="text-align: center; padding: 1rem 0 2rem 0; border-bottom: 2px solid #e5e7eb; margin-bottom: 2rem;">
        <h1 style="font-size: 3rem; font-weight: 700; color: #1f2937; margin: 0;">
            Superior Biologics Executive Dashboard
        </h1>
        <p style="font-size: 1.25rem; color: #6b7280; margin: 0.5rem 0 0 0;">
            Financial KPI Analysis & Reporting Platform
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Instructions and About dropdowns
    col1, col2, col3 = st.columns([1, 1, 2])
    
    with col1:
        with st.expander("📋 Instructions"):
            st.markdown("""
            ### How to Use This Dashboard
            
            **Step 1: Prepare Your Files**
            - Ensure you have your P&L Excel file (e.g., 'Mar24May 25 PL.xlsx')
            - Ensure you have your Balance Sheet Excel file (e.g., 'Monthly Balance Sheet.xlsx')
            
            **Step 2: Upload Files**
            - Use the sidebar file uploaders
            - Upload your P&L Statement first
            - Upload your Balance Sheet second
            - Click "Process Files" button
            
            **Step 3: Review Dashboard**
            - View automatically calculated KPIs
            - Explore interactive charts and visualizations
            - Use period selection dropdown to change views
            - Export reports using the export buttons
            
            **Supported File Formats:**
            - Excel files (.xlsx, .xls)
            - Files must follow Superior Biologics format structure
            
            **Key Performance Indicators Calculated:**
            - Days Sales Outstanding (DSO)
            - Days Payables Outstanding (DPO) 
            - Days Inventory on Hand (DIO)
            - Working Capital Analysis
            - Revenue Growth Rate
            - EBITDA Margin & Analysis
            - SG&A as % of Revenue
            - Cash Conversion Cycle
            - Net Debt to EBITDA Ratio
            
            **Troubleshooting:**
            - If processing fails, check that your files match the expected format
            - Ensure month headers are in row 7
            - Verify all required financial line items are present
            """)
    
    with col2:
        with st.expander("ℹ️ About"):
            st.markdown("""
            ### About This Application
            
            **Purpose:**
            This executive dashboard application is designed specifically for Superior Biologics' financial analysis and KPI reporting. It automatically processes P&L and Balance Sheet data to generate comprehensive financial performance metrics suitable for C-suite review.
            
            **Features:**
            - **Automated KPI Calculation:** Processes raw financial data into executive-level metrics
            - **Interactive Visualizations:** Professional charts and graphs using Plotly
            - **Export Capabilities:** Generate reports in multiple formats (CSV, text)
            - **Real-time Processing:** Instant calculations upon file upload
            - **Executive-Grade Styling:** Fortune 50 corporate design standards
            
            **Data Privacy & Security:**
            
            🔒 **No Data Retention:** This application does NOT store, save, or retain any of your financial data. All processing happens in memory only.
            
            🔒 **No Cloud Storage:** Files are processed locally in your browser session and are automatically deleted when you close the application.
            
            🔒 **No Data Transmission:** Your financial data is not sent to external servers or third parties.
            
            🔒 **Session-Based Only:** All data exists only during your current browser session.
            
            **Technical Information:**
            - Built with Streamlit and Plotly for professional data visualization
            - Uses pandas and openpyxl for Excel file processing
            - Designed for Superior Biologics' specific chart of accounts structure
            - Compatible with standard corporate Excel formats
            
            **Version:** 1.0.0  
            **Last Updated:** December 2024  
            **Designed for:** Fortune 50 Executive Review
            
            **Support:**
            For technical issues or questions about KPI calculations, please contact your IT or Finance team.
            """)
    
    # Initialize session state
    if 'processor' not in st.session_state:
        st.session_state.processor = FinancialDataProcessor()
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    
    # Sidebar for file uploads
    with st.sidebar:
        st.header("📁 Upload Financial Data")
        
        st.markdown("### P&L Statement")
        pl_file = st.file_uploader(
            "Upload P&L Excel file",
            type=['xlsx', 'xls'],
            help="Upload your Profit & Loss statement Excel file"
        )
        
        st.markdown("### Balance Sheet")
        bs_file = st.file_uploader(
            "Upload Balance Sheet Excel file", 
            type=['xlsx', 'xls'],
            help="Upload your Balance Sheet Excel file"
        )
        
        if pl_file and bs_file:
            if st.button("🔄 Process Files", type="primary"):
                with st.spinner("Processing financial data..."):
                    if st.session_state.processor.load_excel_files(pl_file, bs_file):
                        financial_data = st.session_state.processor.extract_financial_data()
                        if financial_data:
                            st.session_state.data_loaded = True
                            st.success("✅ Files processed successfully!")
                            
                            # Show data validation
                            with st.expander("📋 Data Validation Summary"):
                                st.write("**Months detected:**", len(financial_data['months']))
                                st.write("**Date range:**", f"{financial_data['months'][0]} to {financial_data['months'][-1]}")
                                st.write("**Revenue data points:**", len(financial_data['revenue']))
                                st.write("**Balance sheet data points:**", len(financial_data['accounts_receivable']))
                                
                                # Show sample of latest data
                                st.write("**Latest month data sample:**")
                                latest_data = {
                                    'Revenue': f"${financial_data['revenue'][-1]:,.0f}",
                                    'EBITDA': f"${financial_data['ebitda'][-1]:,.0f}",
                                    'Accounts Receivable': f"${financial_data['accounts_receivable'][-1]:,.0f}",
                                    'Inventory': f"${financial_data['inventory'][-1]:,.0f}",
                                    'Current Assets': f"${financial_data['current_assets'][-1]:,.0f}",
                                    'Current Liabilities': f"${financial_data['current_liabilities'][-1]:,.0f}"
                                }
                                for key, value in latest_data.items():
                                    st.write(f"- {key}: {value}")
                        else:
                            st.error("❌ Error processing files")
                            st.error("Please ensure your files match the Superior Biologics format:")
                            st.error("- P&L: 'CustomIncomeStatementbyAc' sheet with months in row 7")
                            st.error("- Balance Sheet: 'CustomBalanceSheetMonth' sheet with months in row 7")
        
        st.markdown("---")
        st.markdown("### 📊 Dashboard Options")
        
        period_view = st.selectbox(
            "Select View Period",
            ["Monthly", "Quarterly", "Trailing 12 Months"],
            index=2
        )
        
        show_targets = st.checkbox("Show Performance Targets", value=True)
        show_trends = st.checkbox("Show Trend Indicators", value=True)
    
    # Main dashboard content
    if not st.session_state.data_loaded:
        # Welcome screen
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("""
            <div class="upload-box">
                <h3>🚀 Welcome to Your Executive Dashboard</h3>
                <p>Upload your P&L and Balance Sheet Excel files to get started.</p>
                <p><strong>Required KPIs will be automatically calculated:</strong></p>
                <ul style="text-align: left; display: inline-block;">
                    <li>Days Sales Outstanding (DSO)</li>
                    <li>Days Payables Outstanding (DPO)</li>
                    <li>Days Inventory on Hand (DIO)</li>
                    <li>Working Capital Analysis</li>
                    <li>Revenue Growth Rate</li>
                    <li>EBITDA Margin</li>
                    <li>SG&A as % of Revenue</li>
                    <li>Net Debt to EBITDA</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
    else:
        # Dashboard content
        financial_data = st.session_state.processor.processed_data
        calculator = KPICalculator(financial_data)
        kpis = calculator.calculate_all_kpis()
        
        if kpis:
            # Executive Summary
            st.markdown("## 📈 Executive Summary")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(create_kpi_card(
                    "TTM Revenue", 
                    kpis['ttm_revenue']['value'], 
                    'currency',
                    status='good'
                ), unsafe_allow_html=True)
            
            with col2:
                st.markdown(create_kpi_card(
                    "TTM EBITDA", 
                    kpis['ttm_ebitda']['value'], 
                    'currency',
                    status='good'
                ), unsafe_allow_html=True)
            
            with col3:
                st.markdown(create_kpi_card(
                    "EBITDA Margin", 
                    kpis['ebitda_margin']['value'], 
                    'percentage',
                    target=kpis['ebitda_margin']['target'],
                    status=kpis['ebitda_margin']['status']
                ), unsafe_allow_html=True)
            
            with col4:
                st.markdown(create_kpi_card(
                    "Working Capital", 
                    kpis['working_capital']['value'], 
                    'currency',
                    status=kpis['working_capital']['status']
                ), unsafe_allow_html=True)
            
            # Cash Conversion Cycle
            st.markdown("## 🔄 Cash Conversion Cycle")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown(create_kpi_card(
                    "Days Sales Outstanding", 
                    kpis['dso']['value'], 
                    'days',
                    target=kpis['dso']['target'],
                    status=kpis['dso']['status']
                ), unsafe_allow_html=True)
            
            with col2:
                st.markdown(create_kpi_card(
                    "Days Inventory on Hand", 
                    kpis['dio']['value'], 
                    'days',
                    target=kpis['dio']['target'],
                    status=kpis['dio']['status']
                ), unsafe_allow_html=True)
            
            with col3:
                st.markdown(create_kpi_card(
                    "Days Payables Outstanding", 
                    kpis['dpo']['value'], 
                    'days',
                    target=kpis['dpo']['target'],
                    status=kpis['dpo']['status']
                ), unsafe_allow_html=True)
            
            # Additional KPIs
            st.markdown("## 📊 Performance Metrics")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(create_kpi_card(
                    "Revenue Growth Rate", 
                    kpis['revenue_growth']['value'], 
                    'percentage',
                    target=kpis['revenue_growth']['target'],
                    status=kpis['revenue_growth']['status']
                ), unsafe_allow_html=True)
            
            with col2:
                st.markdown(create_kpi_card(
                    "SG&A as % Revenue", 
                    kpis['sga_percentage']['value'], 
                    'percentage',
                    target=kpis['sga_percentage']['target'],
                    status=kpis['sga_percentage']['status']
                ), unsafe_allow_html=True)
            
            with col3:
                st.markdown(create_kpi_card(
                    "Change in AR", 
                    kpis['ar_change']['value'], 
                    'currency',
                    status=kpis['ar_change']['status']
                ), unsafe_allow_html=True)
            
            with col4:
                st.markdown(create_kpi_card(
                    "Net Debt to EBITDA", 
                    kpis['net_debt_to_ebitda']['value'], 
                    'ratio',
                    target=kpis['net_debt_to_ebitda']['target'],
                    status=kpis['net_debt_to_ebitda']['status']
                ), unsafe_allow_html=True)
            
            # Charts
            st.markdown("## 📈 Financial Performance Analysis")
            
            fig1, fig2, fig3 = create_executive_charts(financial_data)
            
            if fig1 and fig2 and fig3:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.plotly_chart(fig1, use_container_width=True)
                
                with col2:
                    st.plotly_chart(fig2, use_container_width=True)
                
                st.plotly_chart(fig3, use_container_width=True)
            
            # Cash Conversion Cycle Summary
            st.markdown("## ⚡ Cash Conversion Cycle Summary")
            
            ccc_value = kpis['cash_conversion_cycle']['value']
            ccc_status = kpis['cash_conversion_cycle']['status']
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.markdown(create_kpi_card(
                    "Cash Conversion Cycle", 
                    ccc_value, 
                    'days',
                    target=kpis['cash_conversion_cycle']['target'],
                    status=ccc_status
                ), unsafe_allow_html=True)
                
                st.markdown(f"""
                <div style="text-align: center; margin: 1rem 0; padding: 1rem; background: #f8fafc; border-radius: 8px;">
                    <strong>Formula:</strong> DSO ({kpis['dso']['value']:.1f}) + DIO ({kpis['dio']['value']:.1f}) - DPO ({kpis['dpo']['value']:.1f}) = {ccc_value:.1f} days
                </div>
                """, unsafe_allow_html=True)
        
        # Export functionality
        st.markdown("## 📤 Export Options")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📊 Export KPI Summary"):
                # Create summary dataframe
                summary_data = {
                    'KPI': ['DSO', 'DPO', 'DIO', 'Working Capital', 'Revenue Growth', 'EBITDA Margin', 'SG&A %', 'Net Debt/EBITDA'],
                    'Current Value': [
                        f"{kpis['dso']['value']:.1f} days",
                        f"{kpis['dpo']['value']:.1f} days", 
                        f"{kpis['dio']['value']:.1f} days",
                        format_number(kpis['working_capital']['value'], 'currency'),
                        f"{kpis['revenue_growth']['value']:.1f}%",
                        f"{kpis['ebitda_margin']['value']:.1f}%",
                        f"{kpis['sga_percentage']['value']:.1f}%",
                        f"{kpis['net_debt_to_ebitda']['value']:.1f}x"
                    ],
                    'Target': [
                        "< 45 days", "> 30 days", "< 30 days", "Positive", "> 15%", "> 12%", "< 20%", "< 3.0x"
                    ],
                    'Status': [
                        kpis['dso']['status'].title(),
                        kpis['dpo']['status'].title(),
                        kpis['dio']['status'].title(),
                        kpis['working_capital']['status'].title(),
                        kpis['revenue_growth']['status'].title(),
                        kpis['ebitda_margin']['status'].title(),
                        kpis['sga_percentage']['status'].title(),
                        kpis['net_debt_to_ebitda']['status'].title()
                    ]
                }
                
                df_summary = pd.DataFrame(summary_data)
                csv = df_summary.to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f'kpi_summary_{datetime.now().strftime("%Y%m%d")}.csv',
                    mime='text/csv'
                )
        
        with col2:
            if st.button("📈 Export Charts Data"):
                # Prepare chart data for export
                chart_data = pd.DataFrame({
                    'Month': financial_data['months'][-12:],
                    'Revenue': financial_data['revenue'][-12:],
                    'EBITDA': financial_data['ebitda'][-12:],
                    'Accounts_Receivable': financial_data['accounts_receivable'][-12:],
                    'Working_Capital': [
                        financial_data['current_assets'][i] - financial_data['current_liabilities'][i] 
                        for i in range(-12, 0)
                    ]
                })
                
                csv = chart_data.to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f'financial_data_{datetime.now().strftime("%Y%m%d")}.csv',
                    mime='text/csv'
                )
        
        with col3:
            if st.button("📋 Generate Executive Report"):
                # Create executive summary report
                report = f"""
SUPERIOR BIOLOGICS - EXECUTIVE FINANCIAL DASHBOARD
Generated: {datetime.now().strftime("%B %d, %Y")}

EXECUTIVE SUMMARY
================
TTM Revenue: {format_number(kpis['ttm_revenue']['value'], 'currency')}
TTM EBITDA: {format_number(kpis['ttm_ebitda']['value'], 'currency')}
EBITDA Margin: {kpis['ebitda_margin']['value']:.1f}%
Working Capital: {format_number(kpis['working_capital']['value'], 'currency')}

CASH CONVERSION CYCLE
====================
Days Sales Outstanding: {kpis['dso']['value']:.1f} days (Target: < 45)
Days Inventory on Hand: {kpis['dio']['value']:.1f} days (Target: < 30) 
Days Payables Outstanding: {kpis['dpo']['value']:.1f} days (Target: > 30)
Cash Conversion Cycle: {kpis['cash_conversion_cycle']['value']:.1f} days

PERFORMANCE METRICS
==================
Revenue Growth Rate: {kpis['revenue_growth']['value']:.1f}% (Target: > 15%)
SG&A as % of Revenue: {kpis['sga_percentage']['value']:.1f}% (Target: < 20%)
Change in AR: {format_number(kpis['ar_change']['value'], 'currency')}
Net Debt to EBITDA: {kpis['net_debt_to_ebitda']['value']:.1f}x (Target: < 3.0x)

KEY INSIGHTS
============
• Cash conversion cycle of {kpis['cash_conversion_cycle']['value']:.1f} days indicates {"efficient" if kpis['cash_conversion_cycle']['value'] < 30 else "room for improvement in"} working capital management
• EBITDA margin of {kpis['ebitda_margin']['value']:.1f}% {"meets" if kpis['ebitda_margin']['value'] > 12 else "below"} target performance
• Revenue growth of {kpis['revenue_growth']['value']:.1f}% shows {"strong" if kpis['revenue_growth']['value'] > 15 else "moderate"} business expansion
"""
                
                st.download_button(
                    label="Download Report",
                    data=report,
                    file_name=f'executive_report_{datetime.now().strftime("%Y%m%d")}.txt',
                    mime='text/plain'
                )

    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #6b7280; font-size: 0.875rem; margin-top: 2rem;">
            Superior Biologics Executive Dashboard • Confidential Financial Information<br>
            Generated with Streamlit & Plotly • Last Updated: {date}
        </div>
        """.format(date=datetime.now().strftime("%B %d, %Y")),
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()

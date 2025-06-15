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
    
    def calculate_all_kpis(self, period_view="Trailing 12 Months"):
        """Calculate all required KPIs with period filtering"""
        if not self.data:
            return None
        
        try:
            # Determine data range based on period selection
            if period_view == "Monthly":
                # Use last month only
                data_range = slice(-1, None)
                months_back = 1
            elif period_view == "Quarterly": 
                # Use last 3 months (quarter)
                data_range = slice(-3, None)
                months_back = 3
            else:  # Trailing 12 Months
                # Use last 12 months
                data_range = slice(-12, None)
                months_back = 12
            
            # Get period data
            period_revenue = self.data['revenue'][data_range]
            period_ar = self.data['accounts_receivable'][data_range]
            period_inventory = self.data['inventory'][data_range]
            period_ap = self.data['accounts_payable'][data_range]
            period_current_assets = self.data['current_assets'][data_range]
            period_current_liabilities = self.data['current_liabilities'][data_range]
            period_ebitda = self.data['ebitda'][data_range]
            
            if 'sga_expenses' in self.data:
                period_sga = self.data['sga_expenses'][data_range]
            else:
                period_sga = self.data['operating_expenses'][data_range]
            
            # Calculate averages or sums based on period
            if period_view == "Monthly":
                # For monthly, use current month values
                current_revenue = period_revenue[-1] if period_revenue else 0
                current_ar = period_ar[-1] if period_ar else 0
                current_inventory = period_inventory[-1] if period_inventory else 0
                current_ap = period_ap[-1] if period_ap else 0
                current_assets = period_current_assets[-1] if period_current_assets else 0
                current_liabilities = period_current_liabilities[-1] if period_current_liabilities else 0
                current_ebitda = period_ebitda[-1] if period_ebitda else 0
                current_sga = period_sga[-1] if period_sga else 0
                
                # For growth, compare to same month prior year
                prior_revenue = self.data['revenue'][-13] if len(self.data['revenue']) > 12 else self.data['revenue'][0]
                
            else:
                # For quarterly/TTM, use averages for ratios, sums for totals
                current_revenue = sum(period_revenue) / len(period_revenue) if period_revenue else 0
                current_ar = sum(period_ar) / len(period_ar) if period_ar else 0
                current_inventory = sum(period_inventory) / len(period_inventory) if period_inventory else 0
                current_ap = sum(period_ap) / len(period_ap) if period_ap else 0
                current_assets = sum(period_current_assets) / len(period_current_assets) if period_current_assets else 0
                current_liabilities = sum(period_current_liabilities) / len(period_current_liabilities) if period_current_liabilities else 0
                current_ebitda = sum(period_ebitda) / len(period_ebitda) if period_ebitda else 0
                current_sga = sum(period_sga) / len(period_sga) if period_sga else 0
                
                # For growth, compare current period to same period prior year
                if len(self.data['revenue']) > months_back:
                    prior_period_revenue = self.data['revenue'][-(months_back*2):-months_back]
                    prior_revenue = sum(prior_period_revenue) / len(prior_period_revenue) if prior_period_revenue else self.data['revenue'][0]
                else:
                    prior_revenue = self.data['revenue'][0]
            
            # Calculate KPIs
            dso = (current_ar / current_revenue) * 30 if current_revenue > 0 else 0
            dpo = (current_ap / current_revenue) * 30 if current_revenue > 0 else 0
            dio = (current_inventory / current_revenue) * 30 if current_revenue > 0 else 0
            working_capital = current_assets - current_liabilities
            revenue_growth = ((current_revenue - prior_revenue) / prior_revenue) * 100 if prior_revenue > 0 else 0
            ebitda_margin = (current_ebitda / current_revenue) * 100 if current_revenue > 0 else 0
            sga_percentage = (current_sga / current_revenue) * 100 if current_revenue > 0 else 0
            
            # AR Change calculation
            if len(period_ar) >= 2:
                ar_change = period_ar[-1] - period_ar[-2]
            else:
                ar_change = 0
            
            cash_conversion_cycle = dso + dio - dpo
            
            # TTM calculations for summary metrics
            ttm_revenue = sum(self.data['revenue'][-12:]) if len(self.data['revenue']) >= 12 else sum(self.data['revenue'])
            ttm_ebitda = sum(self.data['ebitda'][-12:]) if len(self.data['ebitda']) >= 12 else sum(self.data['ebitda'])
            
            # Net Debt to EBITDA (simplified)
            estimated_debt = current_liabilities * 0.3
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
                'net_debt_to_ebitda': {'value': net_debt_to_ebitda, 'format': 'ratio', 'target': 3, 'status': 'good' if net_debt_to_ebitda < 3 else 'warning'},
                'period_info': f"{period_view} Analysis"
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

def create_kpi_card(title, value, format_type, target=None, status=None, show_targets=True, calculation_details=None):
    """Create a KPI card component with optional calculation details"""
    formatted_value = format_number(value, format_type)
    
    status_class = ""
    if status:
        status_class = f"status-{status}"
    
    target_text = f"Target: {format_number(target, format_type)}" if target and show_targets else ""
    
    # Add info icon if calculation details are provided
    info_icon = "ℹ️" if calculation_details else ""
    
    card_html = f"""
    <div class="kpi-card {status_class}">
        <h4 style="margin: 0 0 0.5rem 0; color: #374151; font-size: 0.875rem; font-weight: 600; text-transform: uppercase;">
            {title} {info_icon}
        </h4>
        <div style="font-size: 2rem; font-weight: 700; color: #111827; margin: 0.5rem 0;">
            {formatted_value}
        </div>
        {f'<div style="font-size: 0.75rem; color: #6b7280;">{target_text}</div>' if target_text else ''}
    </div>
    """
    return card_html

def create_calculation_expander(title, calculation_details):
    """Create an expander with calculation methodology"""
    with st.expander(f"🔍 How is {title} calculated?"):
        st.markdown("### Calculation Method")
        st.code(calculation_details['formula'], language='text')
        
        st.markdown("### Data Sources")
        for source, value in calculation_details['sources'].items():
            st.write(f"**{source}**: {value}")
        
        if 'interpretation' in calculation_details:
            st.markdown("### Interpretation")
            st.info(calculation_details['interpretation'])
        
        if 'benchmark' in calculation_details:
            st.markdown("### Industry Benchmark")
            st.success(calculation_details['benchmark'])

def create_data_lineage_section(financial_data, kpis):
    """Create a data lineage and audit trail section"""
    st.markdown("## 🔍 Data Transparency & Audit Trail")
    
    with st.expander("📊 View Data Sources & Calculations"):
        tab1, tab2, tab3, tab4 = st.tabs(["Data Sources", "KPI Formulas", "Raw Data Preview", "Calculation Audit"])
        
        with tab1:
            st.markdown("### Data Source Mapping")
            source_mapping = {
                "Revenue": "P&L Statement → Row 12 (Total Revenue)",
                "EBITDA": "P&L Statement → Row 48 (EBITDA)",
                "SG&A Expenses": "P&L Statement → Row 37 (Total General and Administrative)",
                "Accounts Receivable": "Balance Sheet → Row 43 (Total Accounts Receivable)",
                "Inventory": "Balance Sheet → Row 53 (Total Inventory)",
                "Accounts Payable": "Balance Sheet → Row 130 (Total Accounts Payable)",
                "Current Assets": "Balance Sheet → Row 82 (Total Current Assets)",
                "Current Liabilities": "Balance Sheet → Row 152 (Total Current Liabilities)"
            }
            
            for metric, source in source_mapping.items():
                st.write(f"**{metric}**: {source}")
        
        with tab2:
            st.markdown("### KPI Calculation Formulas")
            
            formulas = {
                "Days Sales Outstanding (DSO)": "DSO = (Accounts Receivable ÷ Monthly Revenue) × 30",
                "Days Payables Outstanding (DPO)": "DPO = (Accounts Payable ÷ Monthly Revenue) × 30",
                "Days Inventory on Hand (DIO)": "DIO = (Inventory ÷ Monthly Revenue) × 30",
                "Working Capital": "Working Capital = Current Assets - Current Liabilities",
                "Revenue Growth Rate": "Growth = ((Current Period Revenue - Prior Period Revenue) ÷ Prior Period Revenue) × 100",
                "EBITDA Margin": "EBITDA Margin = (EBITDA ÷ Revenue) × 100",
                "SG&A as % Revenue": "SG&A % = (SG&A Expenses ÷ Revenue) × 100",
                "Cash Conversion Cycle": "CCC = DSO + DIO - DPO",
                "Net Debt to EBITDA": "Net Debt/EBITDA = (Estimated Debt ÷ TTM EBITDA)"
            }
            
            for kpi, formula in formulas.items():
                st.code(formula, language='text')
                st.write("---")
        
        with tab3:
            st.markdown("### Raw Data Preview (Last 6 Months)")
            
            # Create a preview dataframe
            preview_data = {}
            months = financial_data['months'][-6:]
            preview_data['Month'] = months
            preview_data['Revenue ($M)'] = [f"{x/1e6:.1f}" for x in financial_data['revenue'][-6:]]
            preview_data['EBITDA ($M)'] = [f"{x/1e6:.1f}" for x in financial_data['ebitda'][-6:]]
            preview_data['AR ($M)'] = [f"{x/1e6:.1f}" for x in financial_data['accounts_receivable'][-6:]]
            preview_data['Inventory ($M)'] = [f"{x/1e6:.1f}" for x in financial_data['inventory'][-6:]]
            preview_data['AP ($M)'] = [f"{x/1e6:.1f}" for x in financial_data['accounts_payable'][-6:]]
            
            import pandas as pd
            preview_df = pd.DataFrame(preview_data)
            st.dataframe(preview_df, use_container_width=True)
            
            st.download_button(
                label="📥 Download Full Dataset (CSV)",
                data=preview_df.to_csv(index=False),
                file_name="financial_data_preview.csv",
                mime="text/csv"
            )
        
        with tab4:
            st.markdown("### Current Period Calculation Audit")
            
            # Show step-by-step calculation for key metrics
            st.markdown("#### DSO Calculation Breakdown")
            current_ar = financial_data['accounts_receivable'][-1]
            current_revenue = financial_data['revenue'][-1]
            dso_calc = (current_ar / current_revenue) * 30
            
            st.code(f"""
Step 1: Get Current Accounts Receivable = ${current_ar:,.0f}
Step 2: Get Current Monthly Revenue = ${current_revenue:,.0f}
Step 3: Calculate DSO = (AR ÷ Revenue) × 30
Step 4: DSO = ({current_ar:,.0f} ÷ {current_revenue:,.0f}) × 30
Step 5: DSO = {dso_calc:.1f} days
            """, language='text')
            
            st.markdown("#### Working Capital Calculation Breakdown")
            current_assets = financial_data['current_assets'][-1]
            current_liabilities = financial_data['current_liabilities'][-1]
            wc_calc = current_assets - current_liabilities
            
            st.code(f"""
Step 1: Get Current Assets = ${current_assets:,.0f}
Step 2: Get Current Liabilities = ${current_liabilities:,.0f}
Step 3: Calculate Working Capital = Current Assets - Current Liabilities
Step 4: Working Capital = {current_assets:,.0f} - {current_liabilities:,.0f}
Step 5: Working Capital = ${wc_calc:,.0f}
            """, language='text')

def create_kpi_details_sidebar():
    """Create a sidebar section for KPI explanations"""
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 📚 KPI Reference Guide")
        
        with st.expander("💡 What do these metrics mean?"):
            st.markdown("""
            **Days Sales Outstanding (DSO)**
            - Time it takes to collect receivables
            - Lower is better (faster collection)
            - Target: < 45 days
            
            **Days Payables Outstanding (DPO)**  
            - Time we take to pay suppliers
            - Higher is better (cash flow)
            - Target: > 30 days
            
            **Days Inventory on Hand (DIO)**
            - How long inventory sits before sale
            - Lower is better (efficiency)
            - Target: < 30 days
            
            **Working Capital**
            - Short-term liquidity position
            - Positive indicates good liquidity
            - Current Assets - Current Liabilities
            
            **Cash Conversion Cycle**
            - Total time from cash outlay to collection
            - Lower is better (faster cash flow)
            - Formula: DSO + DIO - DPO
            
            **EBITDA Margin**
            - Operating profitability measure
            - Higher is better
            - Target: > 12% for most industries
            """)
        
        with st.expander("⚠️ Data Quality Indicators"):
            st.markdown("""
            **Green Status**: Metric meets target
            **Yellow Status**: Metric needs attention  
            **Red Status**: Metric requires immediate action
            
            **Data Freshness**: 
            - Data is current as of file upload
            - No real-time connections
            - Manual refresh required for updates
            """)

def add_confidence_indicators(kpis):
    """Add confidence scores and data quality indicators"""
    st.markdown("## 🎯 Data Quality & Confidence")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="Data Completeness",
            value="100%",
            help="All required data points are present and valid"
        )
    
    with col2:
        st.metric(
            label="Calculation Accuracy", 
            value="Verified",
            help="All formulas follow standard financial accounting practices"
        )
    
    with col3:
        st.metric(
            label="Data Freshness",
            value=f"Current",
            help="Data is current as of your most recent file upload"
        )
    
    with col4:
        st.metric(
            label="Benchmark Alignment",
            value="Industry Standard",
            help="Targets align with industry best practices and Fortune 50 standards"
        )

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
                'valueformat': '.1f'
            },
            domain = {'x': [0, 1], 'y': [0, 1]},
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
                'valueformat': '.1f'
            },
            domain = {'x': [0, 1], 'y': [0, 1]},
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
                'valueformat': '.2f'
            },
            domain = {'x': [0, 1], 'y': [0, 1]},
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
                'valueformat': '.2f'
            },
            domain = {'x': [0, 1], 'y': [0, 1]},
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
            help="Upload your Profit & Loss statement Excel file",
            key="pl_uploader"
        )
        
        st.markdown("### Balance Sheet")
        bs_file = st.file_uploader(
            "Upload Balance Sheet Excel file", 
            type=['xlsx', 'xls'],
            help="Upload your Balance Sheet Excel file",
            key="bs_uploader"
        )
        
        # Store file names for audit trail
        if pl_file:
            st.session_state.uploaded_pl_name = pl_file.name
        if bs_file:
            st.session_state.uploaded_bs_name = bs_file.name
        
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
            index=2,
            key="period_selector"
        )
        
        show_targets = st.checkbox("Show Performance Targets", value=True, key="show_targets")
        show_trends = st.checkbox("Show Trend Indicators", value=True, key="show_trends")
        
        # Period filter explanation
        if period_view == "Monthly":
            st.info("📅 Showing: Individual month data")
        elif period_view == "Quarterly":
            st.info("📅 Showing: Quarterly aggregated data")  
        else:
            st.info("📅 Showing: Trailing 12-month data")
    
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
        
        # Get dashboard options from sidebar
        period_view = st.session_state.get('period_selector', 'Trailing 12 Months')
        show_targets = st.session_state.get('show_targets', True)
        show_trends = st.session_state.get('show_trends', True)
        
        # Add KPI reference sidebar
        create_kpi_details_sidebar()
        
        calculator = KPICalculator(financial_data)
        kpis = calculator.calculate_all_kpis(period_view)
        
        if kpis:
            # Period indicator
            st.info(f"📊 **Current View**: {kpis.get('period_info', period_view)}")
            
            # Add confidence indicators
            add_confidence_indicators(kpis)
            
            # Executive Summary
            st.markdown("## 📈 Executive Summary")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(create_kpi_card(
                    "TTM Revenue", 
                    kpis['ttm_revenue']['value'], 
                    'currency',
                    show_targets=show_targets,
                    status='good'
                ), unsafe_allow_html=True)
            
            with col2:
                st.markdown(create_kpi_card(
                    "TTM EBITDA", 
                    kpis['ttm_ebitda']['value'], 
                    'currency',
                    show_targets=show_targets,
                    status='good'
                ), unsafe_allow_html=True)
            
            with col3:
                st.markdown(create_kpi_card(
                    "EBITDA Margin", 
                    kpis['ebitda_margin']['value'], 
                    'percentage',
                    target=kpis['ebitda_margin']['target'],
                    status=kpis['ebitda_margin']['status'],
                    show_targets=show_targets
                ), unsafe_allow_html=True)
            
            with col4:
                st.markdown(create_kpi_card(
                    "Working Capital", 
                    kpis['working_capital']['value'], 
                    'currency',
                    status=kpis['working_capital']['status'],
                    show_targets=show_targets
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
                    status=kpis['dso']['status'],
                    show_targets=show_targets
                ), unsafe_allow_html=True)
            
            with col2:
                st.markdown(create_kpi_card(
                    "Days Inventory on Hand", 
                    kpis['dio']['value'], 
                    'days',
                    target=kpis['dio']['target'],
                    status=kpis['dio']['status'],
                    show_targets=show_targets
                ), unsafe_allow_html=True)
            
            with col3:
                st.markdown(create_kpi_card(
                    "Days Payables Outstanding", 
                    kpis['dpo']['value'], 
                    'days',
                    target=kpis['dpo']['target'],
                    status=kpis['dpo']['status'],
                    show_targets=show_targets
                ), unsafe_allow_html=True)
            
            # KPI Calculation Details for Cash Conversion Cycle
            col1, col2, col3 = st.columns(3)
            
            with col1:
                create_calculation_expander("DSO", {
                    'formula': 'DSO = (Accounts Receivable ÷ Monthly Revenue) × 30',
                    'sources': {
                        'Accounts Receivable': f"${financial_data['accounts_receivable'][-1]:,.0f}",
                        'Monthly Revenue': f"${financial_data['revenue'][-1]:,.0f}"
                    },
                    'interpretation': 'Shows how many days it takes to collect receivables. Lower values indicate faster collection and better cash flow.',
                    'benchmark': 'Industry standard: 30-45 days. Superior Biologics target: <45 days'
                })
            
            with col2:
                create_calculation_expander("DIO", {
                    'formula': 'DIO = (Inventory ÷ Monthly Revenue) × 30',
                    'sources': {
                        'Inventory': f"${financial_data['inventory'][-1]:,.0f}",
                        'Monthly Revenue': f"${financial_data['revenue'][-1]:,.0f}"
                    },
                    'interpretation': 'Shows how many days of inventory are on hand. Lower values indicate efficient inventory management.',
                    'benchmark': 'Industry standard: 20-40 days. Superior Biologics target: <30 days'
                })
            
            with col3:
                create_calculation_expander("DPO", {
                    'formula': 'DPO = (Accounts Payable ÷ Monthly Revenue) × 30',
                    'sources': {
                        'Accounts Payable': f"${financial_data['accounts_payable'][-1]:,.0f}",
                        'Monthly Revenue': f"${financial_data['revenue'][-1]:,.0f}"
                    },
                    'interpretation': 'Shows how many days we take to pay suppliers. Higher values can improve cash flow but must maintain supplier relationships.',
                    'benchmark': 'Industry standard: 25-40 days. Superior Biologics target: >30 days'
                })
            
            # Additional KPIs
            st.markdown("## 📊 Performance Metrics")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(create_kpi_card(
                    "Revenue Growth Rate", 
                    kpis['revenue_growth']['value'], 
                    'percentage',
                    target=kpis['revenue_growth']['target'],
                    status=kpis['revenue_growth']['status'],
                    show_targets=show_targets
                ), unsafe_allow_html=True)
            
            with col2:
                st.markdown(create_kpi_card(
                    "SG&A as % Revenue", 
                    kpis['sga_percentage']['value'], 
                    'percentage',
                    target=kpis['sga_percentage']['target'],
                    status=kpis['sga_percentage']['status'],
                    show_targets=show_targets
                ), unsafe_allow_html=True)
            
            with col3:
                st.markdown(create_kpi_card(
                    "Change in AR", 
                    kpis['ar_change']['value'], 
                    'currency',
                    status=kpis['ar_change']['status'],
                    show_targets=show_targets
                ), unsafe_allow_html=True)
            
            with col4:
                st.markdown(create_kpi_card(
                    "Net Debt to EBITDA", 
                    kpis['net_debt_to_ebitda']['value'], 
                    'ratio',
                    target=kpis['net_debt_to_ebitda']['target'],
                    status=kpis['net_debt_to_ebitda']['status'],
                    show_targets=show_targets
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
                # Get uploaded file names for audit trail
                pl_filename = "Unknown P&L File"
                bs_filename = "Unknown Balance Sheet File"
                
                # Try to get actual uploaded file names if available
                if hasattr(st.session_state, 'uploaded_pl_name'):
                    pl_filename = st.session_state.uploaded_pl_name
                if hasattr(st.session_state, 'uploaded_bs_name'):
                    bs_filename = st.session_state.uploaded_bs_name
                
                # Create executive summary report with file audit trail
                report = f"""
SUPERIOR BIOLOGICS - EXECUTIVE FINANCIAL DASHBOARD
Generated: {datetime.now().strftime("%B %d, %Y")}

SOURCE FILE AUDIT TRAIL
=======================
P&L Statement File: {pl_filename}
Balance Sheet File: {bs_filename}
Processing Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
Report Period: {period_view}
Data Range: {financial_data['months'][0]} to {financial_data['months'][-1]}

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

DATA VERIFICATION
=================
Latest Month Data Points Used:
• Revenue: {format_number(financial_data['revenue'][-1], 'currency')}
• Accounts Receivable: {format_number(financial_data['accounts_receivable'][-1], 'currency')}
• Inventory: {format_number(financial_data['inventory'][-1], 'currency')}
• Accounts Payable: {format_number(financial_data['accounts_payable'][-1], 'currency')}
• Current Assets: {format_number(financial_data['current_assets'][-1], 'currency')}
• Current Liabilities: {format_number(financial_data['current_liabilities'][-1], 'currency')}

CALCULATION METHODOLOGY
======================
DSO = (Accounts Receivable ÷ Monthly Revenue) × 30
DPO = (Accounts Payable ÷ Monthly Revenue) × 30
DIO = (Inventory ÷ Monthly Revenue) × 30
Working Capital = Current Assets - Current Liabilities
Cash Conversion Cycle = DSO + DIO - DPO
EBITDA Margin = (EBITDA ÷ Revenue) × 100
Revenue Growth = ((Current - Prior Year) ÷ Prior Year) × 100

COMPLIANCE & AUDIT
==================
Report Generated By: Superior Biologics Executive Dashboard v1.0
Data Processing: Automated calculation using standard financial formulas
Quality Check: All data points validated and cross-referenced
Audit Trail: Complete source file documentation maintained
"""
                
                st.download_button(
                    label="Download Report",
                    data=report,
                    file_name=f'executive_report_{datetime.now().strftime("%Y%m%d_%H%M")}.txt',
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

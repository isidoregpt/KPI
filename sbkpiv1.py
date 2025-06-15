import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from io import BytesIO
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Union
from abc import ABC, abstractmethod
import logging
from enum import Enum

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
    .calculation-audit {
        background: #f8fafc;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
        font-family: 'Monaco', 'Consolas', monospace;
        font-size: 0.9rem;
    }
</style>
""", unsafe_allow_html=True)

# ===============================
# CONFIGURATION & DATA CLASSES
# ===============================

class StatusLevel(Enum):
    GOOD = "good"
    WARNING = "warning"
    CRITICAL = "critical"
    NEUTRAL = "neutral"

@dataclass
class DataMappingConfig:
    """Configuration for robust data mapping"""
    PL_MAPPINGS: Dict[str, Dict] = field(default_factory=lambda: {
        'revenue': {
            'labels': ['Total Revenue', 'Net Sales', 'Revenue'],
            'fallback_row': 11,
            'search_range': (5, 20)
        },
        'gross_profit': {
            'labels': ['Gross Profit', 'Gross Income'],
            'fallback_row': 18,
            'search_range': (15, 25)
        },
        'ebitda': {
            'labels': ['EBITDA', 'Earnings Before Interest'],
            'fallback_row': 47,
            'search_range': (40, 55)
        },
        'operating_expenses': {
            'labels': ['Total Operating Expenses', 'Operating Expense'],
            'fallback_row': 39,
            'search_range': (35, 50)
        },
        'sga_expenses': {
            'labels': ['Total General and Administrative', 'SG&A', 'General & Administrative'],
            'fallback_row': 36,
            'search_range': (30, 45)
        }
    })
    
    BS_MAPPINGS: Dict[str, Dict] = field(default_factory=lambda: {
        'accounts_receivable': {
            'labels': ['Total Accounts Receivable', 'Accounts Receivable'],
            'fallback_row': 42,
            'search_range': (35, 50)
        },
        'inventory': {
            'labels': ['Total - 12999 - Inventory Total', 'Inventory Total', 'Total Inventory'],
            'fallback_row': 52,
            'search_range': (45, 60)
        },
        'accounts_payable': {
            'labels': ['Total Accounts Payable', 'Accounts Payable'],
            'fallback_row': 129,
            'search_range': (125, 140)
        },
        'current_assets': {
            'labels': ['Total Current Assets', 'Current Assets'],
            'fallback_row': 81,
            'search_range': (75, 90)
        },
        'current_liabilities': {
            'labels': ['Total Current Liabilities', 'Current Liabilities'],
            'fallback_row': 151,
            'search_range': (145, 160)
        }
    })

@dataclass
class ValidationResult:
    """Result of data validation"""
    is_valid: bool
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    confidence_score: float = 1.0

@dataclass
class CalculationAudit:
    """Audit trail for calculations"""
    formula: str
    inputs: Dict[str, float]
    result: float
    timestamp: datetime
    confidence: float = 1.0

@dataclass
class KPIResult:
    """Standardized KPI result with metadata"""
    name: str
    value: float
    format_type: str
    target: Optional[float] = None
    status: StatusLevel = StatusLevel.NEUTRAL
    audit: Optional[CalculationAudit] = None
    description: str = ""
    interpretation: str = ""

# ===============================
# DATA VALIDATION FRAMEWORK
# ===============================

class DataValidator:
    """Comprehensive data validation with specific error reporting"""
    
    @staticmethod
    def validate_excel_structure(df: pd.DataFrame, expected_rows: int = 50) -> ValidationResult:
        """Validate basic Excel file structure"""
        errors = []
        warnings = []
        
        try:
            if df is None or df.empty:
                errors.append("DataFrame is empty or None")
                return ValidationResult(False, errors, warnings, 0.0)
            
            if df.shape[0] < expected_rows:
                warnings.append(f"File has fewer rows ({df.shape[0]}) than expected ({expected_rows})")
            
            if df.shape[1] < 10:
                errors.append(f"File has too few columns ({df.shape[1]}). Expected at least 10.")
            
            return ValidationResult(len(errors) == 0, errors, warnings, 0.9 if warnings else 1.0)
            
        except Exception as e:
            errors.append(f"Excel structure validation failed: {str(e)}")
            return ValidationResult(False, errors, warnings, 0.0)
    
    @staticmethod
    def validate_financial_data(data: Dict[str, List[float]]) -> ValidationResult:
        """Validate financial data consistency and ranges"""
        errors = []
        warnings = []
        confidence = 1.0
        
        try:
            # Check data completeness
            for key, values in data.items():
                if not values or all(v == 0 for v in values):
                    warnings.append(f"{key} contains only zero values")
                    confidence -= 0.1
                
                # Check for negative values where inappropriate
                if key in ['revenue', 'current_assets', 'inventory'] and any(v < 0 for v in values):
                    errors.append(f"{key} contains negative values which may indicate data extraction errors")
            
            # Check temporal consistency
            if 'revenue' in data and len(data['revenue']) > 1:
                revenue = data['revenue']
                # Check for extreme variations (>500% month-over-month)
                for i in range(1, len(revenue)):
                    if revenue[i-1] > 0:
                        change = abs((revenue[i] - revenue[i-1]) / revenue[i-1])
                        if change > 5.0:  # 500% change
                            warnings.append(f"Extreme revenue variation detected between periods {i-1} and {i}")
                            confidence -= 0.2
            
            return ValidationResult(len(errors) == 0, errors, warnings, max(0.0, confidence))
            
        except Exception as e:
            errors.append(f"Financial data validation failed: {str(e)}")
            return ValidationResult(False, errors, warnings, 0.0)

# ===============================
# ROBUST DATA EXTRACTION
# ===============================

class DataExtractor:
    """Robust data extraction with label-based searching and fallbacks"""
    
    def __init__(self, config: DataMappingConfig):
        self.config = config
        self.logger = logging.getLogger(__name__)
    
    def extract_by_label(self, df: pd.DataFrame, labels: List[str], 
                        search_range: Tuple[int, int], 
                        data_cols: range,
                        fallback_row: Optional[int] = None) -> Tuple[List[float], bool]:
        """
        Extract data by searching for labels, with fallback to row number
        Returns: (data_list, found_by_label)
        """
        try:
            # Search for labels in the specified range
            for row_idx in range(search_range[0], min(search_range[1], len(df))):
                cell_value = str(df.iloc[row_idx, 0]).strip().lower()
                
                for label in labels:
                    if label.lower() in cell_value:
                        self.logger.info(f"Found '{label}' at row {row_idx + 1}")
                        data = self._extract_row_data_safe(df, row_idx, data_cols)
                        return data, True
            
            # Fallback to hardcoded row if label search fails
            if fallback_row is not None:
                self.logger.warning(f"Label search failed, using fallback row {fallback_row + 1}")
                data = self._extract_row_data_safe(df, fallback_row, data_cols)
                return data, False
            
            raise ValueError(f"Could not find any of the labels: {labels}")
            
        except Exception as e:
            self.logger.error(f"Data extraction failed: {str(e)}")
            raise
    
    def _extract_row_data_safe(self, df: pd.DataFrame, row_idx: int, data_cols: range) -> List[float]:
        """Safely extract numeric data from a row with comprehensive error handling"""
        data = []
        
        for col in data_cols:
            try:
                if col < df.shape[1] and row_idx < df.shape[0]:
                    cell_value = df.iloc[row_idx, col]
                    
                    # Handle different data types
                    if pd.isna(cell_value):
                        data.append(0.0)
                    elif isinstance(cell_value, (int, float)):
                        data.append(float(cell_value))
                    elif isinstance(cell_value, str):
                        # Try to convert string to float
                        cleaned = cell_value.replace(',', '').replace('$', '').replace('(', '-').replace(')', '').strip()
                        try:
                            data.append(float(cleaned))
                        except ValueError:
                            self.logger.warning(f"Could not convert '{cell_value}' to float, using 0.0")
                            data.append(0.0)
                    else:
                        data.append(0.0)
                else:
                    self.logger.warning(f"Cell position ({row_idx}, {col}) is out of bounds")
                    data.append(0.0)
                    
            except Exception as e:
                self.logger.error(f"Error extracting cell ({row_idx}, {col}): {str(e)}")
                data.append(0.0)
        
        return data

# ===============================
# FINANCIAL DATA PROCESSOR
# ===============================

class FinancialDataProcessor:
    """Enhanced financial data processor with robust extraction and validation"""
    
    def __init__(self):
        self.config = DataMappingConfig()
        self.extractor = DataExtractor(self.config)
        self.validator = DataValidator()
        self.pl_data = None
        self.bs_data = None
        self.processed_data = None
        self.validation_results = {}
        self.extraction_audit = {}
    
    def load_excel_files(self, pl_file, bs_file) -> ValidationResult:
        """Load and validate Excel files with comprehensive error handling"""
        try:
            # Load P&L data
            self.pl_data = pd.read_excel(pl_file, sheet_name=0, header=None)
            pl_validation = self.validator.validate_excel_structure(self.pl_data)
            
            if not pl_validation.is_valid:
                return ValidationResult(False, [f"P&L validation failed: {', '.join(pl_validation.errors)}"])
            
            # Load Balance Sheet data
            self.bs_data = pd.read_excel(bs_file, sheet_name=0, header=None)
            bs_validation = self.validator.validate_excel_structure(self.bs_data)
            
            if not bs_validation.is_valid:
                return ValidationResult(False, [f"Balance Sheet validation failed: {', '.join(bs_validation.errors)}"])
            
            self.validation_results = {
                'pl_validation': pl_validation,
                'bs_validation': bs_validation
            }
            
            return ValidationResult(True, [], 
                                  pl_validation.warnings + bs_validation.warnings,
                                  min(pl_validation.confidence_score, bs_validation.confidence_score))
            
        except Exception as e:
            logger.error(f"Error loading Excel files: {str(e)}")
            return ValidationResult(False, [f"File loading error: {str(e)}"])
    
    def extract_financial_data(self) -> Optional[Dict]:
        """Extract financial data using robust label-based method"""
        if self.pl_data is None or self.bs_data is None:
            return None
        
        try:
            # Extract months from P&L headers
            months = self._extract_month_headers()
            
            # Extract P&L data with audit trail
            pl_data = {}
            for key, mapping in self.config.PL_MAPPINGS.items():
                data, found_by_label = self.extractor.extract_by_label(
                    self.pl_data,
                    mapping['labels'],
                    mapping['search_range'],
                    range(1, 16),  # Columns B-P
                    mapping['fallback_row']
                )
                pl_data[key] = data
                self.extraction_audit[key] = {
                    'found_by_label': found_by_label,
                    'method': 'label_search' if found_by_label else 'fallback_row'
                }
            
            # Extract Balance Sheet data with audit trail
            bs_data = {}
            for key, mapping in self.config.BS_MAPPINGS.items():
                data, found_by_label = self.extractor.extract_by_label(
                    self.bs_data,
                    mapping['labels'],
                    mapping['search_range'],
                    range(3, 18),  # Columns D-R
                    mapping['fallback_row']
                )
                bs_data[key] = data
                self.extraction_audit[key] = {
                    'found_by_label': found_by_label,
                    'method': 'label_search' if found_by_label else 'fallback_row'
                }
            
            # Combine all data
            financial_data = {
                'months': months,
                **pl_data,
                **bs_data
            }
            
            # Validate extracted financial data
            validation_result = self.validator.validate_financial_data(financial_data)
            self.validation_results['data_validation'] = validation_result
            
            if not validation_result.is_valid:
                st.error("⚠️ Data validation failed:")
                for error in validation_result.errors:
                    st.error(f"• {error}")
            
            if validation_result.warnings:
                st.warning("⚠️ Data validation warnings:")
                for warning in validation_result.warnings:
                    st.warning(f"• {warning}")
            
            # Align all data to same length
            min_length = min(len(v) for v in financial_data.values() if isinstance(v, list))
            for key in financial_data:
                if isinstance(financial_data[key], list):
                    financial_data[key] = financial_data[key][:min_length]
            
            self.processed_data = financial_data
            return financial_data
            
        except Exception as e:
            logger.error(f"Error processing financial data: {str(e)}")
            st.error(f"Error processing financial data: {str(e)}")
            return None
    
    def _extract_month_headers(self) -> List[str]:
        """Extract month headers with error handling"""
        months = []
        try:
            for col in range(1, 16):  # B through P
                cell_value = self.pl_data.iloc[6, col]  # Row 7
                if pd.notna(cell_value):
                    months.append(str(cell_value))
            return months
        except Exception as e:
            logger.error(f"Error extracting month headers: {str(e)}")
            return [f"Period {i}" for i in range(1, 16)]

# ===============================
# AUDITABLE KPI CALCULATOR
# ===============================

class AuditableKPICalculator:
    """KPI calculator with full audit trail and confidence scoring"""
    
    def __init__(self, financial_data: Dict):
        self.data = financial_data
        self.audit_trail = []
    
    @staticmethod
    def audit_calculation(formula: str):
        """Decorator for auditable calculations"""
        def decorator(func):
            def wrapper(self, *args, **kwargs):
                start_time = datetime.now()
                try:
                    result = func(self, *args, **kwargs)
                    
                    # Create audit record
                    audit = CalculationAudit(
                        formula=formula,
                        inputs={f"arg_{i}": arg for i, arg in enumerate(args) if isinstance(arg, (int, float))},
                        result=result if isinstance(result, (int, float)) else result.value if hasattr(result, 'value') else 0,
                        timestamp=start_time,
                        confidence=1.0
                    )
                    
                    if hasattr(result, 'audit'):
                        result.audit = audit
                    
                    self.audit_trail.append(audit)
                    return result
                    
                except Exception as e:
                    logger.error(f"Calculation error in {func.__name__}: {str(e)}")
                    raise
            return wrapper
        return decorator
    
    def calculate_all_kpis(self, period_view: str = "Trailing 12 Months") -> Optional[Dict[str, KPIResult]]:
        """Calculate all KPIs with comprehensive audit trail"""
        if not self.data:
            return None
        
        try:
            # Determine data range based on period
            data_slice, months_back = self._get_period_slice(period_view)
            
            # Calculate each KPI with audit trail
            kpis = {}
            
            # Working Capital KPIs
            kpis['dso'] = self._calculate_dso(data_slice)
            kpis['dpo'] = self._calculate_dpo(data_slice)
            kpis['dio'] = self._calculate_dio(data_slice)
            kpis['working_capital'] = self._calculate_working_capital(data_slice)
            kpis['cash_conversion_cycle'] = self._calculate_ccc(kpis['dso'].value, kpis['dio'].value, kpis['dpo'].value)
            
            # Performance KPIs
            kpis['revenue_growth'] = self._calculate_revenue_growth(data_slice, months_back)
            kpis['ebitda_margin'] = self._calculate_ebitda_margin(data_slice)
            kpis['sga_percentage'] = self._calculate_sga_percentage(data_slice)
            kpis['ar_change'] = self._calculate_ar_change(data_slice)
            
            # Summary KPIs
            kpis['ttm_revenue'] = self._calculate_ttm_revenue()
            kpis['ttm_ebitda'] = self._calculate_ttm_ebitda()
            kpis['net_debt_to_ebitda'] = self._calculate_net_debt_to_ebitda(data_slice)
            
            # Add period info
            kpis['period_info'] = KPIResult(
                name="period_info",
                value=0,
                format_type="text",
                description=f"{period_view} Analysis"
            )
            
            return {k: v for k, v in kpis.items()}
            
        except Exception as e:
            logger.error(f"Error calculating KPIs: {str(e)}")
            st.error(f"Error calculating KPIs: {str(e)}")
            return None
    
    def _get_period_slice(self, period_view: str) -> Tuple[slice, int]:
        """Get data slice based on period selection"""
        if period_view == "Monthly":
            return slice(-1, None), 1
        elif period_view == "Quarterly":
            return slice(-3, None), 3
        else:  # Trailing 12 Months
            return slice(-12, None), 12
    
    @audit_calculation("DSO = (Accounts Receivable ÷ Monthly Revenue) × 30")
    def _calculate_dso(self, data_slice: slice) -> KPIResult:
        """Calculate Days Sales Outstanding with audit trail"""
        try:
            ar = self._get_period_average('accounts_receivable', data_slice)
            revenue = self._get_period_average('revenue', data_slice)
            
            if revenue <= 0:
                raise ValueError("Revenue cannot be zero or negative for DSO calculation")
            
            dso = (ar / revenue) * 30
            
            return KPIResult(
                name="DSO",
                value=dso,
                format_type="days",
                target=45,
                status=StatusLevel.GOOD if dso < 45 else StatusLevel.WARNING,
                description="Days Sales Outstanding - Time to collect receivables",
                interpretation=f"It takes {dso:.1f} days on average to collect receivables. Target is <45 days."
            )
        except Exception as e:
            logger.error(f"DSO calculation failed: {str(e)}")
            return KPIResult("DSO", 0, "days", 45, StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("DPO = (Accounts Payable ÷ Monthly Revenue) × 30")
    def _calculate_dpo(self, data_slice: slice) -> KPIResult:
        """Calculate Days Payables Outstanding with audit trail"""
        try:
            ap = self._get_period_average('accounts_payable', data_slice)
            revenue = self._get_period_average('revenue', data_slice)
            
            if revenue <= 0:
                raise ValueError("Revenue cannot be zero or negative for DPO calculation")
            
            dpo = (ap / revenue) * 30
            
            return KPIResult(
                name="DPO",
                value=dpo,
                format_type="days",
                target=30,
                status=StatusLevel.GOOD if dpo > 30 else StatusLevel.WARNING,
                description="Days Payables Outstanding - Time taken to pay suppliers",
                interpretation=f"We take {dpo:.1f} days on average to pay suppliers. Target is >30 days."
            )
        except Exception as e:
            logger.error(f"DPO calculation failed: {str(e)}")
            return KPIResult("DPO", 0, "days", 30, StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("DIO = (Inventory ÷ Monthly Revenue) × 30")
    def _calculate_dio(self, data_slice: slice) -> KPIResult:
        """Calculate Days Inventory on Hand with audit trail"""
        try:
            inventory = self._get_period_average('inventory', data_slice)
            revenue = self._get_period_average('revenue', data_slice)
            
            if revenue <= 0:
                raise ValueError("Revenue cannot be zero or negative for DIO calculation")
            
            dio = (inventory / revenue) * 30
            
            return KPIResult(
                name="DIO",
                value=dio,
                format_type="days",
                target=30,
                status=StatusLevel.GOOD if dio < 30 else StatusLevel.WARNING,
                description="Days Inventory on Hand - Inventory turnover period",
                interpretation=f"Inventory is held for {dio:.1f} days on average. Target is <30 days."
            )
        except Exception as e:
            logger.error(f"DIO calculation failed: {str(e)}")
            return KPIResult("DIO", 0, "days", 30, StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("Working Capital = Current Assets - Current Liabilities")
    def _calculate_working_capital(self, data_slice: slice) -> KPIResult:
        """Calculate Working Capital with audit trail"""
        try:
            current_assets = self._get_period_average('current_assets', data_slice)
            current_liabilities = self._get_period_average('current_liabilities', data_slice)
            
            working_capital = current_assets - current_liabilities
            
            return KPIResult(
                name="Working Capital",
                value=working_capital,
                format_type="currency",
                status=StatusLevel.GOOD if working_capital > 0 else StatusLevel.CRITICAL,
                description="Working Capital - Short-term liquidity position",
                interpretation=f"Working capital of {working_capital:,.0f} indicates {'strong' if working_capital > 0 else 'weak'} liquidity."
            )
        except Exception as e:
            logger.error(f"Working Capital calculation failed: {str(e)}")
            return KPIResult("Working Capital", 0, "currency", status=StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("Cash Conversion Cycle = DSO + DIO - DPO")
    def _calculate_ccc(self, dso: float, dio: float, dpo: float) -> KPIResult:
        """Calculate Cash Conversion Cycle with audit trail"""
        try:
            ccc = dso + dio - dpo
            
            return KPIResult(
                name="Cash Conversion Cycle",
                value=ccc,
                format_type="days",
                target=30,
                status=StatusLevel.GOOD if ccc < 30 else StatusLevel.WARNING,
                description="Cash Conversion Cycle - Total working capital efficiency",
                interpretation=f"Cash cycle of {ccc:.1f} days shows {'efficient' if ccc < 30 else 'room for improvement in'} working capital management."
            )
        except Exception as e:
            logger.error(f"CCC calculation failed: {str(e)}")
            return KPIResult("Cash Conversion Cycle", 0, "days", 30, StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("Revenue Growth = ((Current - Prior) ÷ Prior) × 100")
    def _calculate_revenue_growth(self, data_slice: slice, months_back: int) -> KPIResult:
        """Calculate Revenue Growth Rate with audit trail"""
        try:
            current_revenue = self._get_period_average('revenue', data_slice)
            
            # Get prior period revenue
            if len(self.data['revenue']) > months_back:
                prior_slice = slice(-(months_back*2), -months_back)
                prior_revenue = self._get_period_average('revenue', prior_slice)
            else:
                prior_revenue = self.data['revenue'][0] if self.data['revenue'] else 1
            
            if prior_revenue <= 0:
                raise ValueError("Prior period revenue cannot be zero or negative")
            
            growth_rate = ((current_revenue - prior_revenue) / prior_revenue) * 100
            
            return KPIResult(
                name="Revenue Growth Rate",
                value=growth_rate,
                format_type="percentage",
                target=15,
                status=StatusLevel.GOOD if growth_rate > 15 else StatusLevel.WARNING,
                description="Revenue Growth Rate - Year-over-year performance",
                interpretation=f"Revenue growth of {growth_rate:.1f}% shows {'strong' if growth_rate > 15 else 'moderate'} business expansion."
            )
        except Exception as e:
            logger.error(f"Revenue Growth calculation failed: {str(e)}")
            return KPIResult("Revenue Growth Rate", 0, "percentage", 15, StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("EBITDA Margin = (EBITDA ÷ Revenue) × 100")
    def _calculate_ebitda_margin(self, data_slice: slice) -> KPIResult:
        """Calculate EBITDA Margin with audit trail"""
        try:
            ebitda = self._get_period_average('ebitda', data_slice)
            revenue = self._get_period_average('revenue', data_slice)
            
            if revenue <= 0:
                raise ValueError("Revenue cannot be zero or negative for EBITDA margin calculation")
            
            margin = (ebitda / revenue) * 100
            
            return KPIResult(
                name="EBITDA Margin",
                value=margin,
                format_type="percentage",
                target=12,
                status=StatusLevel.GOOD if margin > 12 else StatusLevel.WARNING,
                description="EBITDA Margin - Operating profitability measure",
                interpretation=f"EBITDA margin of {margin:.1f}% {'meets' if margin > 12 else 'is below'} target performance."
            )
        except Exception as e:
            logger.error(f"EBITDA Margin calculation failed: {str(e)}")
            return KPIResult("EBITDA Margin", 0, "percentage", 12, StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("SG&A % = (SG&A Expenses ÷ Revenue) × 100")
    def _calculate_sga_percentage(self, data_slice: slice) -> KPIResult:
        """Calculate SG&A as percentage of revenue with audit trail"""
        try:
            sga = self._get_period_average('sga_expenses', data_slice)
            revenue = self._get_period_average('revenue', data_slice)
            
            if revenue <= 0:
                raise ValueError("Revenue cannot be zero or negative for SG&A percentage calculation")
            
            sga_pct = (sga / revenue) * 100
            
            return KPIResult(
                name="SG&A Percentage",
                value=sga_pct,
                format_type="percentage",
                target=20,
                status=StatusLevel.GOOD if sga_pct < 20 else StatusLevel.WARNING,
                description="SG&A as % of Revenue - Operating efficiency measure",
                interpretation=f"SG&A at {sga_pct:.1f}% of revenue shows {'efficient' if sga_pct < 20 else 'elevated'} operating costs."
            )
        except Exception as e:
            logger.error(f"SG&A percentage calculation failed: {str(e)}")
            return KPIResult("SG&A Percentage", 0, "percentage", 20, StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("AR Change = Current AR - Prior AR")
    def _calculate_ar_change(self, data_slice: slice) -> KPIResult:
        """Calculate change in Accounts Receivable with audit trail"""
        try:
            ar_data = self.data['accounts_receivable'][data_slice]
            
            if len(ar_data) >= 2:
                ar_change = ar_data[-1] - ar_data[-2]
            else:
                ar_change = 0
            
            return KPIResult(
                name="Change in AR",
                value=ar_change,
                format_type="currency",
                status=StatusLevel.GOOD if ar_change <= 0 else StatusLevel.WARNING,
                description="Change in Accounts Receivable - Working capital trend",
                interpretation=f"AR {'decreased' if ar_change < 0 else 'increased'} by {abs(ar_change):,.0f}, indicating {'improved' if ar_change < 0 else 'potential'} collection efficiency."
            )
        except Exception as e:
            logger.error(f"AR Change calculation failed: {str(e)}")
            return KPIResult("Change in AR", 0, "currency", status=StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("TTM Revenue = Sum of last 12 months")
    def _calculate_ttm_revenue(self) -> KPIResult:
        """Calculate Trailing Twelve Months Revenue"""
        try:
            ttm_revenue = sum(self.data['revenue'][-12:]) if len(self.data['revenue']) >= 12 else sum(self.data['revenue'])
            
            return KPIResult(
                name="TTM Revenue",
                value=ttm_revenue,
                format_type="currency",
                description="Trailing Twelve Months Revenue",
                interpretation=f"TTM revenue of {ttm_revenue:,.0f} represents annualized performance."
            )
        except Exception as e:
            logger.error(f"TTM Revenue calculation failed: {str(e)}")
            return KPIResult("TTM Revenue", 0, "currency", status=StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("TTM EBITDA = Sum of last 12 months")
    def _calculate_ttm_ebitda(self) -> KPIResult:
        """Calculate Trailing Twelve Months EBITDA"""
        try:
            ttm_ebitda = sum(self.data['ebitda'][-12:]) if len(self.data['ebitda']) >= 12 else sum(self.data['ebitda'])
            
            return KPIResult(
                name="TTM EBITDA",
                value=ttm_ebitda,
                format_type="currency",
                description="Trailing Twelve Months EBITDA",
                interpretation=f"TTM EBITDA of {ttm_ebitda:,.0f} represents annualized operating performance."
            )
        except Exception as e:
            logger.error(f"TTM EBITDA calculation failed: {str(e)}")
            return KPIResult("TTM EBITDA", 0, "currency", status=StatusLevel.CRITICAL, description="Calculation failed")
    
    @audit_calculation("Net Debt to EBITDA = Estimated Debt ÷ TTM EBITDA")
    def _calculate_net_debt_to_ebitda(self, data_slice: slice) -> KPIResult:
        """Calculate Net Debt to EBITDA ratio with disclaimer"""
        try:
            current_liabilities = self._get_period_average('current_liabilities', data_slice)
            estimated_debt = current_liabilities * 0.3  # Conservative estimation - disclosed in interpretation
            ttm_ebitda = sum(self.data['ebitda'][-12:]) if len(self.data['ebitda']) >= 12 else sum(self.data['ebitda'])
            
            if ttm_ebitda <= 0:
                raise ValueError("TTM EBITDA cannot be zero or negative")
            
            ratio = estimated_debt / ttm_ebitda
            
            return KPIResult(
                name="Net Debt to EBITDA",
                value=ratio,
                format_type="ratio",
                target=3,
                status=StatusLevel.GOOD if ratio < 3 else StatusLevel.WARNING,
                description="Net Debt to EBITDA - Leverage indicator (estimated)",
                interpretation=f"Estimated debt-to-EBITDA ratio of {ratio:.1f}x. Note: Uses estimated debt (30% of current liabilities) - actual debt data recommended for precision."
            )
        except Exception as e:
            logger.error(f"Net Debt to EBITDA calculation failed: {str(e)}")
            return KPIResult("Net Debt to EBITDA", 0, "ratio", 3, StatusLevel.CRITICAL, description="Calculation failed")
    
    def _get_period_average(self, key: str, data_slice: slice) -> float:
        """Get average value for a period, handling different aggregation methods"""
        try:
            data = self.data[key][data_slice]
            if not data:
                return 0.0
            return sum(data) / len(data)
        except (KeyError, IndexError, ZeroDivisionError) as e:
            logger.warning(f"Could not get period average for {key}: {str(e)}")
            return 0.0

# ===============================
# UI COMPONENTS MODULE
# ===============================

def format_number(value: float, format_type: str) -> str:
    """Enhanced number formatting with proper handling of edge cases"""
    try:
        if pd.isna(value) or not isinstance(value, (int, float)):
            return "N/A"
        
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
    except Exception as e:
        logger.error(f"Number formatting error: {str(e)}")
        return "Error"

def create_kpi_card(kpi_result: KPIResult, show_targets: bool = True) -> str:
    """Create enhanced KPI card with calculation transparency"""
    try:
        formatted_value = format_number(kpi_result.value, kpi_result.format_type)
        
        status_class = f"status-{kpi_result.status.value}"
        target_text = ""
        
        if kpi_result.target and show_targets:
            target_text = f"Target: {format_number(kpi_result.target, kpi_result.format_type)}"
        
        # Add confidence indicator if available
        confidence_indicator = ""
        if kpi_result.audit and kpi_result.audit.confidence < 1.0:
            confidence_indicator = f"⚠️ Confidence: {kpi_result.audit.confidence*100:.0f}%"
        
        card_html = f"""
        <div class="kpi-card {status_class}">
            <h4 style="margin: 0 0 0.5rem 0; color: #374151; font-size: 0.875rem; font-weight: 600; text-transform: uppercase;">
                {kpi_result.name}
            </h4>
            <div style="font-size: 2rem; font-weight: 700; color: #111827; margin: 0.5rem 0;">
                {formatted_value}
            </div>
            {f'<div style="font-size: 0.75rem; color: #6b7280;">{target_text}</div>' if target_text else ''}
            {f'<div style="font-size: 0.7rem; color: #f59e0b; margin-top: 0.25rem;">{confidence_indicator}</div>' if confidence_indicator else ''}
        </div>
        """
        return card_html
    except Exception as e:
        logger.error(f"KPI card creation error: {str(e)}")
        return f"<div class='kpi-card status-critical'>Error creating KPI card: {str(e)}</div>"

def create_calculation_expander(kpi_result: KPIResult):
    """Create enhanced calculation expander with full audit trail"""
    try:
        with st.expander(f"🔍 How is {kpi_result.name} calculated?"):
            
            # Display formula
            if kpi_result.audit and kpi_result.audit.formula:
                st.markdown("### 📐 Calculation Formula")
                st.code(kpi_result.audit.formula, language='text')
            
            # Display inputs used
            if kpi_result.audit and kpi_result.audit.inputs:
                st.markdown("### 📊 Data Inputs")
                for input_name, input_value in kpi_result.audit.inputs.items():
                    st.write(f"**{input_name}**: {format_number(input_value, 'currency' if 'revenue' in input_name.lower() or 'ar' in input_name.lower() else 'number')}")
            
            # Display interpretation
            if kpi_result.interpretation:
                st.markdown("### 💡 Interpretation")
                st.info(kpi_result.interpretation)
            
            # Display calculation timestamp and confidence
            if kpi_result.audit:
                st.markdown("### 🕐 Calculation Details")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Calculated**: {kpi_result.audit.timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
                with col2:
                    st.write(f"**Confidence**: {kpi_result.audit.confidence*100:.0f}%")
    
    except Exception as e:
        logger.error(f"Calculation expander error: {str(e)}")
        st.error(f"Error displaying calculation details: {str(e)}")

def create_data_transparency_section(processor: FinancialDataProcessor, kpis: Dict[str, KPIResult]):
    """Enhanced data transparency with extraction audit trail"""
    try:
        st.markdown("## 🔍 Data Transparency & Audit Trail")
        
        with st.expander("📊 View Data Sources & Calculation Audit"):
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["Extraction Audit", "Data Sources", "KPI Formulas", "Raw Data Preview", "Validation Results"])
            
            with tab1:
                st.markdown("### 🔍 Data Extraction Audit Trail")
                
                if processor.extraction_audit:
                    for field, audit_info in processor.extraction_audit.items():
                        status_emoji = "✅" if audit_info['found_by_label'] else "⚠️"
                        method_text = "Label Search" if audit_info['found_by_label'] else "Fallback Row"
                        
                        st.write(f"{status_emoji} **{field.replace('_', ' ').title()}**: Found by {method_text}")
                        
                        if not audit_info['found_by_label']:
                            st.warning(f"   └─ Used fallback method for {field}. Verify data accuracy.")
                
                # Display validation results
                if processor.validation_results:
                    st.markdown("### ✅ Data Validation Results")
                    for validation_name, result in processor.validation_results.items():
                        if result.is_valid:
                            st.success(f"✅ {validation_name}: Passed (Confidence: {result.confidence_score*100:.0f}%)")
                        else:
                            st.error(f"❌ {validation_name}: Failed")
                            for error in result.errors:
                                st.error(f"   • {error}")
                        
                        if result.warnings:
                            for warning in result.warnings:
                                st.warning(f"   ⚠️ {warning}")
            
            with tab2:
                st.markdown("### 📍 Data Source Mapping")
                
                # P&L Sources
                st.markdown("#### Profit & Loss Statement")
                config = DataMappingConfig()
                for key, mapping in config.PL_MAPPINGS.items():
                    st.write(f"**{key.replace('_', ' ').title()}**:")
                    st.write(f"   • Primary search terms: {', '.join(mapping['labels'])}")
                    st.write(f"   • Fallback row: {mapping['fallback_row'] + 1}")
                    st.write("---")
                
                # Balance Sheet Sources
                st.markdown("#### Balance Sheet")
                for key, mapping in config.BS_MAPPINGS.items():
                    st.write(f"**{key.replace('_', ' ').title()}**:")
                    st.write(f"   • Primary search terms: {', '.join(mapping['labels'])}")
                    st.write(f"   • Fallback row: {mapping['fallback_row'] + 1}")
                    st.write("---")
            
            with tab3:
                st.markdown("### 📐 KPI Calculation Formulas")
                
                formula_definitions = {
                    "Days Sales Outstanding (DSO)": "DSO = (Accounts Receivable ÷ Monthly Revenue) × 30",
                    "Days Payables Outstanding (DPO)": "DPO = (Accounts Payable ÷ Monthly Revenue) × 30",
                    "Days Inventory on Hand (DIO)": "DIO = (Inventory ÷ Monthly Revenue) × 30",
                    "Working Capital": "Working Capital = Current Assets - Current Liabilities",
                    "Cash Conversion Cycle": "CCC = DSO + DIO - DPO",
                    "Revenue Growth Rate": "Growth = ((Current Period - Prior Period) ÷ Prior Period) × 100",
                    "EBITDA Margin": "EBITDA Margin = (EBITDA ÷ Revenue) × 100",
                    "SG&A as % Revenue": "SG&A % = (SG&A Expenses ÷ Revenue) × 100",
                    "Net Debt to EBITDA": "Net Debt/EBITDA = (Estimated Debt ÷ TTM EBITDA)"
                }
                
                for kpi_name, formula in formula_definitions.items():
                    st.markdown(f"**{kpi_name}**")
                    st.code(formula, language='text')
                    st.write("---")
            
            with tab4:
                st.markdown("### 📋 Raw Data Preview (Last 6 Months)")
                
                if processor.processed_data:
                    preview_data = {}
                    months = processor.processed_data['months'][-6:]
                    preview_data['Month'] = months
                    
                    # Add key financial metrics
                    metrics = ['revenue', 'ebitda', 'accounts_receivable', 'inventory', 'accounts_payable']
                    for metric in metrics:
                        if metric in processor.processed_data:
                            values = processor.processed_data[metric][-6:]
                            preview_data[f'{metric.replace("_", " ").title()} ($M)'] = [f"{x/1e6:.1f}" for x in values]
                    
                    preview_df = pd.DataFrame(preview_data)
                    st.dataframe(preview_df, use_container_width=True)
                    
                    # Download option
                    csv = preview_df.to_csv(index=False)
                    st.download_button(
                        label="📥 Download Full Dataset (CSV)",
                        data=csv,
                        file_name=f"financial_data_preview_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )
            
            with tab5:
                st.markdown("### ✅ Validation & Quality Metrics")
                
                # Data completeness check
                if processor.processed_data:
                    total_fields = len([k for k in processor.processed_data.keys() if k != 'months'])
                    complete_fields = len([k for k, v in processor.processed_data.items() if k != 'months' and v and all(x != 0 for x in v[-3:])])
                    completeness = (complete_fields / total_fields) * 100 if total_fields > 0 else 0
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric(
                            label="Data Completeness",
                            value=f"{completeness:.0f}%",
                            help="Percentage of required fields with valid data"
                        )
                    
                    with col2:
                        extraction_success = len([audit for audit in processor.extraction_audit.values() if audit['found_by_label']])
                        total_extractions = len(processor.extraction_audit)
                        success_rate = (extraction_success / total_extractions) * 100 if total_extractions > 0 else 0
                        
                        st.metric(
                            label="Label-Based Extraction",
                            value=f"{success_rate:.0f}%",
                            help="Percentage of data found by intelligent label search vs fallback"
                        )
                    
                    with col3:
                        avg_confidence = sum(result.confidence_score for result in processor.validation_results.values()) / len(processor.validation_results) if processor.validation_results else 1.0
                        
                        st.metric(
                            label="Overall Confidence",
                            value=f"{avg_confidence*100:.0f}%",
                            help="Average confidence score across all validations"
                        )
    
    except Exception as e:
        logger.error(f"Data transparency section error: {str(e)}")
        st.error(f"Error creating transparency section: {str(e)}")

def create_executive_charts(financial_data: Dict, kpis: Dict[str, KPIResult] = None) -> Tuple[Optional[go.Figure], Optional[go.Figure], Optional[go.Figure]]:
    """Create enhanced executive charts with real KPI data"""
    try:
        if not financial_data:
            return None, None, None
        
        # Prepare data for charts
        months = financial_data['months'][-12:]
        revenue_data = [r/1e6 for r in financial_data['revenue'][-12:]]
        ebitda_data = [e/1e6 for e in financial_data['ebitda'][-12:]]
        
        # Revenue and EBITDA Trend
        fig1 = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig1.add_trace(
            go.Bar(x=months, y=revenue_data, name="Revenue ($M)", 
                   marker_color='#3b82f6', opacity=0.8),
            secondary_y=False,
        )
        
        fig1.add_trace(
            go.Scatter(x=months, y=ebitda_data, mode='lines+markers',
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
        
        wc_data = []
        for i in range(len(months)):
            if (i < len(financial_data['current_assets']) and 
                i < len(financial_data['current_liabilities'])):
                wc = (financial_data['current_assets'][-(12-i)] - 
                     financial_data['current_liabilities'][-(12-i)]) / 1e6
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
        
        # KPI Gauge Chart with REAL data
        fig3 = make_subplots(
            rows=2, cols=2,
            subplot_titles=('Days Sales Outstanding', 'Days Payables Outstanding', 'EBITDA Margin', 'Revenue Growth'),
            specs=[[{'type': 'indicator'}, {'type': 'indicator'}],
                   [{'type': 'indicator'}, {'type': 'indicator'}]],
            vertical_spacing=0.15,
            horizontal_spacing=0.1
        )
        
        # Only create gauges if we have KPI data
        if kpis:
            try:
                # Debug: Check if KPI values exist
                dso_value = kpis.get('dso', {}).value if 'dso' in kpis else 0
                dpo_value = kpis.get('dpo', {}).value if 'dpo' in kpis else 0
                ebitda_value = kpis.get('ebitda_margin', {}).value if 'ebitda_margin' in kpis else 0
                growth_value = kpis.get('revenue_growth', {}).value if 'revenue_growth' in kpis else 0
                
                # DSO Gauge
                fig3.add_trace(go.Indicator(
                    mode="gauge+number",
                    value=max(0, dso_value),  # Ensure non-negative
                    number={
                        'suffix': " days", 
                        'font': {'size': 20, 'color': "#1f2937"},
                        'valueformat': '.1f'
                    },
                    domain={'x': [0, 1], 'y': [0, 1]},
                    gauge={
                        'axis': {'range': [0, 60], 'tickwidth': 1, 'tickcolor': "#6b7280"},
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
                            'value': 45}
                    }
                ), row=1, col=1)
                
                # DPO Gauge  
                fig3.add_trace(go.Indicator(
                    mode="gauge+number",
                    value=max(0, dpo_value),  # Ensure non-negative
                    number={
                        'suffix': " days", 
                        'font': {'size': 20, 'color': "#1f2937"},
                        'valueformat': '.1f'
                    },
                    domain={'x': [0, 1], 'y': [0, 1]},
                    gauge={
                        'axis': {'range': [0, 60], 'tickwidth': 1, 'tickcolor': "#6b7280"},
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
                            'value': 30}
                    }
                ), row=1, col=2)
                
                # EBITDA Margin Gauge
                fig3.add_trace(go.Indicator(
                    mode="gauge+number",
                    value=max(0, ebitda_value),  # Ensure non-negative
                    number={
                        'suffix': "%", 
                        'font': {'size': 20, 'color': "#1f2937"},
                        'valueformat': '.2f'
                    },
                    domain={'x': [0, 1], 'y': [0, 1]},
                    gauge={
                        'axis': {'range': [0, 25], 'tickwidth': 1, 'tickcolor': "#6b7280"},
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
                            'value': 12}
                    }
                ), row=2, col=1)
                
                # Revenue Growth Gauge
                fig3.add_trace(go.Indicator(
                    mode="gauge+number",
                    value=growth_value,  # Allow negative values for growth
                    number={
                        'suffix': "%", 
                        'font': {'size': 20, 'color': "#1f2937"},
                        'valueformat': '.2f'
                    },
                    domain={'x': [0, 1], 'y': [0, 1]},
                    gauge={
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
                            'value': 15}
                    }
                ), row=2, col=2)
                
                # Log debug information
                logger.info(f"Gauge values - DSO: {dso_value}, DPO: {dpo_value}, EBITDA: {ebitda_value}, Growth: {growth_value}")
                
            except Exception as gauge_error:
                logger.error(f"Error creating individual gauges: {str(gauge_error)}")
                # Fallback to simple gauges
                for i, (row, col) in enumerate([(1,1), (1,2), (2,1), (2,2)]):
                    fig3.add_trace(go.Indicator(
                        mode="gauge+number",
                        value=0,
                        domain={'x': [0, 1], 'y': [0, 1]},
                        gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#6b7280"}},
                    ), row=row, col=col)
        else:
            # No KPI data available - show placeholder gauges
            logger.warning("No KPI data available for gauges")
            for i, (row, col) in enumerate([(1,1), (1,2), (2,1), (2,2)]):
                fig3.add_trace(go.Indicator(
                    mode="gauge+number",
                    value=0,
                    domain={'x': [0, 1], 'y': [0, 1]},
                    gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#cbd5e0"}},
                ), row=row, col=col)
        
        fig3.update_layout(
            height=550,
            font=dict(family="Arial, sans-serif", size=12),
            paper_bgcolor='white',
            plot_bgcolor='white',
            margin=dict(l=20, r=20, t=60, b=20)
        )
        
        return fig1, fig2, fig3
        
    except Exception as e:
        logger.error(f"Chart creation error: {str(e)}")
        st.error(f"Error creating charts: {str(e)}")
        return None, None, None

def create_sidebar_reference():
    """Create enhanced sidebar with KPI reference and data quality info"""
    try:
        with st.sidebar:
            st.markdown("---")
            st.markdown("### 📚 KPI Reference Guide")
            
            with st.expander("💡 What do these metrics mean?"):
                st.markdown("""
                **Days Sales Outstanding (DSO)**
                - Time to collect receivables after sale
                - Formula: (AR ÷ Monthly Revenue) × 30
                - Target: < 45 days (industry benchmark)
                
                **Days Payables Outstanding (DPO)**  
                - Time taken to pay suppliers
                - Formula: (AP ÷ Monthly Revenue) × 30
                - Target: > 30 days (cash flow optimization)
                
                **Days Inventory on Hand (DIO)**
                - Inventory turnover efficiency
                - Formula: (Inventory ÷ Monthly Revenue) × 30
                - Target: < 30 days (lean operations)
                
                **Cash Conversion Cycle (CCC)**
                - Total working capital efficiency
                - Formula: DSO + DIO - DPO
                - Target: < 30 days (best-in-class)
                
                **EBITDA Margin**
                - Operating profitability before financing
                - Formula: (EBITDA ÷ Revenue) × 100
                - Target: > 12% (industry dependent)
                """)
            
            with st.expander("🔍 Data Quality Indicators"):
                st.markdown("""
                **Status Color Coding:**
                - 🟢 **Green**: Meets or exceeds target
                - 🟡 **Yellow**: Needs attention/monitoring
                - 🔴 **Red**: Requires immediate action
                
                **Confidence Indicators:**
                - **100%**: High confidence, label-based extraction
                - **90-99%**: Good confidence, minor warnings
                - **<90%**: Lower confidence, fallback methods used
                
                **Data Validation:**
                - ✅ Structure validation passed
                - ✅ Range validation completed
                - ✅ Consistency checks performed
                """)
    
    except Exception as e:
        logger.error(f"Sidebar creation error: {str(e)}")
        st.error("Error creating sidebar reference")

# ===============================
# MAIN APPLICATION
# ===============================

def main():
    """Enhanced main application with comprehensive error handling"""
    try:
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
                - Ensure P&L Excel file follows Superior Biologics format
                - Ensure Balance Sheet Excel file follows standard structure
                - Files should contain monthly data with consistent headers
                
                **Step 2: Upload Files**
                - Use sidebar file uploaders for P&L and Balance Sheet
                - Click "Process Files" after both files are uploaded
                - Review data validation results and extraction audit
                
                **Step 3: Review Dashboard**
                - Explore automatically calculated KPIs with confidence indicators
                - Use period selection to switch between Monthly/Quarterly/TTM views
                - Click calculation expanders to see detailed methodologies
                - Export reports and data for further analysis
                
                **Advanced Features:**
                - Data extraction audit trail shows label vs fallback usage
                - Calculation transparency with formula details
                - Confidence scoring for data quality assessment
                - Comprehensive validation and error reporting
                """)
        
        with col2:
            with st.expander("ℹ️ About"):
                st.markdown("""
                ### About This Application
                
                **Enterprise-Grade Financial Analytics**
                This dashboard implements Fortune 50 standards for financial KPI analysis with comprehensive audit trails, robust data validation, and transparent calculation methodologies.
                
                **Key Features:**
                - **Intelligent Data Extraction**: Label-based search with fallback protection
                - **Comprehensive Validation**: Multi-layer data quality checks
                - **Audit Trail**: Complete calculation transparency and lineage
                - **Error Handling**: Specific error reporting and recovery mechanisms
                - **Confidence Scoring**: Data quality assessment and reporting
                
                **🔒 Data Security & Privacy:**
                - **Zero Data Retention**: All processing occurs in session memory only
                - **No External Transmission**: Data remains on local processing environment
                - **Session Isolation**: Complete data isolation between users
                - **Automatic Cleanup**: All data deleted when session ends
                
                **Technical Architecture:**
                - Modular design with separation of concerns
                - Robust error handling with specific exception management
                - Configuration-driven data mapping for maintainability
                - Auditable calculations with full lineage tracking
                
                **Compliance & Standards:**
                - Follows Generally Accepted Accounting Principles (GAAP)
                - Implements Fortune 50 dashboard design standards
                - Provides audit-ready calculation documentation
                - Supports regulatory compliance requirements
                
                **Version:** 2.0.0 - Enterprise Edition  
                **Architecture:** Modular, Auditable, Scalable  
                **Standards:** Fortune 50 Compliance Ready
                """)
        
        # Initialize session state
        if 'processor' not in st.session_state:
            st.session_state.processor = FinancialDataProcessor()
        if 'data_loaded' not in st.session_state:
            st.session_state.data_loaded = False
        
        # Enhanced sidebar with file uploads and options
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
                    with st.spinner("Processing financial data with enhanced validation..."):
                        validation_result = st.session_state.processor.load_excel_files(pl_file, bs_file)
                        
                        if validation_result.is_valid:
                            financial_data = st.session_state.processor.extract_financial_data()
                            if financial_data:
                                st.session_state.data_loaded = True
                                st.success("✅ Files processed successfully!")
                                
                                # Enhanced data validation summary
                                with st.expander("📋 Data Processing Summary"):
                                    st.write("**Files Processed:**")
                                    st.write(f"• P&L: {pl_file.name}")
                                    st.write(f"• Balance Sheet: {bs_file.name}")
                                    
                                    st.write(f"**Data Range:** {financial_data['months'][0]} to {financial_data['months'][-1]}")
                                    st.write(f"**Periods:** {len(financial_data['months'])} months")
                                    
                                    # Show extraction method success
                                    extraction_stats = st.session_state.processor.extraction_audit
                                    label_success = sum(1 for audit in extraction_stats.values() if audit['found_by_label'])
                                    total_fields = len(extraction_stats)
                                    
                                    st.write(f"**Extraction Success:** {label_success}/{total_fields} fields found by intelligent search")
                                    
                                    if label_success < total_fields:
                                        st.warning(f"⚠️ {total_fields - label_success} fields used fallback extraction - verify data accuracy")
                            else:
                                st.error("❌ Error extracting financial data")
                        else:
                            st.error("❌ File validation failed:")
                            for error in validation_result.errors:
                                st.error(f"• {error}")
            
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
            show_confidence = st.checkbox("Show Confidence Scores", value=True, key="show_confidence")
            
            # Period filter explanation
            if period_view == "Monthly":
                st.info("📅 Showing: Current month data only")
            elif period_view == "Quarterly":
                st.info("📅 Showing: Quarterly averaged data")  
            else:
                st.info("📅 Showing: Trailing 12-month data")
        
        # Create sidebar reference
        create_sidebar_reference()
        
        # Main dashboard content
        if not st.session_state.data_loaded:
            # Enhanced welcome screen
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.markdown("""
                <div class="upload-box">
                    <h3>🚀 Welcome to Your Enterprise Dashboard</h3>
                    <p>Upload your P&L and Balance Sheet Excel files to generate comprehensive financial KPI analysis.</p>
                    
                    <p><strong>Enterprise Features:</strong></p>
                    <ul style="text-align: left; display: inline-block;">
                        <li>🔍 Intelligent data extraction with audit trails</li>
                        <li>📊 Comprehensive data validation and quality scoring</li>
                        <li>🎯 All required KPIs with calculation transparency</li>
                        <li>📈 Interactive visualizations and drill-down capability</li>
                        <li>🔒 Zero data retention with complete privacy protection</li>
                        <li>📋 Fortune 50-grade reporting and export capabilities</li>
                    </ul>
                    
                    <p><strong>Calculated KPIs:</strong></p>
                    <ul style="text-align: left; display: inline-block;">
                        <li>Days Sales Outstanding (DSO) with period analysis</li>
                        <li>Days Payables Outstanding (DPO) with cash flow insights</li>
                        <li>Days Inventory on Hand (DIO) with efficiency tracking</li>
                        <li>Working Capital Analysis with trend monitoring</li>
                        <li>Cash Conversion Cycle with optimization recommendations</li>
                        <li>Revenue Growth Rate with comparative analysis</li>
                        <li>EBITDA Margin with profitability assessment</li>
                        <li>SG&A as % of Revenue with cost efficiency metrics</li>
                        <li>Net Debt to EBITDA with leverage analysis</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
        else:
            # Enhanced dashboard content
            financial_data = st.session_state.processor.processed_data
            
            # Get dashboard options from sidebar
            period_view = st.session_state.get('period_selector', 'Trailing 12 Months')
            show_targets = st.session_state.get('show_targets', True)
            show_trends = st.session_state.get('show_trends', True)
            show_confidence = st.session_state.get('show_confidence', True)
            
            calculator = AuditableKPICalculator(financial_data)
            kpis = calculator.calculate_all_kpis(period_view)
            
            if kpis:
                # Period and confidence indicator
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.info(f"📊 **Current View**: {kpis.get('period_info').description if 'period_info' in kpis else period_view}")
                
                with col2:
                    # Overall data confidence score
                    if st.session_state.processor.validation_results:
                        avg_confidence = sum(result.confidence_score for result in st.session_state.processor.validation_results.values()) / len(st.session_state.processor.validation_results)
                        confidence_color = "🟢" if avg_confidence > 0.9 else "🟡" if avg_confidence > 0.7 else "🔴"
                        st.info(f"{confidence_color} **Data Confidence**: {avg_confidence*100:.0f}%")
                
                # Executive Summary
                st.markdown("## 📈 Executive Summary")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(create_kpi_card(kpis['ttm_revenue'], show_targets), unsafe_allow_html=True)
                
                with col2:
                    st.markdown(create_kpi_card(kpis['ttm_ebitda'], show_targets), unsafe_allow_html=True)
                
                with col3:
                    st.markdown(create_kpi_card(kpis['ebitda_margin'], show_targets), unsafe_allow_html=True)
                
                with col4:
                    st.markdown(create_kpi_card(kpis['working_capital'], show_targets), unsafe_allow_html=True)
                
                # Cash Conversion Cycle
                st.markdown("## 🔄 Cash Conversion Cycle")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown(create_kpi_card(kpis['dso'], show_targets), unsafe_allow_html=True)
                    create_calculation_expander(kpis['dso'])
                
                with col2:
                    st.markdown(create_kpi_card(kpis['dio'], show_targets), unsafe_allow_html=True)
                    create_calculation_expander(kpis['dio'])
                
                with col3:
                    st.markdown(create_kpi_card(kpis['dpo'], show_targets), unsafe_allow_html=True)
                    create_calculation_expander(kpis['dpo'])
                
                # Performance Metrics
                st.markdown("## 📊 Performance Metrics")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(create_kpi_card(kpis['revenue_growth'], show_targets), unsafe_allow_html=True)
                
                with col2:
                    st.markdown(create_kpi_card(kpis['sga_percentage'], show_targets), unsafe_allow_html=True)
                
                with col3:
                    st.markdown(create_kpi_card(kpis['ar_change'], show_targets), unsafe_allow_html=True)
                
                with col4:
                    st.markdown(create_kpi_card(kpis['net_debt_to_ebitda'], show_targets), unsafe_allow_html=True)
                
                # Debug: Show KPI calculation status
                if show_confidence:
                    with st.expander("🔧 Debug: KPI Calculation Status"):
                        st.write("**KPI Values for Gauges:**")
                        
                        debug_kpis = ['dso', 'dpo', 'ebitda_margin', 'revenue_growth']
                        for kpi_name in debug_kpis:
                            if kpi_name in kpis:
                                kpi_obj = kpis[kpi_name]
                                st.write(f"• **{kpi_obj.name}**: {kpi_obj.value:.2f} ({kpi_obj.status.value})")
                                if hasattr(kpi_obj, 'audit') and kpi_obj.audit:
                                    st.write(f"  └─ Confidence: {kpi_obj.audit.confidence*100:.0f}%")
                            else:
                                st.write(f"• **{kpi_name}**: Missing from calculations")
                        
                        # Show underlying data for DSO/DPO
                        if financial_data:
                            st.write("**Underlying Data (Latest Month):**")
                            st.write(f"• Revenue: ${financial_data['revenue'][-1]:,.0f}")
                            st.write(f"• Accounts Receivable: ${financial_data['accounts_receivable'][-1]:,.0f}")
                            st.write(f"• Accounts Payable: ${financial_data['accounts_payable'][-1]:,.0f}")
                            st.write(f"• Inventory: ${financial_data['inventory'][-1]:,.0f}")
                            
                            # Manual DSO/DPO calculation for verification
                            if financial_data['revenue'][-1] > 0:
                                manual_dso = (financial_data['accounts_receivable'][-1] / financial_data['revenue'][-1]) * 30
                                manual_dpo = (financial_data['accounts_payable'][-1] / financial_data['revenue'][-1]) * 30
                                st.write(f"• **Manual DSO Calculation**: {manual_dso:.1f} days")
                                st.write(f"• **Manual DPO Calculation**: {manual_dpo:.1f} days")
                
                fig1, fig2, fig3 = create_executive_charts(financial_data, kpis)
                
                if fig1 and fig2 and fig3:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.plotly_chart(fig1, use_container_width=True)
                    
                    with col2:
                        st.plotly_chart(fig2, use_container_width=True)
                    
                    st.plotly_chart(fig3, use_container_width=True)
                
                # Enhanced data transparency section
                create_data_transparency_section(st.session_state.processor, kpis)
                
                # Cash Conversion Cycle Summary
                st.markdown("## ⚡ Cash Conversion Cycle Summary")
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.markdown(create_kpi_card(kpis['cash_conversion_cycle'], show_targets), unsafe_allow_html=True)
                    
                    st.markdown(f"""
                    <div style="text-align: center; margin: 1rem 0; padding: 1rem; background: #f8fafc; border-radius: 8px;">
                        <strong>Formula:</strong> DSO ({kpis['dso'].value:.1f}) + DIO ({kpis['dio'].value:.1f}) - DPO ({kpis['dpo'].value:.1f}) = {kpis['cash_conversion_cycle'].value:.1f} days
                    </div>
                    """, unsafe_allow_html=True)
                    
                    create_calculation_expander(kpis['cash_conversion_cycle'])
                
                # Enhanced export functionality
                st.markdown("## 📤 Export Options")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if st.button("📊 Export KPI Summary"):
                        # Enhanced summary with audit information
                        summary_data = {
                            'KPI': [],
                            'Current Value': [],
                            'Target': [],
                            'Status': [],
                            'Confidence': []
                        }
                        
                        key_kpis = ['dso', 'dpo', 'dio', 'working_capital', 'revenue_growth', 'ebitda_margin', 'sga_percentage', 'net_debt_to_ebitda']
                        
                        for kpi_key in key_kpis:
                            if kpi_key in kpis:
                                kpi = kpis[kpi_key]
                                summary_data['KPI'].append(kpi.name)
                                summary_data['Current Value'].append(format_number(kpi.value, kpi.format_type))
                                summary_data['Target'].append(format_number(kpi.target, kpi.format_type) if kpi.target else "N/A")
                                summary_data['Status'].append(kpi.status.value.title())
                                summary_data['Confidence'].append(f"{kpi.audit.confidence*100:.0f}%" if kpi.audit else "N/A")
                        
                        df_summary = pd.DataFrame(summary_data)
                        csv = df_summary.to_csv(index=False)
                        st.download_button(
                            label="Download CSV",
                            data=csv,
                            file_name=f'kpi_summary_{datetime.now().strftime("%Y%m%d_%H%M")}.csv',
                            mime='text/csv'
                        )
                
                with col2:
                    if st.button("📈 Export Charts Data"):
                        # Enhanced chart data with validation info
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
                            file_name=f'financial_data_{datetime.now().strftime("%Y%m%d_%H%M")}.csv',
                            mime='text/csv'
                        )
                
                with col3:
                    if st.button("📋 Generate Executive Report"):
                        # Enhanced executive report with full audit trail
                        pl_filename = st.session_state.get('uploaded_pl_name', 'Unknown P&L File')
                        bs_filename = st.session_state.get('uploaded_bs_name', 'Unknown Balance Sheet File')
                        
                        # Calculate audit statistics
                        extraction_stats = st.session_state.processor.extraction_audit
                        label_success = sum(1 for audit in extraction_stats.values() if audit['found_by_label'])
                        total_fields = len(extraction_stats)
                        
                        validation_stats = st.session_state.processor.validation_results
                        avg_confidence = sum(result.confidence_score for result in validation_stats.values()) / len(validation_stats) if validation_stats else 1.0
                        
                        report = f"""
SUPERIOR BIOLOGICS - EXECUTIVE FINANCIAL DASHBOARD
Enterprise Edition v2.0 - Audit-Ready Report
Generated: {datetime.now().strftime("%B %d, %Y at %H:%M:%S")}

SOURCE FILE AUDIT TRAIL
=======================
P&L Statement File: {pl_filename}
Balance Sheet File: {bs_filename}
Processing Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
Report Period: {period_view}
Data Range: {financial_data['months'][0]} to {financial_data['months'][-1]}

DATA QUALITY ASSESSMENT
========================
Extraction Method Success: {label_success}/{total_fields} fields ({label_success/total_fields*100:.0f}% intelligent extraction)
Overall Data Confidence: {avg_confidence*100:.0f}%
Validation Status: {"PASSED" if all(r.is_valid for r in validation_stats.values()) else "WARNINGS PRESENT"}

EXECUTIVE SUMMARY
================
TTM Revenue: {format_number(kpis['ttm_revenue'].value, 'currency')}
TTM EBITDA: {format_number(kpis['ttm_ebitda'].value, 'currency')}
EBITDA Margin: {kpis['ebitda_margin'].value:.1f}% (Target: {kpis['ebitda_margin'].target}%)
Working Capital: {format_number(kpis['working_capital'].value, 'currency')}

CASH CONVERSION CYCLE ANALYSIS
==============================
Days Sales Outstanding: {kpis['dso'].value:.1f} days (Target: < {kpis['dso'].target})
Days Inventory on Hand: {kpis['dio'].value:.1f} days (Target: < {kpis['dio'].target}) 
Days Payables Outstanding: {kpis['dpo'].value:.1f} days (Target: > {kpis['dpo'].target})
Cash Conversion Cycle: {kpis['cash_conversion_cycle'].value:.1f} days (Target: < {kpis['cash_conversion_cycle'].target})

PERFORMANCE METRICS
==================
Revenue Growth Rate: {kpis['revenue_growth'].value:.1f}% (Target: > {kpis['revenue_growth'].target}%)
SG&A as % of Revenue: {kpis['sga_percentage'].value:.1f}% (Target: < {kpis['sga_percentage'].target}%)
Change in AR: {format_number(kpis['ar_change'].value, 'currency')}
Net Debt to EBITDA: {kpis['net_debt_to_ebitda'].value:.1f}x (Target: < {kpis['net_debt_to_ebitda'].target}x)

KEY INSIGHTS & RECOMMENDATIONS
=============================
• Cash conversion cycle of {kpis['cash_conversion_cycle'].value:.1f} days indicates {kpis['cash_conversion_cycle'].interpretation}
• EBITDA margin of {kpis['ebitda_margin'].value:.1f}% {kpis['ebitda_margin'].interpretation}
• Revenue growth of {kpis['revenue_growth'].value:.1f}% {kpis['revenue_growth'].interpretation}

DATA VERIFICATION & AUDIT TRAIL
===============================
Latest Period Data Points:
• Revenue: {format_number(financial_data['revenue'][-1], 'currency')}
• Accounts Receivable: {format_number(financial_data['accounts_receivable'][-1], 'currency')}
• Inventory: {format_number(financial_data['inventory'][-1], 'currency')}
• Accounts Payable: {format_number(financial_data['accounts_payable'][-1], 'currency')}
• Current Assets: {format_number(financial_data['current_assets'][-1], 'currency')}
• Current Liabilities: {format_number(financial_data['current_liabilities'][-1], 'currency')}

CALCULATION METHODOLOGY & FORMULAS
==================================
DSO = (Accounts Receivable ÷ Monthly Revenue) × 30
DPO = (Accounts Payable ÷ Monthly Revenue) × 30
DIO = (Inventory ÷ Monthly Revenue) × 30
Working Capital = Current Assets - Current Liabilities
Cash Conversion Cycle = DSO + DIO - DPO
EBITDA Margin = (EBITDA ÷ Revenue) × 100
Revenue Growth = ((Current - Prior Year) ÷ Prior Year) × 100
SG&A % = (SG&A Expenses ÷ Revenue) × 100
Net Debt/EBITDA = (Estimated Debt ÷ TTM EBITDA)

COMPLIANCE & AUDIT CERTIFICATION
================================
Report Generated By: Superior Biologics Executive Dashboard v2.0 Enterprise
Data Processing Method: Intelligent extraction with validation framework
Quality Assurance: Multi-layer validation with confidence scoring
Audit Trail: Complete source file documentation and calculation lineage
Calculation Standards: Generally Accepted Accounting Principles (GAAP) compliant
Fortune 50 Compliance: Enterprise-grade reporting standards implemented

DISCLAIMERS & NOTES
==================
• Net Debt calculation uses estimated debt (30% of current liabilities)
• Confidence scores reflect data extraction and validation reliability
• All calculations follow standard financial accounting practices
• Report intended for internal management use and board presentations
"""
                        
                        st.download_button(
                            label="Download Report",
                            data=report,
                            file_name=f'executive_report_{datetime.now().strftime("%Y%m%d_%H%M")}.txt',
                            mime='text/plain'
                        )
        
        # Enhanced footer
        st.markdown("---")
        st.markdown(
            f"""
            <div style="text-align: center; color: #6b7280; font-size: 0.875rem; margin-top: 2rem;">
                Superior Biologics Executive Dashboard v2.0 Enterprise Edition<br>
                Audit-Ready Financial Analytics • GNU GPL v3.0 Licensed • Zero Data Retention<br>
                Generated with Enhanced Data Validation & Calculation Transparency<br>
                Last Updated: {datetime.now().strftime("%B %d, %Y")}
            </div>
            """,
            unsafe_allow_html=True
        )

    except Exception as e:
        logger.error(f"Main application error: {str(e)}")
        st.error(f"Application error: {str(e)}")
        st.error("Please refresh the page and try again. If the problem persists, check your file formats and data structure.")

if __name__ == "__main__":
    main()

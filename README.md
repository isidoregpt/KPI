# Superior Biologics Executive Dashboard

A professional-grade financial KPI dashboard designed for Fortune 50 executive review. This Streamlit application automatically processes P&L and Balance Sheet data to generate comprehensive financial performance metrics and interactive visualizations.

![Dashboard Preview](https://img.shields.io/badge/Streamlit-Dashboard-FF6B6B?style=for-the-badge&logo=streamlit)
![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python)
![Plotly](https://img.shields.io/badge/Plotly-Interactive-3F4F75?style=for-the-badge&logo=plotly)

## 🎯 Features

### Executive-Level KPIs
- **Days Sales Outstanding (DSO)** - Receivables collection efficiency
- **Days Payables Outstanding (DPO)** - Supplier payment optimization  
- **Days Inventory on Hand (DIO)** - Inventory turnover analysis
- **Working Capital Analysis** - Liquidity and operational efficiency
- **Revenue Growth Rate** - Year-over-year performance tracking
- **EBITDA Margin** - Profitability analysis
- **SG&A as % Revenue** - Operational cost management
- **Cash Conversion Cycle** - Complete working capital efficiency
- **Net Debt to EBITDA** - Leverage and financial risk assessment

### Professional Visualizations
- Interactive revenue and EBITDA trend charts
- Working capital waterfall analysis
- KPI gauge dashboards with performance targets
- Color-coded status indicators (green/yellow/red)
- Executive-grade styling and typography

### Data Security & Privacy
- 🔒 **No data retention** - All processing in memory only
- 🔒 **No cloud storage** - Files deleted when session ends
- 🔒 **No external transmission** - Data stays local
- 🔒 **Session-based processing** - Complete privacy protection

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- Excel files in Superior Biologics format (P&L and Balance Sheet)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/superior-biologics-dashboard.git
   cd superior-biologics-dashboard
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   streamlit run financial_dashboard.py
   ```

4. **Access the dashboard**
   - Open your browser to `http://localhost:8501`
   - Upload your P&L and Balance Sheet Excel files
   - View automatically generated KPIs and visualizations

## 📊 Supported File Formats

### P&L Statement Requirements
- **File format**: Excel (.xlsx, .xls)
- **Sheet name**: `CustomIncomeStatementbyAc`
- **Structure**: Months in row 7, columns B-P
- **Required data**: Revenue, EBITDA, Operating Expenses, SG&A

### Balance Sheet Requirements  
- **File format**: Excel (.xlsx, .xls)
- **Sheet name**: `CustomBalanceSheetMonth`
- **Structure**: Months in row 7, columns D-R
- **Required data**: Accounts Receivable, Inventory, Accounts Payable, Current Assets/Liabilities

## 🎨 Dashboard Sections

### Executive Summary
- TTM Revenue and EBITDA performance
- EBITDA margin with target benchmarks
- Working capital position analysis

### Cash Conversion Cycle
- Detailed DSO, DPO, and DIO analysis
- Performance vs. industry targets
- Trend indicators and status alerts

### Performance Metrics
- Revenue growth tracking
- Operational efficiency ratios
- Financial leverage indicators

### Interactive Charts
- Revenue and EBITDA trend analysis
- Working capital component visualization
- KPI gauge dashboard with targets

## 📁 File Structure

```
superior-biologics-dashboard/
├── financial_dashboard.py      # Main Streamlit application
├── requirements.txt           # Python dependencies
├── README.md                 # This file
├── assets/                   # Optional: Screenshots and documentation
└── examples/                # Optional: Sample Excel files (anonymized)
```

## 🔧 Configuration

### Customizing KPI Targets
Edit the target values in the `KPICalculator.calculate_all_kpis()` method:

```python
'dso': {'value': dso, 'format': 'days', 'target': 45, 'status': 'good' if dso < 45 else 'warning'},
'dpo': {'value': dpo, 'format': 'days', 'target': 30, 'status': 'good' if dpo > 30 else 'warning'},
# ... other KPIs
```

### Modifying Chart Styles
Update colors and styling in the `create_executive_charts()` function to match corporate branding.

## 📈 KPI Calculations

### Working Capital Metrics
- **DSO** = (Accounts Receivable ÷ Monthly Revenue) × 30
- **DPO** = (Accounts Payable ÷ Monthly Revenue) × 30  
- **DIO** = (Inventory ÷ Monthly Revenue) × 30
- **Cash Conversion Cycle** = DSO + DIO - DPO

### Performance Ratios
- **EBITDA Margin** = (EBITDA ÷ Revenue) × 100
- **SG&A %** = (SG&A Expenses ÷ Revenue) × 100
- **Revenue Growth** = ((Current Revenue - Prior Year Revenue) ÷ Prior Year Revenue) × 100

## 🛠️ Troubleshooting

### Common Issues

**File Processing Errors**
- Verify Excel files match the expected format
- Check that month headers are in row 7
- Ensure all required financial line items are present

**Missing KPI Data**
- Confirm P&L includes Total Revenue, EBITDA, and Operating Expenses
- Verify Balance Sheet contains AR, Inventory, AP, and Current Assets/Liabilities

**Chart Display Issues**
- Update your browser to the latest version
- Clear browser cache and reload the application

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the **GNU General Public License v3.0** - see the [LICENSE](LICENSE) file for details.

### License Summary
- ✅ **Use**: You can use this software for any purpose, including commercial use
- ✅ **Modify**: You can modify the source code
- ✅ **Distribute**: You can distribute copies of the software
- ⚠️ **Share Alike**: Any derivative works must also be licensed under GPL v3.0
- ⚠️ **Source Required**: You must provide source code when distributing
- ❌ **No Proprietary Forks**: Cannot create closed-source commercial versions

**Important for Commercial Users**: While you can use this software internally within your organization, any distribution or modification requires compliance with GPL v3.0 terms, including sharing your source code modifications.

## 🏢 Enterprise Usage & Licensing

This dashboard is designed for enterprise financial analysis and executive reporting. It follows Fortune 50 design standards and provides professional-grade visualizations expected in corporate environments.

### Commercial Use Guidelines
- **Internal Use**: Organizations may use this software internally without restriction
- **Modifications**: Any customizations must be made available under GPL v3.0 if distributed
- **Distribution**: Sharing the software (modified or unmodified) requires GPL v3.0 compliance
- **Consulting/Services**: You may provide consulting services using this software
- **SaaS/Hosting**: If offering as a service, source code must be made available to users

### Prohibited Uses Under GPL v3.0
- ❌ Creating proprietary/closed-source commercial versions
- ❌ Incorporating into proprietary software without GPL compliance
- ❌ Selling licenses for closed-source usage
- ❌ Removing copyright notices or license information

### Security Considerations
- All data processing occurs locally
- No data is transmitted to external servers
- Files are automatically deleted after session ends
- Suitable for sensitive financial information
- GPL v3.0 ensures transparency and security through open source requirements

## 📞 Support

For technical support or feature requests:
- Create an issue in this repository
- Contact your IT or Finance team for deployment assistance
- Review the Instructions dropdown in the application for user guidance

## 🙏 Acknowledgments

- Built with [Streamlit](https://streamlit.io/) for the web framework
- Powered by [Plotly](https://plotly.com/) for interactive visualizations
- Data processing with [pandas](https://pandas.pydata.org/) and [openpyxl](https://openpyxl.readthedocs.io/)
- Designed for Superior Biologics executive reporting requirements

---

**Version**: 1.0.0  
**Last Updated**: December 2024  
**Designed for**: Fortune 50 Executive Review

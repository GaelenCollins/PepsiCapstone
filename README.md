# E80 Rejected Units Analyzer

A comprehensive Python application for analyzing rejected units data from the E80 Automated Warehouse system at Pepsi's Stone Mountain facility. This tool helps identify root causes of pallet rejections and provides actionable insights to reduce the current 3% rejection rate to below the corporate standard of <1%.

## Features

### 📊 Dashboard
- **Key Metrics**: Total rejections, quantities, estimated rejection rates
- **Visual Summary**: Colorful metric cards with real-time statistics
- **Analysis Overview**: Comprehensive summary of rejection patterns

### 🔍 Root Cause Analysis
- **Advanced Filtering**: Filter by date range, production line, and rejection reason
- **Pattern Detection**: Automatic root cause classification
- **Detailed Breakdown**: Line-by-line analysis of rejection data

### 📈 Trend Analysis
- **Interactive Charts**: Multiple visualization types (pie, bar, line, horizontal bar)
- **Temporal Analysis**: Monthly trends and seasonal patterns
- **Performance Tracking**: Day-of-week and time-based analysis

### 📋 Report Generation
- **Summary Reports**: Executive-level overview with key findings
- **Detailed Analysis**: Comprehensive breakdown with statistics
- **Trend Reports**: Forecasting and trend direction analysis

## Installation

1. **Clone or download** the application files
2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Launch the application**:
   ```bash
   python rejected_units_analyzer.py
   ```

2. **Load your data**:
   - Click "Choose File" in the Upload tab
   - Select your E80 rejected units Excel file (.xlsx)
   - Click "Analyze Data" to process

3. **Explore the analysis**:
   - **Dashboard**: View key metrics and summary
   - **Root Cause Analysis**: Filter and examine detailed data
   - **Trends**: Analyze patterns and trends
   - **Reports**: Generate comprehensive reports

## Data Format

The application expects Excel files with the following columns:
- `Date`: Rejection date
- `Line`: Production line identifier
- `Product`: Product name/code
- `Rejection Reason`: Reason for rejection
- `Quantity`: Number of units rejected
- `Shift`: Production shift (optional)
- `Operator`: Operator identifier (optional)

## Key Benefits

### For Operations Teams
- **Quick Identification** of top rejection reasons
- **Line-specific Analysis** to target improvement efforts
- **Trend Monitoring** to track progress over time

### For Management
- **Cost Impact Analysis** with estimated financial losses
- **Executive Reports** for stakeholder communication
- **Performance Metrics** to measure improvement initiatives

### For Engineering
- **Root Cause Classification** for systematic problem-solving
- **Pattern Recognition** to identify systemic issues
- **Data-driven Insights** for process optimization

## Technical Features

- **Modern UI**: Dark/light theme support with professional styling
- **Responsive Design**: Adapts to different screen sizes
- **Data Validation**: Automatic format checking and error handling
- **Export Capabilities**: Generate reports in multiple formats
- **Real-time Analysis**: Instant processing and visualization updates

## Cost Impact

Based on industry standards, each rejected pallet costs approximately $50-100 in rework and disposal. With the current 3% rejection rate, this tool helps identify opportunities to reduce the estimated annual loss of $80,000–$90,000.

## Future Enhancements

- Machine learning integration for predictive analysis
- Real-time data connectivity
- Advanced statistical modeling
- Custom dashboard configuration
- Automated alert systems

## Support

For technical support or feature requests, please contact the development team.

---

**Version**: 1.0  
**Last Updated**: 2025-01-10  
**Compatibility**: Python 3.7+, Windows/macOS/Linux

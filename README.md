# Transportation Analysis Script

A Python script for analyzing transportation data and generating reports comparing needed vs scheduled loads for Amazon facilities.

## Features

- **Data Analysis**: Compares cart consumption data with scheduled truck loads
- **Multi-site Support**: Handles FC (Fulfillment Centers), SC (Sort Centers), and DS (Delivery Stations)
- **Date Range Processing**: Analyzes current and future dates from Excel serial dates
- **HTML Report Generation**: Creates professional-looking HTML reports with styling
- **Email Integration**: Sends reports via Outlook (Windows) or saves for manual sending
- **Logging**: Comprehensive logging for debugging and monitoring
- **Error Handling**: Robust error handling with graceful fallbacks

## Requirements

- Python 3.7+
- pandas
- openpyxl
- pywin32 (Windows only, for Outlook integration)

## Installation

1. Clone or download the script files
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

```python
python transportation_analysis_improved.py
```

### Configuration

Edit the file paths in the `main()` function:

```python
# File paths (local)
carts_file = r'\\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias & Underbias_Trigger\Carts_Consumption.xlsx'
fmc_file = r'\\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias & Underbias_Trigger\fmc-data-source(5).xlsx'

# Current serial date (August 28, 2025)
current_serial = 45898
```

### Programmatic Usage

```python
from transportation_analysis_improved import TransportationAnalyzer

# Create analyzer
analyzer = TransportationAnalyzer(carts_file, fmc_file, current_serial)

# Run analysis
success = analyzer.run_analysis()
```

## Input Data Format

### Carts Consumption File (`Carts_Consumption.xlsx`)
- Sheet name: `Final_Sheet`
- Required columns:
  - `Region`: Site region
  - `Type`: Site type (FC, SC, DS)
  - `Site`: Site identifier
  - Columns D:J: Consumption data (dates 45894-45900)
  - Columns K:Q: Needed TFR 80 trucks
  - Columns R:X: Needed TFR 93 trucks

### FMC Data File (`fmc-data-source(5).xlsx`)
- Sheet name: `fmc-data-source(5)`
- Required columns:
  - `Stop`: Site identifier
  - `Action Type`: DROPOFF or PICKUP
  - `Planned Dock Arrival`: Excel serial date

## Output

### HTML Report
- Saved as `transportation_report.html`
- Includes summary statistics
- Professional styling with hover effects
- Color-coded differences

### Log File
- Saved as `transportation_analysis.log`
- Detailed execution logs
- Error tracking

### Email
- Sent via Outlook (Windows)
- HTML formatted
- Includes summary and detailed table

## Analysis Logic

1. **Data Loading**: Reads Excel files and validates data
2. **Date Processing**: Converts Excel serial dates to readable format
3. **Site Filtering**: Processes only FC, SC, and DS sites
4. **Truck Counting**:
   - FC/SC: Counts DROPOFF actions
   - DS: Counts PICKUP actions
5. **Comparison**: Calculates difference between needed and scheduled trucks
6. **Reporting**: Generates HTML report with analysis results

## Error Handling

- File access errors
- Data format issues
- Missing dependencies
- Email sending failures
- Network connectivity issues

## Cross-Platform Support

- **Windows**: Full functionality including Outlook integration
- **Linux/macOS**: All features except Outlook email (saves HTML for manual sending)

## Troubleshooting

### Common Issues

1. **File not found**: Check file paths and network connectivity
2. **Permission denied**: Ensure read access to Excel files
3. **Outlook not available**: Script will save HTML report for manual sending
4. **Data format errors**: Verify Excel file structure matches expected format

### Debug Mode

Enable debug logging by modifying the logging level:

```python
logging.basicConfig(level=logging.DEBUG)
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This script is provided as-is for internal Amazon use.

## Author

Vipin Dhiman (dhimv@amazon.com)

## Version History

- **v2.0**: Improved version with better error handling, logging, and cross-platform support
- **v1.0**: Original script with basic functionality
# Healthcare Provider Schedule Balancer

A Streamlit application that automatically balances healthcare provider chatting schedules according to specific business rules.

## Features

- **Automatic Balancing**: Balances provider schedules to ensure 24-hour coverage for each individual (DD, DM, OT)
- **Business Rule Compliance**: Implements priority-based balancing rules with exception handling
- **Color-Coded Output**: Visual indicators for all changes made during balancing
- **Excel Preservation**: Maintains original formatting while highlighting modifications
- **Summary Reporting**: Detailed log of all changes made during processing

## Business Rules Implementation

### Core Rules (Priority Order)
1. Each provider maximum: 16 hours per day (minimum: 2 hours)
2. Each individual total: exactly 24 hours per day
3. Provider supplements: Add supplemental providers when needed
4. Weekly oversight: Ensure oversight entries by authorized personnel

### Exception Handling
1. Modify non-zero entries for existing providers
2. Temporarily raise provider cap to 18 hours
3. Add "Carolyn Porter, LPN" entries (OT only)
4. Flag impossible-to-balance days with red highlighting

### Color Coding System
- **Red highlight**: Days that cannot be fully balanced
- **Green highlight**: New additions or entries changed from 0 to positive
- **Orange highlight**: Reduced or modified non-zero entries  
- **Green font**: Names of newly added providers

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
streamlit run app.py
```

## Usage

1. **Upload Excel File**: Upload your healthcare schedule Excel file (.xlsx format)
2. **Process Schedule**: Click the "Process Schedule" button to automatically balance
3. **Download Results**: Download the processed file with all balancing applied

## File Structure Expected

The application expects Excel files with a "DailyMatrix" sheet containing:
- Date rows (MM/DD/YYYY format)
- Provider rows with hours for DD, DM, OT columns
- Total hours and pending hours calculation rows
- Proper Excel formulas for calculations

## Technical Components

- `app.py`: Main Streamlit application
- `data_validator.py`: Excel file structure validation and parsing
- `schedule_balancer.py`: Core business logic implementation
- `excel_formatter.py`: Output formatting and Excel generation

## Requirements

- Python 3.6+
- Streamlit 1.10+
- pandas 1.1+
- openpyxl 3.1+
- xlsxwriter 3.0+

## Author

Built for healthcare schedule management according to specified business requirements.

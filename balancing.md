# Create a Streamlit App for Healthcare Provider Chatting Schedule Balance Management

Build a Streamlit application that processes Excel files containing healthcare provider chatting schedules and automatically balances them according to specific business rules. The app should handle data validation, automatic balancing, and generate a properly formatted output Excel file with color coding.

## Core Requirements

### 1. File Upload & Processing
- Accept Excel file upload through Streamlit file uploader
- Parse Excel data containing provider schedules with columns for dates, providers, and hours
- data/july_summary.xlsx is a sample of the actual input file. Evaluate its structure and formating and restart the task. Ensure you recreate the same structure or edit the existing excel
- Ensure to examine the existing formatting, formulas, and colors in the input Excel file to preserve them properly

### 2. Data Structure
- Handle scheduling data for 3 individuals (should be able to add more): DD, DM, OT
- Process daily entries with provider names and hours worked
- Track "Total hrs pending in a 24hr period" and "Provider Total" for each day

### 3. Core Balancing Rules (implement in order of priority)
- Each provider maximum: 16 hours per day (minimum: 2 hours)
- Reduce over-allocated providers first (bring them down to ≤16 hours)
- Each individual total: exactly 24 hours per day
- Provider supplements: Add "Charles Sagini, RN/House Manager", "Josephine Sagini, RN/Program Manager" or "Faith Murerwa, RN/House Supervisor" when totals are under-maxed
- Add supplemental providers or emergency providers to fill gaps
- You can add more than 1 supplemental provider per day if necessary
- Weekly oversight: Ensure there are 8-hour oversight entries by either "Josephine" or "Charles" for each individual
- Modify existing 0-hour entries for supplemental providers when possible

### 4. Exception Handling Rules (apply when standard balancing fails)
1. Modify any non-zero entry for any existing provider to fill the gap. If the supplemental providers are not maxed out, use them first, then consider for the existing ones even if they are not supplemental. Just don't make it zero.
2. Temporarily raise a core provider's per-day cap from 16 hrs to 18 hrs.
3. If still unbalanced, add "Carolyn Porter, LPN" entries—but only to supplement OT. Accept input from the user any second level additional providers to use for rebalancing
4. If absolutely impossible to balance to a total of 16 (or 18) hrs per provider and zero "hrs pending" for every 24-hr period, flag that day as "unbalanced." Mark impossible-to-balance days with red highlighting at the date entry cell.

### 5. Color Coding System
- **Red highlight**: Days that cannot be fully balanced
- **Green highlight**: Modified entries (new additions or changed from 0 to positive)
- **Orange highlight**: Reduced or modified non-zero entries
- **Green font**: Names of newly added providers for that day

### 6. User Interface Features
- File upload section
- All processing is done and ready for download. No preview of either input or output
- Download button for processed Excel file
- Summary statistics showing total adjustments made (Optional)

### 7. Output Requirements
- Generate Excel file with all color coding preserved
- Maintain original data structure while adding balancing modifications
- Maintain the formulas for totals
- Include summary sheet with balancing actions taken
- Ensure "Total hrs pending in a 24hr period" equals 0 for balanced days
- Ensure "Provider Total" equals 16 (or 18 in exceptions)

## Technical Specifications

- Use pandas for data manipulation
- Use openpyxl or xlsxwriter to apply cell-level formatting (fills, fonts)
- Organize your logic into clear functions (e.g., `balance_day()`, `apply_remedies()`, `highlight_cells()`)
- Ensure your Streamlit script runs with `streamlit run app.py` and works cross-platform
- Implement validation to ensure business rules are followed
- Include error handling for malformed data
- Make the interface intuitive for healthcare administrators

## Additional Features

- Add progress indicators for processing
- Data validation warnings before processing
- Export processing log showing all changes made

## Final Requirements

The app should be user-friendly for healthcare administrators who may not be technically savvy, with clear instructions and intuitive workflow.

**Deliver complete, production-ready code for `app.py` and any helper modules. Make sure to include comments explaining each major step.**
# Fund Analysis Excel Processor

A production-ready Python class to process fund analysis Excel files, calculate monthly stock returns and contributions, and generate pivot table summaries.

## Features

- Loads Excel files with comprehensive validation
- Calculates derived columns:
  - **Start Price**: Previous month's price for each instrument
  - **Monthly Stock Return%**: Calculated as `(Current Price / Start Price) - 1`
  - **Start wt%**: Previous month's holding percentage (converted to decimal)
  - **Stock Monthly Contribution %**: Product of Start wt% and Monthly Stock Return%
- Generates pivot tables for:
  - Company analysis (grouped by Instrument Name)
  - Sector analysis (grouped by Instrument Sector)
  - Market Cap analysis (grouped by Instrument SEBI Mcap Type)
- Exports results to Excel with multiple sheets

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Streamlit Web UI (Recommended)

The easiest way to use this tool is through the Streamlit web interface:

1. Install dependencies (if not already installed):
```bash
pip install -r requirements.txt
```

2. Run the Streamlit app:
```bash
streamlit run streamlit_app.py
```

3. The app will open in your browser. You can:
   - Upload your Excel file using the file uploader
   - Click "Process File" to process your data
   - Download the processed file with all calculated columns and pivot tables

### Basic Usage (Python API)

```python
from fund_analysis_processor import FundAnalysisProcessor

# Initialize processor
processor = FundAnalysisProcessor(
    input_file_path="path/to/input.xlsx",
    output_file_path="path/to/output.xlsx"  # Optional, auto-generated if not provided
)

# Run complete processing pipeline
output_file = processor.process()
```

### Step-by-Step Usage

```python
from fund_analysis_processor import FundAnalysisProcessor

# Initialize processor
processor = FundAnalysisProcessor(input_file_path="test.xlsx")

# Load and validate data
df = processor.load_data()

# Calculate derived columns
df = processor.calculate_derived_columns()

# Create pivot tables
pivot_tables = processor.create_pivot_tables()

# Save output to Excel
output_path = processor.save_output()
```

## Input File Requirements

The input Excel file must contain the following columns (in order):

1. `Scheme Code` - Numeric scheme IDs
2. `Scheme Name` - Scheme name text
3. `Month` - Month in YYYYMM format (e.g., 202506)
4. `Month End` - Month-end date
5. `Instrument Name` - Stock/instrument name
6. `Holding (%)` - Holding percentage
7. `Instrument Sector` - Sector/category
8. `Instrument SEBI Mcap` - SEBI market capitalization value
9. `Instrument SEBI Mcap Type` - Market cap type (Small Cap, Mid Cap, etc.)
10. `NSE Symbol` - NSE stock symbol
11. `Price` - Current month's closing price

## Output File Structure

The output Excel file contains 4 sheets:

1. **Processed Data**: All original columns plus calculated columns (L-O)
2. **Company Pivot**: Pivot table grouped by Instrument Name and Month End
3. **Sector Pivot**: Pivot table grouped by Instrument Sector and Month End
4. **Market Cap Pivot**: Pivot table grouped by Instrument SEBI Mcap Type and Month End

## Calculation Logic

### Start Price (Column L)
- For each instrument, the Start Price is the previous month's Price
- Blank/NaN for the first entry of each instrument

### Monthly Stock Return% (Column M)
- Calculated as: `(Current Price / Start Price) - 1`
- Blank/NaN if Start Price is missing or zero

### Start wt% (Column N)
- Previous month's Holding (%) converted to decimal
- Automatically handles percentage format conversion
- 0 for the first entry of each instrument

### Stock Monthly Contribution % (Column O)
- Calculated as: `Start wt% × Monthly Stock Return %`
- 0.00% if either input is 0 or blank

## Error Handling

The class includes comprehensive error handling for:
- File not found errors
- Missing required columns
- Invalid data types
- Division by zero in return calculations
- Empty dataframes
- Excel write permissions

## Logging

The class uses Python's logging module to track processing steps. Logs include:
- Data loading status
- Column validation results
- Calculation progress
- Pivot table generation
- Output file saving

## Example

```python
# Process the test file
processor = FundAnalysisProcessor(
    input_file_path=r"d:\car analysis\fund analysis\test.xlsx"
)

output_file = processor.process()
print(f"Output saved to: {output_file}")
```

## Requirements

- Python 3.8+
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- numpy >= 1.24.0, < 2.0.0
- streamlit >= 1.28.0 (for web UI)


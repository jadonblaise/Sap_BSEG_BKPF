<h2>SAPAnalyzer Class Documentation</h2>
Overview
SAPAnalyzer is a Python class designed to load, merge, validate, and compare SAP financial data from the BSEG and BKPF tables against a provided summary balance sheet. The goal is to ensure the SAP data totals per general ledger account (hauptbuch) align with the summary sheet’s reported balances (endsaldo).

Key Functionalities
Initialization
Accepts file paths for three input data sources:

BSEG (SAP line item data)

BKPF (SAP header data)

Summary (external balance sheet in Excel)

Data Loading (load_data)
Loads BSEG and BKPF CSV files with appropriate encoding and delimiter.

Loads the summary Excel sheet.

Standardizes column names by stripping whitespace and lowercasing for consistency.

Handles exceptions during file reading with error messages.

Data Merging (merge_tables)
Merges BSEG and BKPF datasets on the key column belegnr (document number).

Uses inner join to retain only matching records.

Data Validation (validate_data)
Checks for missing belegnr columns or null keys.

Identifies unmatched records by performing an outer merge with indicator.

Warns if BSEG lines lack BKPF headers or vice versa, highlighting potential data issues.

Summary Comparison (compare_summary)
Validates presence of required columns (betrag hauswähr, hauptbuch, endsaldo).

Normalizes and cleans amount columns by replacing thousand separators and decimal marks to parse floats correctly.

Aggregates betrag hauswähr from merged data by hauptbuch to compute balances.

Normalizes hauptbuch account numbers to a 6-digit zero-padded string format.

Merges computed balances with summary balances, preserving the original summary row order.

Calculates the difference between computed and summary balances.

Stores comparison results in self.comparison_df.

Outputs a trimmed comparison with hauptbuch and computed balance for verification.

Export Results (export_results)
Exports the full merged SAP data to an Excel sheet named AnalyzedData.

Exports a focused comparison table with only hauptbuch and endsaldo_computed to a sheet named BalancedComparison.

Provides confirmation messages on export success.

Usage
The class is instantiated with file paths to input data. Calling the following methods in order runs the full analysis and export:

python
Copy
Edit
Analyzer.load_data()
Analyzer.merge_tables()
Analyzer.validate_data()
Analyzer.compare_summary()
Analyzer.export_results("Abstimmung.xlsx")
Notes
The class is designed to be robust to missing or malformed data entries.

Column normalization and string cleaning ensure compatibility across SAP export formats.

Summary sheet rows with no matching SAP data will show NaN in computed balances.

The export limits output columns to those most relevant for review.

# Project Report: Data Cleaning & Reporting Automation

## Introduction
Spreadsheet datasets often contain missing values, duplicate rows, inconsistent text, and unclear column names. Manual cleaning is repetitive and error-prone. CleanReport Studio makes common cleanup and reporting tasks repeatable.

## Problem Statement
Build an interactive application that accepts CSV/Excel files, identifies data-quality issues, applies configurable cleaning rules, visualizes data, and exports useful reports.

## Objectives
- Inspect quality before cleaning.
- Handle missing values and duplicate records.
- Standardize column names and text formatting.
- Display before/after statistics.
- Generate cleaned CSV and Excel reports that can be opened in Excel or imported into Power BI.

## Tools
- Python: core language.
- Pandas: loading, transforming, and profiling tables.
- Streamlit: interactive browser UI.
- OpenPyXL: formatted multi-sheet Excel output.
- Matplotlib: numeric visualizations.
- NumPy: numeric data support.

## Workflow
1. Upload CSV/Excel.
2. Load into a Pandas DataFrame.
3. Profile missing cells, unique values, and duplicate rows.
4. Apply user-selected cleanup rules.
5. Compare before/after metrics.
6. Visualize numeric distributions.
7. Download CSV and Excel report.

## Modules
- File reader: reads supported file types.
- Data profiler: calculates per-column quality metrics.
- Cleaning engine: standardizes names, trims text, fills/drops missing values, removes duplicates.
- Visualization module: charts numeric distributions and remaining missing values.
- Report generator: exports cleaned data, quality profile, summary, and applied actions.

## Testing Checklist
- Upload valid CSV and XLSX files.
- Confirm missing values and duplicate rows are reported.
- Verify selected cleanup options change output as expected.
- Download the cleaned CSV.
- Open the Excel report and verify its four worksheets.
- Test an empty or unsupported file and confirm a helpful message appears.

## Limitations
- Automated rules cannot infer the business meaning of every column; review output.
- Power BI integration is through exported files, not a direct API.
- The prototype handles one dataset at a time and has no user accounts/database.

## Future Enhancements
- Combine multiple input files.
- Add per-column cleaning rules, PDF output, scheduled jobs, and audit history.
- Add direct database connectors and a Power BI refresh workflow.

## Conclusion
CleanReport Studio demonstrates an end-to-end data-cleaning workflow with quality inspection, configurable transformations, visual summaries, and automated report generation.

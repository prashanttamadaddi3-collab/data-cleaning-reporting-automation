# CleanReport Studio — Data Cleaning & Reporting Automation

A Python project that cleans CSV/Excel files, profiles data quality, visualizes numeric data, and exports formatted Excel reports.

## Features
- Upload CSV, XLSX, or XLS data.
- Profile missing values, unique values, duplicates, and data types.
- Standardize column names, trim text, change text case, remove duplicates, and handle missing values.
- Compare before/after metrics and view numeric distributions.
- Download cleaned CSV and a multi-sheet Excel report (`Cleaned Data`, `Data Quality`, `Summary`, `Actions`).
- Includes a deliberately imperfect sample dataset.

## Technologies
Python, Streamlit, Pandas, NumPy, Matplotlib, OpenPyXL.

## Run on Windows / VS Code
1. Extract the ZIP and open the extracted folder in VS Code.
2. In the terminal, make sure you are in the folder containing `app.py`.
3. Recommended virtual environment:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```
4. Install packages:
   ```powershell
   python -m pip install -r requirements.txt
   ```
5. Run:
   ```powershell
   python -m streamlit run app.py
   ```
6. Open the local URL printed by Streamlit, usually `http://localhost:8501`.

## Demo
Upload `data/sample_sales.csv`, inspect missing/duplicate counts, apply cleaning rules, compare metrics, and download the generated report.

## Workflow
Upload → Profile → Choose rules → Clean → Visualize → Export CSV/Excel.

## Power BI
The generated CSV and Excel files can be imported into Power BI. This starter project does not automatically refresh Power BI or connect to its APIs.

## Important notes
- Review cleaning rules before using results for business decisions.
- Mean/median/mode are general-purpose defaults; the best approach depends on the meaning of each column.
- Avoid uploading sensitive data to an untrusted hosted instance.

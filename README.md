
# Commercial Data Analysis (Python -> Excel Dashboard)


## Overview
Year-to-date (YTD) compliance reporting pack built from a SharePoint List containing commercial sales audit records. Python (pandas) cleans and reshapes the export, calculates KPIs, and outputs an Excel file with multiple sheets plus a one-page dashboard for stakeholders.


## Problem
- Compliance needed a YTD view of FCA-related metrics (vulnerability, remediation, selected question pass rates).
- Raw SharePoint exports contained many non-compliance fields and required significant cleaning/reshaping before reporting.
- Manual filtering in Excel would be time-consuming and inconsistent.


## Approach
- **Extract:** exported SharePoint List audit data for the required YTD scope.
- **Transform (Python/pandas):** cleaned fields, handled types, reshaped compliance question results into reporting-friendly tables, and generated KPI outputs.
- **Deliver (Excel):** wrote cleaned datasets + KPI tables to an Excel workbook (multiple sheets) and built a one-page dashboard with charts and headline KPIs.


## Tools
- **Python (pandas, openpyxl)**
- **SharePoint**
- **Excel**


## Results
- Delivered a clean YTD compliance summary with the requested FCA-focused metrics in a single file.
- Reduced manual effort and improved consistency by automating the cleaning and KPI generation steps.
- Provided a simple dashboard view for quick stakeholder reporting.

## Files
- `Commercial Data 2025.xlsx` – cleaned datasets, KPI tables and dashboard
- `Commercial_Dashboard_Screenshot.PNG` – dashboard screenshot
- `Commercial_Clean_KPI_Script.py` – Python cleaning + KPI generation script


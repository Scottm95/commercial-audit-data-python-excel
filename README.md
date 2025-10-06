
# Commercial Data Analysis (Python -> Excel Dashboard)

## Business Problem
The compliance team requested a **year-to-date (YTD) view** of sales audits, focuing on specific FCA related metrics such as customer vulnerability, remediation rates and certain compliance question pass rates. 
 
All Commerical audit data is stored in a **Sharepoint List**, which contains thousands of records.
However, the raw export includes all audit information (not just the compliance related fields) and significant cleaning and transformation was required prior to creating the report. 


## Solution
- **Python and Pandas** - used to clean raw SharePoint data and prepare **YTD Metrics**.
- Generated **key KPIs**:
- Total Audits YTD
- Vulnerable Customer % YTD
- Remediation % YTD
- Pass rates for selected compliance questions
- Device orders, interaction types, customer type splits
- Reshaped compliance question results into tidy tables for easier reporting.
- Exported all cleaned datasets and KPI results into **Excel with multiple sheets**. 
- Built a **one-page Excel dashboard** with charts and KPI highlights, designed to be FCA-reporting friendly.


## Impact
- Delivered a **clean, year-to-date summary** with key information included to make reporting easier.
- Automated a process that would otherwise require manual filtering and transformations in Excel. 
- Produced a professional, single-page dashboard that gives compliance and management quick access to the requested insights.  


## Files
- 'Commercial Data 2025.xslx' - Contains cleaned dataset, KPIs and dashboard.
- 'Commercial_Dashboard_Screenshot.PNG' - Dashboard screenshot
- 'README.md' - project documentation.
- 'Commercial_Clean_KPI_Script.py' - Python script (data cleaning + KPI generation)


## Tech Stack
- Python (pandas, openyxl)

- Excel (dashboard creation, charts, formatting)

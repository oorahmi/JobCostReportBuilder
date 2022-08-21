# Job Cost Report Builder
Program to take in a .xslx for the total job workbook and Revenue report workbook file and fill out all Job Cost Report Sheets.
Currently requires the files to be in the JobCostReportBuilder directory

### Installation
To download navigate to the top and download under green 'code' button.

or in the terminal

    git clone https://github.com/oorahmi/jobCostReportBuilder.git
    
Requires openpyxl, in your terminal run:

    pip install openpyxl

Usage:  

    python createJobWorkbook.py path_to_job_workbook_file   path_to_revenue_report_workbook-file

The processed total job workbook will be saved as a copy in /processed



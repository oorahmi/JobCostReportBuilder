# Job Cost Report Builder
Program to take in a .xslx for the total job workbook and Revenue report workbook file and fill out all Job Cost Report Sheets.

* Currently requires the input .xslx files to be in the JobCostReportBuilder directory

### Installation
**- Written in python 3.10x**

To download navigate to the top and download under green 'code' button.

or in the terminal

    git clone https://github.com/oorahmi/jobCostReportBuilder.git

Requires openpyxl, in your terminal run:

    pip install openpyxl
    
### Job Cost Report
Usage:  

    python createJobWorkbook.py path_to_cost_detail_workbook_file   path_to_revenue_report_workbook-file

The processed total job workbook will be saved as a copy in /processed


### EVA Cost report

Usage: 

    python createEVAJobWorkbook.py path_to_eva_total_workbook_file  
    
The processed total job workbook will be saved as a copy in /processed

### WIP report

Usage:

    python createWIPReport.py path_to_processed_eva_workbook quarter_to_process (1-4) year
    
The processed WIP report will be saved in /processed

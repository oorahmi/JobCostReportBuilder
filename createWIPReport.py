'''
Program to take in a eva report and output a wip report

Usage:  

    python createWIPReport.py path_to_eva_workbook_file 

The wip report will be saved in /processed

Requires openpyxl

    pip install openpyxl

Author: Brian Wright
10-19-2022

'''
from xmlrpc.client import MAXINT
from util import draw_line, set_border, copySheet

import sys
import openpyxl
from openpyxl.styles import Font
import os 
import shutil
from copy import copy
from datetime import datetime
from collections import OrderedDict
       

def createWIPReport(eva_wb_path):
    eva_total_wb = openpyxl.load_workbook(eva_wb_path) 
    if not eva_total_wb:
        print("Error: failed to open workbook: ", eva_wb_path)
        return

    # copy empty wip
    wip_blank_path = os.getcwd() + "/data/wip_blank.xlsx"
    processed_file_path = os.getcwd() +'/processed/wip_processed.xlsx'
    shutil.copyfile(wip_blank_path, processed_file_path)

    wip_report_wb = openpyxl.load_workbook(wip_blank_path) 
    if not wip_report_wb:
        print("Error: failed to open data workbook: ", wip_blank_path)
        sys.exit()
        return

    actual_cost_detail_sheet    = eva_total_wb.worksheets[0]
    revenue_sheet               = eva_total_wb.worksheets[1]
    estimate_cost_detail_sheet  = eva_total_wb.worksheets[2]

    wip_report_sheet = wip_report_wb.worksheets[0]

    # define all columns
    WIP_JOB_NUMBER_COlUMN = 1
    WIP_JOB_NAME_COlUMN = 2
    WIP_CONTRACT_PRICE_COlUMN = 4
    WIP_APPROVED_CO_COlUMN = 5
    WIP_TOTAL_PRICE_COlUMN = 6
    WIP_ESTIMATED_COST_COlUMN = 7
    WIP_ESTIMATED_PROFIT_COlUMN = 8
    WIP_ESTIMATED_PROFIT_PERC_COlUMN = 9
    WIP_ACTUAL_COST_TO_DATE_COlUMN = 10
    WIP_PERC_COMPLETION_COlUMN = 11
    WIP_REVENUES_EARNED_TO_DATE_COlUMN = 12
    WIP_BILLINGS_TO_DATE_COlUMN = 13
    WIP_RETAINAGE_COlUMN = 14
    WIP_COST_IN_EXEC_BILLINGS_COlUMN = 15
    WIP_BILLINGS_IN_EXCESS_COlUMN = 16
    WIP_BACKLOG_COlUMN = 17
    WIP_Q_REVENUES_COlUMN = 18
    WIP_Q_COSTS_COlUMN = 19
    WIP_Q_LAB_OVERHEAD_COlUMN = 20
    WIP_Q_OTH_OVERHEAD_COlUMN = 21
    WIP_Q_PROFIT_LOSS_COlUMN = 22


    fl_job_numbers = []
    m_fl_job_numbers = []
    other_job_numbers = []
    # find order of jobs
    # get all food lion jobs first, sort ascending, then rest sort ascending
    for i in range(3, len(eva_total_wb.worksheets)):
        job_sheet = eva_total_wb.worksheets[i]
        job_desc = job_sheet.cell(row = 2, column = 1).value
        job_number = job_sheet.title
        if "Food Lion" in job_desc:
            if "Maintenance" in job_desc:
                m_fl_job_numbers.append(job_number)
            else:
                fl_job_numbers.append(job_number)
        else:
            other_job_numbers.append(job_number)

    fl_job_numbers.sort()
    m_fl_job_numbers.sort()
    other_job_numbers.sort()
    
    # build sorted list from ordered job numbers
    ordered_job_numbers = fl_job_numbers + m_fl_job_numbers + other_job_numbers

    # NOTE: dont need to store these seperately but nicer to work with?
    ordered_sheets = []
    for job_number in ordered_job_numbers:
        # get correct sheet
        for i in range(3, len(eva_total_wb.worksheets)): 
            job_sheet = eva_total_wb.worksheets[i]
            if job_number == job_sheet.title:
                ordered_sheets.append(job_sheet)
    
    # write job data
    i = 7 # after header
    for job_sheet in ordered_sheets:
        
        i = str(i)
        # write formulaic columns
        wip_report_sheet.cell(row = i, column = WIP_TOTAL_PRICE_COlUMN).value = "=D" + i + "+E" + i
        wip_report_sheet.cell(row = i, column = WIP_ESTIMATED_PROFIT_COlUMN).value = "=+F" + i + "-G" + i
        wip_report_sheet.cell(row = i, column = WIP_ESTIMATED_PROFIT_PERC_COlUMN).value = "=H" + i + "/G" + i
        wip_report_sheet.cell(row = i, column = WIP_PERC_COMPLETION_COlUMN).value = "=+J" + i + "/G" + i
        wip_report_sheet.cell(row = i, column = WIP_REVENUES_EARNED_TO_DATE_COlUMN).value = "=F" + i + "*K" + i
        wip_report_sheet.cell(row = i, column = WIP_COST_IN_EXEC_BILLINGS_COlUMN).value = '=IF(L'+i+'>M'+i+',L'+i+'-M'+i+',IF((L'+i+'-M'+i+')<-1,"-",0))'
        wip_report_sheet.cell(row = i, column = WIP_BILLINGS_IN_EXCESS_COlUMN).value = '=IF(L'+i+'>M'+i+',"-",L'+i+'-M'+i+')'
        wip_report_sheet.cell(row = i, column = WIP_BACKLOG_COlUMN).value = "=+F"+i+"-L"+i
        wip_report_sheet.cell(row = i, column = WIP_Q_PROFIT_LOSS_COlUMN).value = "=R"+i+"-S"+i
        i = int(i)

        job_number = job_sheet.title

        job_name = job_sheet.cell(row = 2, column = 1).value
        # trimming off "Job Estimates vs. Actuals Detail for "
        job_name = job_name[37:]
        name_index = job_name.find(job_number) + 6 # job number length
        job_name = job_name[name_index:]
        if "Food Lion" in job_name:
            if job_number in m_fl_job_numbers:
                wip_report_sheet.cell(row = i, column = WIP_JOB_NAME_COlUMN).value = "Maint. FL" + job_number[3:] + " " + job_name
                wip_report_sheet.cell(row = i, column = WIP_JOB_NUMBER_COlUMN).value = "M" + job_number
            else:
                # only write 4 digits of the job number and the rest of the text.
                wip_report_sheet.cell(row = i, column = WIP_JOB_NAME_COlUMN).value = "FL" + job_number[3:] + " " + job_name
                wip_report_sheet.cell(row = i, column = WIP_JOB_NUMBER_COlUMN).value = job_number
        else:
            # only write the rest of the text after job number 
            wip_report_sheet.cell(row = i, column = WIP_JOB_NAME_COlUMN).value = job_name


        JOB_ESTIMATE_COLUMN = 5
        estimate_total = 0
        change_order = 0
        orig_contract = 0
        other_job_income = 0
        billed_to_date = 0
        total_cost_w_oh = 0
        retainage = 0
        labor_oh = 0
        other_oh = 0


        # get several numbers from job sheet
        for j in range(7, job_sheet.max_row + 1): # could optimize by not doing all rows

            if job_sheet.cell(row = j, column = 2).value == "Total":
                estimate_total = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 3).value == "Orig Contract":
                orig_contract= job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 3).value == "Change Order":
                change_order = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 3).value == "Other Job Income":
                other_job_income = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 3).value == "Total Billed to Date":
                billed_to_date = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 4).value == "Total Cost w/ OH":
                total_cost_w_oh = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 3).value == "Retainage Held by Customer":
                retainage = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 4).value == "Labor OH":
                labor_oh = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value
            elif job_sheet.cell(row = j, column = 4).value == "Other OH":
                other_oh = job_sheet.cell(row = j, column = JOB_ESTIMATE_COLUMN).value

        
        if estimate_total > 0:
            wip_report_sheet.cell(row = i, column = WIP_CONTRACT_PRICE_COlUMN).value = estimate_total 
            wip_report_sheet.cell(row = i, column = WIP_ESTIMATED_COST_COlUMN).value = estimate_total - orig_contract + (change_order*.95)
            wip_report_sheet.cell(row = i, column = WIP_APPROVED_CO_COlUMN).value = change_order 
        else: # no estimate
            wip_report_sheet.cell(row = i, column = WIP_CONTRACT_PRICE_COlUMN).value = estimate_total 
            wip_report_sheet.cell(row = i, column = WIP_ESTIMATED_COST_COlUMN).value = 


        # fill in simple columns
        wip_report_sheet.cell(row = i, column = WIP_ACTUAL_COST_TO_DATE_COlUMN).value = total_cost_w_oh 
        wip_report_sheet.cell(row = i, column = WIP_BILLINGS_TO_DATE_COlUMN).value = billed_to_date 
        wip_report_sheet.cell(row = i, column = WIP_RETAINAGE_COlUMN).value = retainage 
        wip_report_sheet.cell(row = i, column = WIP_Q_LAB_OVERHEAD_COlUMN).value = labor_oh 
        wip_report_sheet.cell(row = i, column = WIP_Q_OTH_OVERHEAD_COlUMN).value = other_oh
        

        i += 1 # next job

                


    # -------------------------------------------------------------------------------- #



def main(argv):
    if len(argv) == 0 or len(argv) > 1:
        print("Error - usage: supply the processed EVA Job Workbook")
        return

    eva_wb_path = os.path.abspath(argv[0])

    if os.path.isfile(eva_wb_path):
        createWIPReport() 
    elif not os.path.isfile():
        print("Error: eva workbook path does not exist?: ", eva_wb_path)
    else:
        print("Error: wrong input")

if __name__ == "__main__":
   main(sys.argv[1:])

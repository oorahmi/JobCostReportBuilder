'''
Program to take in a .xslx for the estimated cost detail, actual cost detail, and Revenue report workbook file and fill out all Job Cost Report Sheets.

Usage:  

    python createEVAJobWorkbook.py path_to_eva_workbook_file 

The processed total job workbook will be saved as a copy in /processed

Requires openpyxl

    pip install openpyxl

Author: Brian Wright
9-6-2022

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
    # copy wb and work on copy
    processed_file_path = os.path.split(eva_wb_path)[0]  +'/processed/' + os.path.basename(eva_wb_path).split('.')[0] + '_processed.xlsx'
    shutil.copyfile(eva_wb_path, processed_file_path)

    eva_total_wb = openpyxl.load_workbook(processed_file_path) 
    if not eva_total_wb:
        print("Error: failed to open workbook: ", processed_file_path)
        return

    actual_cost_detail_sheet    = eva_total_wb.worksheets[0]
    revenue_sheet               = eva_total_wb.worksheets[1]
    estimate_cost_detail_sheet  = eva_total_wb.worksheets[2]

    #job_str_set = set()
    job_str_set = OrderedDict()
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

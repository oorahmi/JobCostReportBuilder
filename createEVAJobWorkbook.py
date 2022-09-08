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
from util import set_border, copySheet

import sys
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
import os 
import shutil
from copy import copy
from datetime import datetime
from collections import OrderedDict
       

def createEVAJobWorkbook(eva_total_wb_path):
    # copy wb and work on copy
    processed_file_path = os.path.split(eva_total_wb_path)[0]  +'/processed/' + os.path.basename(eva_total_wb_path).split('.')[0] + '_processed.xlsx'
    shutil.copyfile(eva_total_wb_path, processed_file_path)

    eva_total_wb = openpyxl.load_workbook(processed_file_path) 
    if not eva_total_wb:
        print("Error: failed to open workbook: ", processed_file_path)
        return

    actual_cost_detail_sheet    = eva_total_wb.worksheets[0]
    revenue_sheet               = eva_total_wb.worksheets[1]
    estimate_cost_detail_sheet  = eva_total_wb.worksheets[2]

    #job_str_set = set()
    job_str_set = OrderedDict()

    ACTUAL_NAME_COLUMN = 11      

    # add new sheet for each unique job
    # column
    for i in range(1, actual_cost_detail_sheet.max_row + 1): 
        job_data = actual_cost_detail_sheet.cell(row = i, column = ACTUAL_NAME_COLUMN).value

        # format is currrently:   job_name:job_number type
        # is a job string? 
        if job_data and len(job_data.split(":")) > 1:
            job_number = job_data.split(":")[1].split(' ')[0]    # could be better?
            job_str_set[job_number] = None

    job_numbers = list(job_str_set.keys())
    
    if len(job_numbers) == 0:
        print("Error: failed to find jobs in workbook: ", processed_file_path)
        return


    # add new sheet for each job number
    for job_number in job_numbers:
        eva_total_wb.create_sheet(title=job_number)

    # copy empty job cost sheet
    eva_jc_wb = openpyxl.load_workbook(os.getcwd() + "/data/eva_jc_blank.xlsx") 
    if not eva_jc_wb:
        print("Error: failed to open data workbook: /data/eva_jc_blank.xlsx")
        sys.exit()
        return

    ACTUAL_DATE_COLUMN = 9
    ACTUAL_ITEM_COLUMN = 15      
    ACTUAL_AMOUNT_COLUMN = 19

    class JobItem:
        def __init__(self, item_name=""):
            self.item_name = item_name
            self.actual_amount    = 0  # used for tracking value for non sub-type
            self.estimate_amount = 0  

            self.hasSub = False
            self.actual_sub_items   = OrderedDict() 
            self.estimate_sub_items = OrderedDict() # contains list of (name, amount) pairs
        
        def processSubItem(self, sub_item_name, amount, actual=False):
            if not self.hasSub:
                print("error: tried to process subitem on top level item")
                return

            if sub_item_name not in self.actual_sub_items.keys():
                # initialize sub-item in both dicts
                if actual:
                    self.actual_sub_items[sub_item_name] = amount 
                    self.estimate_sub_items[sub_item_name] = 0
                else:
                    self.actual_sub_items[sub_item_name] = 0 
                    self.estimate_sub_items[sub_item_name] = amount 
            else: # already in?
                if actual:
                    self.actual_sub_items[sub_item_name] += amount
                else:
                    self.estimate_sub_items[sub_item_name] += amount
            
    # SCOPED
    def createEVAJobCostSheet(sheet):
        # create job sheet
        job_number = sheet.title

        min_date = datetime.max
        max_date = datetime.min

        job_name = None
        job_items = []
        # get all job actual costs 
        for i in range(1, actual_cost_detail_sheet.max_row + 1):    # could optimize by not doing all rows
            j_name = actual_cost_detail_sheet.cell(row = i, column = ACTUAL_NAME_COLUMN).value

            if j_name and job_number in j_name:

                date = actual_cost_detail_sheet.cell(row = i, column = ACTUAL_DATE_COLUMN).value
                if date:
                    min_date = min(min_date, date)
                    max_date = max(max_date, date)
                else:
                    print("Warn: Job without a date: ", j_name)

                if not job_name:
                    job_name = j_name 

                j_item = actual_cost_detail_sheet.cell(row = i, column = ACTUAL_ITEM_COLUMN).value
                j_actual_amount = actual_cost_detail_sheet.cell(row = i, column = ACTUAL_AMOUNT_COLUMN).value
                if not j_item:
                    print("Warn: have job entry with no item data", j_name)
                    continue

                if not j_actual_amount:
                    print("Warn: have job entry with no amount, ", j_name)
                    continue

                item_name = ""
                sub_item_name = None
                # won't have sub item types without :
                if ":" not in j_item:
                    item_name = j_item
                elif ":" in j_item:
                    item_name, sub_item_name = j_item.split(":")
                else:
                    print("Warn: unhandled job item: ", j_item)

                job_item = None
                # find job_item if it already exists
                for j_item in job_items:
                    if j_item.item_name == item_name:
                        job_item = j_item
                if not job_item:
                    job_item = JobItem(item_name)
                    if sub_item_name:
                        job_item.hasSub = True
                        job_item.processSubItem(sub_item_name, j_actual_amount, actual=True)
                    else:
                        job_item.actual_amount += j_actual_amount
                    job_items.append(job_item)
                else: # 
                    if job_item.hasSub:
                        job_item.processSubItem(sub_item_name, j_actual_amount, actual=True)
                    else:
                        job_item.actual_amount += j_actual_amount

        ESTIMATE_NAME_COLUMN   = 10
        ESTIMATE_ITEM_COLUMN   = 12
        ESTIMATE_AMOUNT_COLUMN = 16

        # now get estimate amounts for all the job items, should not be any new jobs
        for i in range(1, estimate_cost_detail_sheet.max_row + 1):    # could optimize by not doing all rows
            j_name = estimate_cost_detail_sheet.cell(row = i, column = ESTIMATE_NAME_COLUMN).value

            if j_name and job_number in j_name:

                j_item = estimate_cost_detail_sheet.cell(row = i, column = ESTIMATE_ITEM_COLUMN).value
                j_estimate_amount = estimate_cost_detail_sheet.cell(row = i, column = ESTIMATE_AMOUNT_COLUMN).value

                if not j_item:
                    print("Warn: have job entry with no item data", j_name)
                    continue

                if not j_estimate_amount:
                    print("Warn: have job entry with no estimate amount, ", j_name)
                    continue

                item_name = ""
                sub_item_name = None
                # won't have sub item types without :
                if ":" not in j_item:
                    item_name = j_item
                elif ":" in j_item:
                    item_name, sub_item_name = j_item.split(":")
                else:
                    print("Warn: unhandled job item: ", j_item)

                job_item = None
                # find job_item if it already exists
                for j_item in job_items:
                    if j_item.item_name == item_name:
                        job_item = j_item
                if not job_item:
                    print("Warn: found job with an estimate but no actual cost, ", item_name, " ", sub_item_name)
                    job_item = JobItem(item_name)
                    if sub_item_name:
                        job_item.hasSub = True
                        job_item.processSubItem(sub_item_name, j_estimate_amount, actual=False)
                    else:
                        job_item.estimate_amount += j_estimate_amount
                    job_items.append(job_item)


                if job_item.hasSub:
                    job_item.processSubItem(sub_item_name, j_estimate_amount, actual=False)
                else:
                    job_item.estimate_amount += j_estimate_amount


                        
        # append job name at top text
        sheet.cell(row = 2, column = 1).value = sheet.cell(row = 2, column = 1).value + " " + job_name

        # write date range
        sheet.cell(row = 3, column = 1).value = "Transactions from: " + min_date.strftime("%m/%d/%y") + " to " + max_date.strftime("%m/%d/%y")

        ITEM_NAME_COLUMN        = 3
        SUBITEM_NAME_COLUMN     = 4
        ESTIMATE_COST_COLUMN    = 5
        ACT_COST_COLUMN         = 7
        DIFF_COLUMN             = 9

        # date and time
        sheet.cell(row = 1, column = DIFF_COLUMN).value = datetime.today().strftime("%H:%M %p")
        sheet.cell(row = 2, column = DIFF_COLUMN).value = datetime.today().strftime("%B %d, %Y")

        i = 7  # starting point after 'Service' row

        total_actual_labor_cost    = 0
        total_estimate_labor_cost  = 0
        total_labor_cost_no_temp   = 0

        total_estimate_cost = 0
        total_actual_cost   = 0

        # write job cost data
        for job_item in job_items:
            if not job_item.hasSub:

                if "labor" in job_item.item_name.lower() and "temp" not in job_item.item_name.lower():
                    total_labor_cost_no_temp += job_item.actual_amount
                    total_actual_labor_cost += job_item.actual_amount
                    total_estimate_labor_cost += job_item.estimate_amount
                elif "labor" in job_item.item_name.lower():
                    total_actual_labor_cost += job_item.actual_amount
                    total_estimate_labor_cost += job_item.estimate_amount

                total_actual_cost += job_item.actual_amount
                total_estimate_cost += job_item.estimate_amount
                diff = job_item.estimate_amount - job_item.actual_amount

                sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = job_item.item_name 
                sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = job_item.estimate_amount
                sheet.cell(row = i, column = ACT_COST_COLUMN).value = job_item.actual_amount
                sheet.cell(row = i, column = DIFF_COLUMN).value = diff 
                i += 1
            else:
                sheet.cell(row = i, column = 3).value = job_item.item_name 
                i += 1
                sub_actual_total = 0
                sub_estimate_total = 0
                for s_item_name, s_actual_amount in job_item.actual_sub_items.items():
                    
                    s_estimate_amount = job_item.estimate_sub_items[s_item_name] 

                    if "labor" in s_item_name.lower() and "temp" not in s_item_name.lower():
                        total_labor_cost_no_temp += s_actual_amount
                        total_actual_labor_cost += s_actual_amount
                        total_estimate_labor_cost += s_estimate_amount
                    elif "labor" in s_item_name.lower():
                        total_actual_labor_cost += s_actual_amount
                        total_estimate_labor_cost += s_estimate_amount

                    total_actual_cost += s_actual_amount
                    total_estimate_cost += s_estimate_amount
                    sub_actual_total += s_actual_amount
                    sub_estimate_total += s_estimate_amount
                    s_diff = s_estimate_amount - s_actual_amount

                    sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = s_item_name
                    sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = s_estimate_amount
                    sheet.cell(row = i, column = ACT_COST_COLUMN).value = s_actual_amount
                    sheet.cell(row = i, column = DIFF_COLUMN).value = s_diff
                    i += 1
                # write out total for the subs
                sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Total " + job_item.item_name
                sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = sub_estimate_total
                sheet.cell(row = i, column = ACT_COST_COLUMN).value = sub_actual_total
                sheet.cell(row = i, column = DIFF_COLUMN).value = sub_estimate_total - sub_actual_total
                i += 1

        # total service, same as total??
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Total Service"
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_estimate_cost
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_actual_cost 
        sheet.cell(row = i, column = DIFF_COLUMN).value = total_estimate_cost - total_actual_cost
        i += 1

        # Other Charges
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Other Charges"
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = 0.0
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = 0.0
        sheet.cell(row = i, column = DIFF_COLUMN).value = 0.0
        i += 1

        # total
        sheet.cell(row = i, column = 2).value = "Total"
        sheet.cell(row = i, column = 2).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_estimate_cost
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_actual_cost
        sheet.cell(row = i, column = DIFF_COLUMN).value = total_estimate_cost - total_actual_cost
        # total font bold
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = DIFF_COLUMN).font = Font(bold=True)
        i += 1
        # whitespace
        i += 1

        sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Estimated Contract Revenue"
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_estimate_cost
        i += 1

        sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Actual Revenue To Date"
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).font = Font(bold=True)
        i += 1

        NAME_REVENUE_COLUMN = 11
        MEMO_REVENUE_COLUMN = 13
        ITEM_REVENUE_COLUMN = 15
        AMOUNT_REVENUE_COLUMN = 19
 
        # get revenue info 

        total_orig_contract    = 0
        total_change_order     = 0
        total_other_job_income = 0
        total_revenue          = 0
        total_retainage        = 0

        for j in range(6, revenue_sheet.max_row + 1):    
            j_name = revenue_sheet.cell(row = j, column = NAME_REVENUE_COLUMN).value
            if j_name and job_number in j_name:
                amount = revenue_sheet.cell(row = j, column = AMOUNT_REVENUE_COLUMN).value
                item_str = revenue_sheet.cell(row = j, column = ITEM_REVENUE_COLUMN).value
                if not item_str:
                    continue
                item_str = item_str.lower()

                if "orig contract" in item_str:
                    total_orig_contract += amount
                elif "change order" in item_str:
                    total_change_order += amount
                elif "other job income" in item_str:
                    total_other_job_income += amount
                elif "retainage" in item_str:
                    total_retainage += amount
                
                # some amounts are negative
                total_revenue += amount

        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Orig Contract"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_orig_contract
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Change Order"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_change_order 
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Other Job Income"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_other_job_income 
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total Revenue Recognized to Date"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_revenue
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Retainage Held by Customer"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = -total_retainage
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total Actual Revenue Collected to Date"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_revenue - total_retainage
        i += 1
        i += 1 # whitespace

        sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Total Estimated Labor including Temp Labor"
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_estimate_labor_cost
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = "% Complete"
        i += 1
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Total Actual Labor including Temp Labor"
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_actual_labor_cost
        if total_estimate_labor_cost > 0:
            sheet.cell(row = i, column = ACT_COST_COLUMN).value = str(round((total_actual_labor_cost/total_estimate_labor_cost) * 100, 2)) + "%"
        else:
            sheet.cell(row = i, column = ACT_COST_COLUMN).value = str(0.0) + "%"
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Estimated vs Actual Difference"
        sheet.cell(row = i, column = ITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_estimate_labor_cost - total_actual_labor_cost
        i += 1
        i += 1 # whitespace

    
        # summary box
        '''
        Calculation details for the summary box:
        Total Labor: Add all labor costs except temp labor
        Labor OH: Multiply 30% to the total labor calculated above
        Other OH: Multiply 0.005 to all costs
        Total Cost w/OH: Total Costs + Labor OH + Other OH
        '''
        labor_oh = total_labor_cost_no_temp * 0.3
        other_oh = total_actual_cost * 0.005
        total_cost_w_oh = total_actual_cost + labor_oh + other_oh

        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total Labor"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_labor_cost_no_temp 
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Labor OH"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = labor_oh
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Other OH"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = other_oh
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total Cost w/ OH"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_cost_w_oh
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Billed To Date"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_revenue
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).font = Font(bold=True)
        i += 1

        # column letter row number : column letter row number  for top left, bottom right
        cell_range = "D" + str(i-5) + ":E" + str(i-1)
        set_border(sheet, cell_range)

        # whitespace
        i += 1

        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Billed To Date"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        i += 1

       # Grab all billed
       # Could avoid doing this iteration twice
        for j in range(6, revenue_sheet.max_row + 1):    
            j_name = revenue_sheet.cell(row = j, column = NAME_REVENUE_COLUMN).value
            if j_name and job_number in j_name:
                memo_cell = revenue_sheet.cell(row = j, column = MEMO_REVENUE_COLUMN) 
                item_cell = revenue_sheet.cell(row = j, column = ITEM_REVENUE_COLUMN) 
                amount_cell = revenue_sheet.cell(row = j, column = AMOUNT_REVENUE_COLUMN) 

                sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = memo_cell.value
                sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = amount_cell.value
                i += 1

        # write total income for the last time
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total" 
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).value = total_revenue
        sheet.cell(row = i, column = ESTIMATE_COST_COLUMN).font = Font(bold=True)
        i += 1

        # clear out extra rows
        sheet.delete_rows(i, sheet.max_row - i)

        # trim printable area to data?
        sheet._print_area = "A1:I"+str(i)

    # -------------------------------------------------------------------------------- #

    # create and fill all job sheet data
    # skip first 3
    for i in range(3,len(eva_total_wb.sheetnames)):
        sheet = eva_total_wb.worksheets[i]
       # copy initial format into empty sheet
        copySheet(eva_jc_wb.active, sheet)
        createEVAJobCostSheet(sheet)

    eva_total_wb.save(processed_file_path)



def main(argv):
    if len(argv) == 0 or len(argv) > 1:
        print("Error - usage: supply cost detail workbook , revenue workbook, and actual cost workbook")
        return

    eva_total_wb_path = os.path.abspath(argv[0])

    if os.path.isfile(eva_total_wb_path):
        createEVAJobWorkbook(eva_total_wb_path) 
    elif not os.path.isfile(eva_total_wb_path):
        print("Error: eva path does not exist?: ", eva_total_wb_path)
    else:
        print("Error: wrong input")

if __name__ == "__main__":
   main(sys.argv[1:])

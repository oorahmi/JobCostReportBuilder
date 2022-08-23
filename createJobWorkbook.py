'''
1. Start with Job In Progress Cost Detail as of 6-30-2022 file and name the worksheet "Total".

2. Add new worksheets to this file: the Name column shows you the long job name.
 I need a new worksheet added to the file with only the job number as the worksheet's name.
  The job numbers are usually XX-XXXX but occasionally we do have numbers that are XX-XXXX.X.
   I think if you get the strings after the ":" and 
   before the next blank space you should be able to get all the job numbers extracted. 

3. Each worksheet should show the job costs of each job in the format of the sample jc report I attached. 
You can skip the blank columns but I would like to have the rest pretty much identical. 
I don't know how much formatting you can do but keeping the width to fit one page wide and
 the margins at .5 inch will also help as I will be printing out the whole workbook. 
The bottom summary box is what my excel macro adds to the Quickbook's job cost report. 
However, my macro only works on individual reports and with the blank columns and formatting, 
hence this project I came up with for you to help me so I can get it done with one click :)
Following is the calculation details for the summary box:
Total Labor: Add all labor costs except temp labor
Labor OH: Multiply 30% to the total labor calculated above
Other OH: Multiply 0,5% to all costs
Total Cost w/OH: Total Costs + Labor OH + Other OH

The Billed To Date section is something I add manually but
 if you feel up to adding that as well you will need the Actual Revenue Detail csv file which 
 I attached.

3. I'd like to see the check sum at the bottom "Total" sheet. 
Meaning add the total cost from each worksheets you added and show it below the total cost of the
 "Total" sheet. Of course the two numbers will have to be equal otherwise we have an error. :)
  In addition, I'd like the total labor and labor OH shown below the "Total" sheet as well for my info.
----------------------------------------------------------------------------------------------------------------------


Program to take in a .xslx for the total job workbook and Revenue report workbook file and fill out all Job Cost Report Sheets.

Usage:  

    python createJobWorkbook.py path_to_job_workbook_file   path_to_revenue_report_workbook-file

The processed total job workbook will be saved as a copy in /processed

Requires openpyxl

    pip install openpyxl

Author: Brian Wright
8-18-2022

'''

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

def set_border(ws, cell_range):
# https://stackoverflow.com/questions/34520764/apply-border-to-range-of-cells-using-openpyxl
    rows = ws[cell_range]
    side = Side(border_style='thick', color="FF000000")

    rows = list(rows)  # we convert iterator to list for simplicity, but it's not memory efficient solution
    max_y = len(rows) - 1  # index of the last row
    for pos_y, cells in enumerate(rows):
        max_x = len(cells) - 1  # index of the last cell
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )
            if pos_x == 0:
                border.left = side
            if pos_x == max_x:
                border.right = side
            if pos_y == 0:
                border.top = side
            if pos_y == max_y:
                border.bottom = side

            # set new border only if it's one of the edge cells
            if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                cell.border = border

# cant copy sheet from one workbook to another without a deep level copy
# https://stackoverflow.com/questions/42344041/how-to-copy-worksheet-from-one-workbook-to-another-one-using-openpyxl
def copySheet(source_sheet, target_sheet):

    def copy_sheet_attributes(source_sheet, target_sheet):
        target_sheet.sheet_format = copy(source_sheet.sheet_format)
        target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
        target_sheet.merged_cells = copy(source_sheet.merged_cells)
        target_sheet.page_margins = copy(source_sheet.page_margins)
        target_sheet.freeze_panes = copy(source_sheet.freeze_panes)
        # set row dimensions
        # So you cannot copy the row_dimensions attribute. Does not work (because of meta data in the attribute I think). So we copy every row's row_dimensions. That seems to work.
        for rn in range(len(source_sheet.row_dimensions)):
            target_sheet.row_dimensions[rn] = copy(source_sheet.row_dimensions[rn])

        if source_sheet.sheet_format.defaultColWidth is None:
            print('Unable to copy default column wide')
        else:
            target_sheet.sheet_format.defaultColWidth = copy(source_sheet.sheet_format.defaultColWidth)

        # set specific column width and hidden property
        # we cannot copy the entire column_dimensions attribute so we copy selected attributes
        for key, value in source_sheet.column_dimensions.items():
            target_sheet.column_dimensions[key].min = copy(source_sheet.column_dimensions[key].min)   # Excel actually groups multiple columns under 1 key. Use the min max attribute to also group the columns in the targetSheet
            target_sheet.column_dimensions[key].max = copy(source_sheet.column_dimensions[key].max)  # https://stackoverflow.com/questions/36417278/openpyxl-can-not-read-consecutive-hidden-columns discussed the issue. Note that this is also the case for the width, not onl;y the hidden property
            target_sheet.column_dimensions[key].width = copy(source_sheet.column_dimensions[key].width) # set width for every column
            target_sheet.column_dimensions[key].hidden = copy(source_sheet.column_dimensions[key].hidden)

    def copy_cells(source_sheet, target_sheet):
        for (row, col), source_cell in source_sheet._cells.items():
            target_cell = target_sheet.cell(column=col, row=row)
            target_cell._value = source_cell._value
            target_cell.data_type = source_cell.data_type
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.border = copy(source_cell.border)
                target_cell.fill = copy(source_cell.fill)
                target_cell.number_format = copy(source_cell.number_format)
                target_cell.protection = copy(source_cell.protection)
                target_cell.alignment = copy(source_cell.alignment)

            if source_cell.hyperlink:
                target_cell._hyperlink = copy(source_cell.hyperlink)

            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)

    copy_cells(source_sheet, target_sheet)  # copy all the cel values and styles
    copy_sheet_attributes(source_sheet, target_sheet)
        


def createJobWorkbook(total_job_wb_path, revenue_file_path):
    # copy wb and work on copy
    processed_file_path = os.path.split(total_job_wb_path)[0]  +'/processed/' + os.path.basename(total_job_wb_path).split('.')[0] + '_processed.xlsx'
    shutil.copyfile(total_job_wb_path, processed_file_path)

    total_wb = openpyxl.load_workbook(processed_file_path) 
    if not total_wb:
        print("Error: failed to open workbook: ", processed_file_path)
        return

    total_sheet = total_wb.active
    total_sheet.title = "Total" # rename

    revenue_wb = openpyxl.load_workbook(revenue_file_path) 
    if not total_wb:
        print("Error: failed to open workbook: ", revenue_file_path)
        return
    
    revenue_sheet = revenue_wb.active

    #job_str_set = set()
    job_str_set = OrderedDict()
    NAME_COLUMN = 8      # find by name? column H

    # add new sheet for each unique job
    # column
    for i in range(1, total_sheet.max_row + 1): 
        job_data = total_sheet.cell(row = i, column = NAME_COLUMN).value

        # format is currrently:   job_name:job_number type
        # is a job string? 
        if job_data and len(job_data.split(":")) > 1:
            job_number = job_data.split(":")[1].split(' ')[0]    # could be better?
            job_str_set[job_number] = None

    job_numbers = list(job_str_set.keys())
    #job_numbers.sort()
    
    if len(job_numbers) == 0:
        print("Error: failed to find jobs in workbook: ", processed_file_path)
        return


    # add new sheet for each job number
    for job_number in job_numbers:
        total_wb.create_sheet(title=job_number)

    # copy empty job cost sheet
    jc_wb = openpyxl.load_workbook(os.getcwd() + "/data/jc_blank.xlsx") 
    if not jc_wb:
        print("Error: failed to open data workbook: /data/jc_blank.xlsx")
        sys.exit()
        return

    DATE_COLUMN = 6
    ITEM_COLUMN = 10      
    AMOUNT_COLUMN = 16
 
    class JobItem:
        def __init__(self, item_name=""):
            self.item_name = item_name
            self.amount = 0  # used for tracking value for non sub-type

            self.hasSub = False
            self.sub_items = OrderedDict() # contains list of (name, amount) pairs
        
        def processSubItem(self, sub_item_name, amount):
            if not self.hasSub:
                print("error: tried to process subitem on top level item")
                return

            if sub_item_name not in self.sub_items.keys():
                self.sub_items[sub_item_name] = amount 
            else:
                self.sub_items[sub_item_name] += amount
            
    # SCOPED
    def createJobCostSheet(sheet):
        # create job sheet
        job_number = sheet.title

        min_date = datetime.max
        max_date = datetime.min

        job_name = None
        job_items = []
        # get all job progress entries 
        for i in range(1, total_sheet.max_row + 1):    # could optimize by not doing all rows
            j_name = total_sheet.cell(row = i, column = NAME_COLUMN).value

            if j_name and job_number in j_name:

                date = total_sheet.cell(row = i, column = DATE_COLUMN).value
                if date:
                    min_date = min(min_date, date)
                    max_date = max(max_date, date)
                else:
                    print("Warn: Job without a date: ", j_name)

                if not job_name:
                    job_name = j_name 

                j_item = total_sheet.cell(row = i, column = ITEM_COLUMN).value
                j_amount = total_sheet.cell(row = i, column = AMOUNT_COLUMN).value
                if not j_item:
                    print("Warn: have job entry with no item data", j_name)
                    continue

                if not j_amount:
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
                        job_item.processSubItem(sub_item_name, j_amount)
                    else:
                        job_item.amount += j_amount
                    job_items.append(job_item)
                else: # 
                    if job_item.hasSub:
                        job_item.processSubItem(sub_item_name, j_amount)
                    else:
                        job_item.amount += j_amount
                        
        # append job name at top text
        sheet.cell(row = 2, column = 1).value = sheet.cell(row = 2, column = 1).value + " " + job_name

        # write date range
        sheet.cell(row = 3, column = 1).value = "Transactions from: " + min_date.strftime("%m/%d/%y") + " to " + max_date.strftime("%m/%d/%y")

        ITEM_NAME_COLUMN    = 3
        SUBITEM_NAME_COLUMN = 4
        ACT_COST_COLUMN     = 5
        ACT_REVENUE_COLUMN  = 7
        DIFF_COLUMN         = 9

        # date and time
        sheet.cell(row = 1, column = DIFF_COLUMN).value = datetime.today().strftime("%H:%M %p")
        sheet.cell(row = 2, column = DIFF_COLUMN).value = datetime.today().strftime("%B %d, %Y")

        i = 7  # starting point after 'Service' row

        total_labor_cost = 0
        total_cost = 0
        # write job cost data
        for job_item in job_items:
            if not job_item.hasSub:

                if "labor" in job_item.item_name.lower() and "temp" not in job_item.item_name.lower():
                    total_labor_cost += job_item.amount

                total_cost += job_item.amount
                sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = job_item.item_name 
                sheet.cell(row = i, column = ACT_COST_COLUMN).value = job_item.amount
                sheet.cell(row = i, column = ACT_REVENUE_COLUMN).value = 0.0
                sheet.cell(row = i, column = DIFF_COLUMN).value = -job_item.amount
                i += 1
            else:
                sheet.cell(row = i, column = 3).value = job_item.item_name 
                i += 1
                sub_total = 0
                for s_item_name, s_amount in job_item.sub_items.items():

                    if "labor" in s_item_name.lower() and "temp" not in s_item_name.lower():
                        total_labor_cost += s_amount 

                    total_cost += s_amount
                    sub_total += s_amount
                    sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = s_item_name
                    sheet.cell(row = i, column = ACT_COST_COLUMN).value = s_amount
                    sheet.cell(row = i, column = ACT_REVENUE_COLUMN).value = 0.0
                    sheet.cell(row = i, column = DIFF_COLUMN).value = -s_amount
                    i += 1
                # write out total for the subs
                sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Total " + job_item.item_name
                sheet.cell(row = i, column = ACT_COST_COLUMN).value = sub_total
                sheet.cell(row = i, column = ACT_REVENUE_COLUMN).value = 0.0
                sheet.cell(row = i, column = DIFF_COLUMN).value = -sub_total
                i += 1

        NUM_REVENUE_COLUMN = 11
        NAME_REVENUE_COLUMN = 13
        MEMO_REVENUE_COLUMN = 15
        ITEM_REVENUE_COLUMN = 17
        AMOUNT_REVENUE_COLUMN = 19
 
        # get total income
        total_revenue_income = 0
        for j in range(6, revenue_sheet.max_row + 1):    
            j_name = revenue_sheet.cell(row = j, column = NAME_REVENUE_COLUMN).value
            if j_name and job_number in j_name:
                amount_cell = revenue_sheet.cell(row = j, column = AMOUNT_REVENUE_COLUMN) 
                total_revenue_income += amount_cell.value


        # total service
        sheet.cell(row = i, column = 2).value = "Total Service"
        sheet.cell(row = i, column = 2).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_labor_cost
        sheet.cell(row = i, column = ACT_REVENUE_COLUMN).value = 0.0
        sheet.cell(row = i, column = DIFF_COLUMN).value = 0.0
        i += 1

        # total income
        sheet.cell(row = i, column = 2).value = "Total Income"
        sheet.cell(row = i, column = 2).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = 0.0
        sheet.cell(row = i, column = ACT_REVENUE_COLUMN).value = total_revenue_income
        sheet.cell(row = i, column = DIFF_COLUMN).value = 0.0
        i += 1

        # total
        sheet.cell(row = i, column = 1).value = "Total"
        sheet.cell(row = i, column = 1).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_cost
        sheet.cell(row = i, column = ACT_REVENUE_COLUMN).value = total_revenue_income
        sheet.cell(row = i, column = DIFF_COLUMN).value = total_revenue_income - total_cost
        # total font bold
        sheet.cell(row = i, column = ACT_COST_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_REVENUE_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = DIFF_COLUMN).font = Font(bold=True)

        i += 1
        # whitespace
        i += 1
    
        # summary box
        '''
        Calculation details for the summary box:
        Total Labor: Add all labor costs except temp labor
        Labor OH: Multiply 30% to the total labor calculated above
        Other OH: Multiply 0.005 to all costs
        Total Cost w/OH: Total Costs + Labor OH + Other OH
        '''
        labor_oh = total_labor_cost * 0.3
        other_oh = total_cost * 0.005
        total_cost_w_oh = total_cost + labor_oh + other_oh

        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total Labor"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value =  total_labor_cost
        sheet.cell(row = i, column = ACT_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Labor OH"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = labor_oh
        sheet.cell(row = i, column = ACT_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Other OH"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = other_oh
        sheet.cell(row = i, column = ACT_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total Cost w/ OH"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_cost_w_oh
        sheet.cell(row = i, column = ACT_COST_COLUMN).font = Font(bold=True)
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Billed To Date"
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_revenue_income
        sheet.cell(row = i, column = ACT_COST_COLUMN).font = Font(bold=True)
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
                sheet.cell(row = i, column = ACT_COST_COLUMN).value = amount_cell.value
                i+=1

        # write total income for the last time
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total" 
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).font = Font(bold=True)
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_revenue_income

    # -------------------------------------------------------------------------------- #

    # create and fill all job sheet data
    for sheet in total_wb:
        if sheet.title == 'Total':
            continue
        # copy initial format into empty sheet
        copySheet(jc_wb.active, sheet)
        createJobCostSheet(sheet)

    total_wb.save(processed_file_path)


def main(argv):
    if len(argv) == 0 or len(argv) > 2:
        print("Error - usage: supply one job total workbook and one revenue workbook")
        return

    job_wb_path = os.path.abspath(argv[0])
    revenue_wb_path = os.path.abspath(argv[1])

    if os.path.isfile(job_wb_path) and os.path.isfile(revenue_wb_path ):
        createJobWorkbook(job_wb_path, revenue_wb_path)
    elif not os.path.isfile(job_wb_path):
        print("Error: job total workbook does not exist?")
    elif not os.path.isfile(revenue_wb_path):
        print("Error: revenue workbook does not exist?")
    else:
        print("Error: wrong input")

if __name__ == "__main__":
   main(sys.argv[1:])

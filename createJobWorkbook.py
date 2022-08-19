'''
So,I figured we should work on stages for this project.
 The first stage is as mentioned above extracting the job cost reports. 
 I've attached both excel and csv files so you can pick whichever is easier for you to work with. 
 However, it helps me more to have it as an excel file format as I am used to now. 
 I've also attached a sample job cost report for your reference.
  You will find the Job Profitability Report is the summary of the "Detail" files.
   So, Cost Detail files are the expanded detail of the Act. 
   Cost column while the revenue detail file is the expanded detail of the Act. 
   Revenue column.
Here is what I need from your program:

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


Program to take in a .xslx for the total job workbook and fill out all Job Sheets.

Usage:  

    python createJobWorkbook.py path_to_total_job_workbook_file path_to_profitability_report

The processed total job workbook will be saved as a copy in /processed

Requires openpyxl

    pip install openpyxl

Author: Brian Wright
8-18-2022

'''

import sys
import openpyxl
import os 
import shutil
from copy import copy
from datetime import datetime

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
        


def createJobWorkbook(total_job_wb_path):
    # copy wb and work on copy
    processed_file_path = os.path.split(total_job_wb_path)[0]  +'/processed/' + os.path.basename(total_job_wb_path).split('.')[0] + '_processed.xlsx'
    shutil.copyfile(total_job_wb_path, processed_file_path)

    total_wb = openpyxl.load_workbook(processed_file_path) 
    if not total_wb:
        print("Error: failed to open workbook: ", processed_file_path)
        return

    total_sheet = total_wb.active
    total_sheet.title = "Total" # rename

    job_str_set = set()
    NAME_COLUMN = 8      # find by name? column H

    # add new sheet for each unique job
    # column
    for i in range(1, total_sheet.max_row + 1): 
        cell_obj = total_sheet.cell(row = i, column = NAME_COLUMN) 
        job_data = cell_obj.value

        # format is currrently:   job_name:job_number type
        # is a job string? 
        if job_data and len(job_data.split(":")) > 1:
            job_number = job_data.split(":")[1].split(' ')[0]    # could be better?
            job_str_set.add(job_number)

    job_numbers = list(job_str_set)
    job_numbers.sort()
    
    if len(job_numbers) == 0:
        print("Error: failed to find jobs in workbook: ", processed_file_path)
        return


    # add new sheet for each job number
    for job_number in job_numbers:
        total_wb.create_sheet(title=job_number)

    # copy empty job cost sheet
    jc_wb = openpyxl.load_workbook("/data/jc_blank.xlsx") 
    if not jc_wb:
        print("Error: failed to open data workbook: /data/jc_blank.xlsx")
        sys.exit()
        return

    ITEM_COLUMN = 10      
    AMOUNT_COLUMN = 15

    class JobItem:
        def __init__(self, item_name="") -> None:
            self.item_name = item_name

            self.amount = 0  # used for tracking value for non sub-type

            self.hasSub = False
            self.sub_items = [] # contains list of (name, amount) pairs


    def createJobCostSheet(sheet):
        # create job sheet
        job_number = sheet.title

        # date and time
        sheet.cell(row = 1, column = 10).value = datetime.today().strftime("%H:%M %p")
        sheet.cell(row = 2, column = 10).value = datetime.today().strftime("%B %d, %Y")

        job_name = None
        job_items = []
        # get all job progress entries 
        for i in range(1, total_sheet.max_row + 1): 
            name_cell = total_sheet.cell(row = i, column = NAME_COLUMN) 
            if not job_name:
                job_name = name_cell.value

            if job_number in name_cell.value:
                item_cell = total_sheet.cell(row = i, column = ITEM_COLUMN) 
                amount_cell = total_sheet.cell(row = i, column = AMOUNT_COLUMN) 

                item_name = ""
                sub_item_name = None
                # won't have sub item types without :
                if ":" not in item_cell.value:
                    item_name = item_cell.value
                    job_item.item_name = item_name
                elif ":" in item_cell.value:
                    split_item = item_cell.value.split(":")
                    item_name = split_item[0]
                    sub_item_name = split_item[1]
                else:
                    print("Warn: unhandled job item: ", item_cell.value)

                job_item = None
                # find job_item if it already exists
                for j_item in job_items:
                    if j_item.item_name == item_name:
                        job_item = j_item
                if not job_item:
                    job_item = JobItem(item_name)
                    if sub_item_name:
                        job_item.hasSub = True
                        sub_item = None
                        for s_item in job_item.sub_items:
                            if sub_item_name == s_item[0]:
                                sub_item = s_item
                                sub_item[1] += amount_cell.value
                            
                        # create a new sub item
                        if sub_item is None:
                            job_item.sub_items.append((sub_item_name, amount_cell.value))
                    else:
                        job_item.amount += amount_cell.value
                else:
                    job_item


        # append job name
        sheet.cell(row = 2, column = 1).value = sheet.cell(row = 2, column = 1).value + " " + job_name

        ITEM_NAME_COLUMN    = 3
        SUBITEM_NAME_COLUMN = 4
        ACT_COST_COLUMN     = 5
        ACT_REVENUE_COLUMN  = 7
        DIFF_COLUMN         = 9
        i = 7  # starting point after 'Service' row
        # write job cost data
        for job_item in job_items:
            if not job_item.hasSub:
                sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = job_item.item_name 
                sheet.cell(row = i, column = ACT_COST_COLUMN).value = job_item.amount
                sheet.cell(row = i, column = DIFF_COLUMN).value = -job_item.amount
                i += 1
            else:
                sheet.cell(row = i, column = 3).value = job_item.item_name 
                sub_total = 0
                for sub_item in job_item.sub_items:
                    sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = job_item.sub_item[0]
                    sheet.cell(row = i, column = ACT_COST_COLUMN).value = job_item.sub_item[1]
                    sheet.cell(row = i, column = DIFF_COLUMN).value = -job_item.sub_item[1]
                    sub_total += job_item[1]
                    i += 1
                # write out total for the subs
                sheet.cell(row = i, column = ITEM_NAME_COLUMN).value = "Total " + job_item.item_name
                sheet.cell(row = i, column = ACT_COST_COLUMN).value = sub_total
                sheet.cell(row = i, column = DIFF_COLUMN).value = -sub_total
                i += 1

        # handle income?

        # calculate totals
        sheet.cell(row = i, column = 1).value = "Total"
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = "Total Service"
        sheet.cell(row = i, column = ACT_REVENUE_COLUMN).value = "Total Service"
        sheet.cell(row = i, column = DIFF_COLUMN).value = "Total Service"
        i += 1

    
        # summary box
        '''
        Calculation details for the summary box:
        Total Labor: Add all labor costs except temp labor
        Labor OH: Multiply 30% to the total labor calculated above
        Other OH: Multiply 0.5% to all costs
        Total Cost w/OH: Total Costs + Labor OH + Other OH
        '''
        total_labor = 0

        total_cost_w_oh = 0

        for job_item in job_items:
            pass


        labor_oh = total_labor * 0.3
        other_oh = 0.05
        total_cost_w_oh += labor_oh + other_oh

        sheet.cell(row = i, column = SU_BITEM_NAME_COLUMN).value = "Total Service"
        sheet.cell(row = i, column = ACT_COST_COLUMN).value =  total_labor
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Labor OH"
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = labor_oh
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Other OH"
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = other_oh
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Total Cost w/ OH"
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_cost_w_oh
        i += 1
        sheet.cell(row = i, column = SUBITEM_NAME_COLUMN).value = "Billed To Date"
        sheet.cell(row = i, column = ACT_COST_COLUMN).value = total_cost_w_oh


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
        print("Error - usage: supply one job total workbook")
        return

    job_wb_path = argv[0]

    #profitability_path = argv[1]

    if os.path.isfile(job_wb_path):
        createJobWorkbook(job_wb_path)
    else:
        print("Error: wrong input")

    #elif os.path.isdir(path):
        # process multiple workbooks
    #    for file_name in os.listdir(path):
    #        createJobWorkbook(os.path.join(path, file_name))

if __name__ == "__main__":
   main(sys.argv[1:])

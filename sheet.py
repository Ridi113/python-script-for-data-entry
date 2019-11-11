import datetime
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
import uuid

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("GRP-Porject-Creds.json", scope)

client = gspread.authorize(creds)

sheet_1 = client.open("Copy_GRP__hrm__master-data").worksheet('employee_master_info')
sheet_2 = client.open("Copy of GRP__cmn__master-data").worksheet('post')
sheet_3 = client.open("Copy_GRP-HR Upload_14102019_v-1.0.8").worksheet('ALL_Combine')

# data = sheet_1.get_all_records()

data = sheet_1.cell(2, 2).value

# sheet_1.append_row(['fer','rtrret', 'gegeg', '', ''])

# sheet_1.delete_row(105)
num_rows = sheet_1.row_count
print(num_rows)
print(data)

# File to be copied
# wb = openpyxl.load_workbook("GRP - HR Upload_14102019_v-1.0.8(1).xlsx")  # Add file name
# sheet = wb["ALL_Combine"]  # Add Sheet name

sheet_copy = client.open("Copy_GRP-HR Upload_14102019_v-1.0.8").worksheet('ALL_Combine')
num_rows_hr = sheet_copy.row_count
num_col_hr = sheet_copy.col_count

# File to be pasted into
# template = openpyxl.load_workbook("Copy of GRP__hrm__master-data-planning.xlsx")  # Add file name
# temp_sheet = template["employee_master_info"]  # Add Sheet name

sheet_paste_cmn_1 = client.open("Copy of GRP__cmn__master-data").worksheet('office')
num_rows_cmn_1 = sheet_paste_cmn_1.row_count
num_cols_cmn_1 = sheet_paste_cmn_1.col_count

sheet_paste_cmn_2 = client.open("Copy of GRP__cmn__master-data").worksheet('office_unit')
num_rows_cmn_2 = sheet_paste_cmn_2.row_count
num_cols_cmn_2 = sheet_paste_cmn_2.col_count

sheet_paste_cmn_3 = client.open("Copy of GRP__cmn__master-data").worksheet('office_unit_post')
num_rows_cmn_3 = sheet_paste_cmn_3.row_count
num_cols_cmn_3 = sheet_paste_cmn_3.col_count

sheet_paste_cmn_4 = client.open("Copy of GRP__cmn__master-data").worksheet('post')
num_rows_cmn_4 = sheet_paste_cmn_4.row_count
num_cols_cmn_4 = sheet_paste_cmn_4.col_count

sheet_paste_cmn_5 = client.open("Copy of GRP__cmn__master-data").worksheet('office_layer')
num_rows_cmn_5 = sheet_paste_cmn_5.row_count
num_cols_cmn_5 = sheet_paste_cmn_5.col_count

sheet_paste_hrm_1 = client.open("Copy_GRP__hrm__master-data").worksheet('employee_master_info')
num_rows_hrm_1 = sheet_paste_hrm_1.row_count
num_cols_hrm_1 = sheet_paste_hrm_1.col_count

sheet_paste_hrm_2 = client.open("Copy_GRP__hrm__master-data").worksheet('employee_personal_info')
num_rows_hrm_2 = sheet_paste_hrm_2.row_count
num_cols_hrm_2 = sheet_paste_hrm_2.col_count

sheet_paste_hrm_3 = client.open("Copy_GRP__hrm__master-data").worksheet('employee_office')
num_rows_hrm_3 = sheet_paste_hrm_3.row_count
num_cols_hrm_3 = sheet_paste_hrm_3.col_count

print(num_rows_hrm_3)

num_rows_hrm_3 = num_rows_hrm_3 + 666

print(num_rows_hrm_3)


# Copy range of cells as a nested list
# Takes: start cell, end cell, and sheet you want to copy from.
def copyRange(whichCol, whichRow, copyCol, sheet):
    rangeSelected = []
    countRow = 0
    rowCount = sheet.row_count
    # Loops through selected Rows
    for r in range(1, rowCount + 1, 1):
        rowSelected = []
        start_time = datetime.datetime.now()
        if sheet.cell(r, whichCol).value == whichRow:
            rowSelected.append(sheet.cell(r, copyCol).value)
        end_time = datetime.datetime.now()
        if (end_time - start_time).total_seconds() < 1:
            sleep(1.01 - (end_time - start_time).total_seconds())
        rangeSelected.append(rowSelected)
        countRow += 1
        if countRow == 50:
            time.sleep(100)

            countRow = 0
    return rangeSelected
    print('a')


# Paste range
# Paste data from copyRange into template sheet
def pasteRange(whichCol, sheetReceiving, copiedData):
    countRow = 0
    i = sheetReceiving.row_count
    while i > sheetReceiving.row_count:
        start_time = datetime.datetime.now()
        # for j in range(startCol, endCol + 1, 1):
        sheetReceiving.cell(i, whichCol).value = copiedData[countRow][whichCol]
        countRow += 1
        end_time = datetime.datetime.now()
        if (end_time - start_time).total_seconds() < 1:
            sleep(1.01 - (end_time - start_time).total_seconds())
    print('b')


def createData():
    print("Processing...")
    selectedRange = copyRange(1, 'পরিকল্পনা বিভাগ', 5, sheet_copy)  # Change the 4 number values
    pastingRange = pasteRange(2, sheet_paste_hrm_1, selectedRange)  # Change the 4 number values
    # You can save the template as another file to create a new file here too.s
    sheet_paste_hrm_1.save("Copy_GRP__hrm__master-data")

    print("Range copied and pasted!")


copyRange(whichCol=1, whichRow='পরিকল্পনা কমিশন', copyCol=5, sheet=sheet_copy)

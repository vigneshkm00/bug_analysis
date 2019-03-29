import openpyxl
import pandas as pd
import os
bugid_bf=[]
bugid_af=[]
miss_cou=0
def before_process(pid):
    os.chdir("C:\\Users\\vikm\Documents\\testbuganalysis\\before")
    without_cou = 0
    with_cou = 0
    workSheet = openpyxl.load_workbook("SORAExcel_"+pid+".xlsx")
    sheet = workSheet["SORA Bug Details"]
    tot_row = sheet.max_row - 1
    rows = int(sheet.max_row)
    for i in range(2, rows + 1):
        if sheet.cell(row=i, column=3).value == None:
            without_cou = without_cou + 1
        else:
            with_cou = with_cou + 1
            bugid_bf.append(sheet.cell(row=i, column=1).value)
    print("Total bug count Before:", tot_row)
    print("Total no of Blank bugs:", without_cou)
    print("Total no of Bug with comments:", with_cou)

def after_process(pid):
    os.chdir("C:\\Users\\vikm\Documents\\testbuganalysis\\after")
    without_cou = 0
    with_cou = 0
    workSheet = openpyxl.load_workbook("SORAExcel_"+pid+".xlsx")
    sheet = workSheet["SORA Bug Details"]
    tot_row = sheet.max_row - 1
    rows = int(sheet.max_row)
    for i in range(2, rows + 1):
        if sheet.cell(row=i, column=3).value == None:
            without_cou = without_cou + 1
            bugid_af.append(sheet.cell(row=i, column=1).value)
        else:
            with_cou = with_cou + 1
    print("Total bug count Before:", tot_row)
    print("Total no of Blank bugs:", without_cou)
    print("Total no of Bug with comments:", with_cou)

proj_ID=input("Enter the Project ID:")
openpyxl.load_workbook("C:\\Users\\vikm\Documents\\testbuganalysis\\report\\Report.xlsx")
print("BUG analysis for before download:")
before_process(proj_ID)
print("BUG analysis for after download:")
after_process(proj_ID)
print("bug_comment_missing analysis report:")
for i in range(len(bugid_bf)):
    for j in range(len(bugid_af)):
        if bugid_bf[i] == bugid_af[j]:
            miss_cou = miss_cou+1
print("missing count:",miss_cou)
if miss_cou==0: print("No commemts are missing...")

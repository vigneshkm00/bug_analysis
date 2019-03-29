import openpyxl
import pandas as pd
import os
wb = openpyxl.Workbook()
sheet = wb.active
sheet['A1']='Project ID'
sheet['B1']='Before(count)'
sheet['C1']='After(count)'
sheet['D1']='Before_with_comments(count)'
sheet['E1']='After_with_comments(count)'
sheet['F1']='Before_blank(count)'
sheet['G1']='After_blank(count)'
sheet['H1']='Missing(count)'
class excelanalysis:
    def __init__(self,bugID = 0):
        self.pid = bugID

    def before_process(self):
        global bugid_bf
        bugid_bf = []
        os.chdir("C:\\Users\\vikm\Documents\\testbuganalysis\\before")
        global with_cou,without_cou,tot_row
        without_cou = 0
        with_cou = 0
        workSheet = openpyxl.load_workbook("SORAExcel_" + self.pid + ".xlsx")
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

    def after_process(self):
        global bugid_af
        bugid_af = []
        os.chdir("C:\\Users\\vikm\Documents\\testbuganalysis\\after")
        global without_cou1,with_cou1,tot_row1
        without_cou1 = 0
        with_cou1 = 0
        workSheet = openpyxl.load_workbook("SORAExcel_" + self.pid + ".xlsx")
        sheet = workSheet["SORA Bug Details"]
        tot_row1 = sheet.max_row - 1
        rows = int(sheet.max_row)
        for i in range(2, rows + 1):
            if sheet.cell(row=i, column=3).value == None:
                without_cou1 = without_cou1 + 1
                bugid_af.append(sheet.cell(row=i, column=1).value)
            else:
                with_cou1 = with_cou1 + 1
        print("Total bug count Before:", tot_row1)
        print("Total no of Blank bugs:", without_cou1)
        print("Total no of Bug with comments:", with_cou1)
#Initial process of project selection
projectIDs = [] 
print("What you want?\n1.All projects.\n2.selected one.")
flag = int(input())
if  flag == 1:
    create = 1
    print("Okae...Process starts")
    projectIDs = ['726441', '733333', '731437', '727138', '728843', '733563', '728844', '728845', '731438', '726202', '725714']
if  flag == 2:
    create = 1
    NoP=int(input("No of Projects:"))
    for x in range(NoP):
        projectIDs.append(input("Project ID:"))

#loop process of projects
for i in range(len(projectIDs)):
    print("**********%d**********",projectIDs[i])
    miss_cou = 0
    temp_id = projectIDs[i]
    projectIDs[i] = excelanalysis(projectIDs[i])
    print("BUG analysis for before download:")
    projectIDs[i].before_process()
    print("BUG analysis for after download:")
    projectIDs[i].after_process()

    print("bug_comment_missing analysis report:")
    miss_cou = 0
    for i in range(len(bugid_bf)):
        for j in range(len(bugid_af)):
            if bugid_bf[i] == bugid_af[j]:
                miss_cou = miss_cou + 1
    print("missing count:", miss_cou)
    if miss_cou == 0: print("No commemts are missing...")
    valus = (temp_id,tot_row,tot_row1,with_cou,with_cou1,without_cou,without_cou1,miss_cou)
    sheet.append(valus)
    os.chdir("C:\\Users\\vikm\Documents\\testbuganalysis\\report")
    wb.save("Reports.xlsx")
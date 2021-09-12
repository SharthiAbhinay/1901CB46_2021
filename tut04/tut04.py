import os
import csv
from openpyxl import Workbook
from openpyxl import load_workbook
 
def output_by_subject():
    directory = "output_by_subject"
 
    if not os.path.exists(directory): 
        os.makedirs(directory)
    with open(r'tut04\regtable_old.csv','r') as info:
        cinfo=csv.reader(info)
        for data in cinfo:
            del data[4:8]
            del data[2:3]
            if (data[2] =="subno"):continue
            try:
                with open(f"output_by_subject\\{data[2]}.xlsx"): 
                    wb=Workbook()
                    sheet1=wb.active
                    sheet1.append(data)
                    wb.save(f'output_by_subject\\{data[2]}.xlsx')
            except IOError:
                wb=load_workbook(r'output_by_subject\\{}.xlsx'.format(data[2]))
                sheet1=wb.active
                sheet1.append(data)
                wb.save(f'output_by_subject\\{data[2]}.xlsx')
    return
 
def output_individual_roll():
    directory = "output_individual_roll"
 
    if not os.path.exists(directory): 
        os.makedirs(directory)
    with open(r'tut04\regtable_old.csv','r') as info:
        cinfo=csv.reader(info)
        for data in cinfo:
            del data[4:8]
            del data[2:3]
            if (data[0] =="rollno"):continue
            try:
                with open(f"output_individual_roll\\{data[0]}.xlsx"): 
                    wb=Workbook()
                    sheet2=wb.active
                    sheet2.append(data)
                    wb.save(f'output_individual_roll\\{data[0]}.xlsx')
            except IOError:
                wb=load_workbook(r'output_individual_roll\\{}.xlsx'.format(data[0]))
                sheet2=wb.active
                sheet2.append(data)
                wb.save(f'output_individual_roll\\{data[0]}.xlsx')
    return
 
output_by_subject()
output_individual_roll()
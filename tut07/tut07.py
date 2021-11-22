import pandas as pd
import os
import xlsxwriter
from openpyxl import load_workbook
import csv
def feedback_not_submitted():
    col_list1=["rollno","Name","email","aemail","contact"]
    studentinfo_file=pd.read_csv(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut07\studentinfo.csv',usecols=col_list1)
    course_given=pd.read_csv(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut07\course_feedback_submitted_by_students.csv')
    col_list2=["rollno","subno"]
    m_list=["subno","ltp"]
    col_list3=["rollno","register_sem","schedule_sem","subno"]
    course_h=pd.read_csv(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut07\course_registered_by_all_students.csv',usecols=col_list3)
    course_regi=pd.read_csv(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut07\course_registered_by_all_students.csv',usecols=col_list)
    master=pd.read_csv(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut07\course_master_dont_open_in_excel.csv',usecols=m_list)
    d=course_regi.applymap(str).groupby('rollno')['subno'].apply(list).to_dict()
    df={}
    for key,value in d.items():
        for v in value:
            n=df[v]
            if n==0:
                continue
            seriesObj = course_given.apply(lambda x: True if x['stud_roll']==key and x['course_code']==v else False , axis=1)
            numOfRows = len(seriesObj[seriesObj == True].index)
            if n>snumOfRows:
                c1=studentinfo_file.rollno==key
                c2=course_h.rollno==key
                c3=course_h.subno==v
                c4=c2 & c3
                h1=studentinfo_file[c1]
                b=h1.drop("rollno",axis=1)
                a=course_h[c4]
                newdf=pd.concat([a,b],axis=1)
                writer = pd.ExcelWriter(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut07\course_feedback_remaining.xlsx', engine='openpyxl')
                writer.book = load_workbook(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut07\course_feedback_remaining.xlsx')
                writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
                newdf.to_excel(writer,index=False,header=False,startrow=len(reader)+1)
                writer.close()





    ltp_mapping_feedback_type = {1: 'lecture', 2: 'tutorial', 3:'practical'}
    output_file_name = "course_feedback_remaining.xlsx" 
 



feedback_not_submitted()

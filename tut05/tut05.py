import re
import os
import xlsxwriter

try:
  os.mkdir("grades")
except: 
  pass
os.chdir("grades")



grades_file =  open(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut05\grades.csv', 'r')
grades_df = []
for line in grades_file:
  grades_df.append(line[:-1].split(','))

names_file =  open(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut05\names-roll.csv', 'r')
names_df = []
for line in names_file:
  names_df.append(line[:-1].split(','))

subs_file =  open(r'C:\Users\shart\OneDrive\Desktop\Python\1901CB46_2021\tut05\subjects_master.csv', 'r')
subs_df = []
for line in subs_file:
  PATTERN = re.compile(r'''((?:[^,"']|"[^"]*"|'[^']*')+)''')
  subs_df.append(PATTERN.split(line[:-1])[1::2])




grades_dict =  {}
for i in grades_df[0]:
  gdict = {}

for i in grades_df[1:]:
  if i[0] not in grades_dict:
    grades_dict[f"{i[0]}"] = {}

  if i[1] not in grades_dict[f"{i[0]}"]:
      grades_dict[f"{i[0]}"][f"{i[1]}"] = {}

  for j in range(2,len(i)):
    if grades_df[0][j] not in grades_dict[f"{i[0]}"][f"{i[1]}"]:
      grades_dict[f"{i[0]}"][f"{i[1]}"][grades_df[0][j]] = []
    grades_dict[f"{i[0]}"][f"{i[1]}"][grades_df[0][j]].append(i[j])


roll_name = {}
for i in names_df[1:]:
  roll_name[f"{i[0]}"] = i[1]



subcode_details = {}
for i in subs_df[1:]:
  subcode_details[f"{i[0]}"] = {"subname" : i[1], "ltp" : i[2], "crd" : i[2]}



grade_num = {"AA" : 10, "AB" : 9, "BB" : 8, "BC" : 7, "CC" : 6, "CD" : 5, "DD" : 4, "F" : 0, "I" : 0, "F*" : 0, " BB" : 8, "DD*" : 4}





headers = ["Sl No.", "Subject No.", "Subject Name", "L-T-P", "Credit", "Subject Type", "Grade"]

for roll in roll_name.keys():
  stunum = roll
  if roll == "0401ME11": # or roll == "0501CS05":  ## 0401ME11 has a grade "F*" which is invalid, 0501CS05 has a grade " BB" which is invalid(space)
    continue
  print(roll)
  with xlsxwriter.Workbook(f"{roll}.xlsx") as workbook:

    worksheet = workbook.add_worksheet(f"Overall")
    overall = [["Roll No."], ["Name of Student"], ["Discipline"], ["Semester No."], ["Semester wise Credit Taken"], ["SPI"], ["Total Credits Taken"], ["CPI"]]
    overall[0].append(stunum)
    overall[1].append(roll_name[f"{stunum}"])
    overall[2].append(stunum[4:6])

    credsum = 0
    creds = []
    SPI = []
    for j in range(1,len(grades_dict[f"{stunum}"].keys())+1):
      overall[3].append(j)
      creds.append(sum(list(map(int, grades_dict[f"{stunum}"][f"{str(j)}"]["Credit"]))))
      credsum += creds[j-1]
      overall[4].append(creds[j-1])
      SPI.append(round((sum(i[0] * grade_num[i[1]] for i in zip((list(map(int, grades_dict[f"{stunum}"][f"{str(j)}"]["Credit"]))), (grades_dict[f"{stunum}"][f"{str(j)}"]["Grade"])))/creds[j-1]), 2))
      overall[5].append(SPI[j-1])
      overall[6].append(credsum)
      overall[7].append(round((sum(i[0] * i[1] for i in zip(SPI,creds))/credsum), 2))
      
      for row_num, data in enumerate(overall):
        worksheet.write_row(row_num, 0, data)



    for i in range(len(grades_dict[f"{stunum}"].keys())):
      worksheet = workbook.add_worksheet(f"Sem{i+1}")
      new_lst = []
      new_lst.append(headers)
      for j in range(len(grades_dict[f"{stunum}"][f"{str(i+1)}"]["SubCode"])):
        lst = [j+1, grades_dict[f"{stunum}"][f"{str(i+1)}"]["SubCode"][j], subcode_details[grades_dict[f"{stunum}"][f"{str(i+1)}"]["SubCode"][j]]["subname"], subcode_details[grades_dict[f"{stunum}"][f"{str(i+1)}"]["SubCode"][j]]["ltp"], grades_dict[f"{stunum}"][f"{str(i+1)}"]["Credit"][j], grades_dict[f"{stunum}"][f"{str(i+1)}"]["Sub_Type"][j], grades_dict[f"{stunum}"][f"{str(i+1)}"]["Grade"][j]]
        new_lst.append(lst)
      for row_num, data in enumerate(new_lst):
          worksheet.write_row(row_num, 0, data)



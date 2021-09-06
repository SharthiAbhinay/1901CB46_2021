
import os
def write_to_file(data, file_name):
    """Writes list of lists into CSV file"""
    with open(file_name, "w") as f:
        for row in data:
            f.write("\n".join((",".join(row))))
 
 
def output_by_subject():
    roll_dict = {}
    DIRECTORY = "output_by_subject"
 
    with open("regtable_old.csv", "r") as d:
        for row in d:
            row = row.strip().split(",")
            rollno, register_sem, _, subno, _, _, _, _, sub_type = row
            if rollno not in roll_dict:
                roll_dict[rollno] = []
            roll_dict[rollno].append([rollno, register_sem, subno, sub_type])
 
    if not os.path.exists(DIRECTORY):
        os.makedirs(DIRECTORY)
 
    for rollno in roll_dict:
        write_to_file(roll_dict[rollno], os.path.join(DIRECTORY, rollno + ".csv"))
 
 
def output_individual_roll():
    subject_dict = {}
    DIRECTORY = "output_individual_roll"
 
    with open("regtable_old.csv", "r") as f:
        for row in f:
            row = row.strip().split(",")
            rollno, register_sem, _, subno, _, _, _, _, sub_type = row
            if subno not in subject_dict:
                subject_dict[subno] = []
            subject_dict[subno].append([rollno, register_sem, subno, sub_type])
 
    if not os.path.exists(DIRECTORY):
        os.makedirs(DIRECTORY)
 
    for subno in subject_dict:
        write_to_file(subject_dict[subno], os.path.join(DIRECTORY, subno + ".csv"))
 
 
    output_by_subject()
    output_individual_roll()
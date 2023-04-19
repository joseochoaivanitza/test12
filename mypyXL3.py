import csv
with open("fio_salary_kpi.csv", newline = '')as csvfile:
    reader = csv.DictReader(csvfile,delimiter=";")
    for row in reader:
        print(row['full_name'],'|',row['salary'],row['kpi'])

with open("new_file.csv", 'w', newline = '') as csvfile:
    writer = csv.writer(csvfile, delimiter=";")
    writer.writerow(["row 1 el 1", "row 1 el 2", "row 1 el 3"])
    writer.writerow(["row 2 el 1", "row 2 el 2", "row 2 el 3"])
    writer.writerow(["row 3 el 1", "row 3 el 2", "row 3 el 3"])

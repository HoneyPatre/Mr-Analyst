import csv

with open("data/student_data.csv", newline="") as file:
    reader = csv.reader(file)

    # r and c tell us where to grid the labels
    for row in reader:
        if row[0] == "12":
            print(row)
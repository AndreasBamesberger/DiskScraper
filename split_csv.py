from math import ceil

orig_path = "../C-.csv"
lines_per_file = 25000
first_row = None
data_rows = list()

print("reading in existing csv file...")
with open(orig_path, 'r', encoding="utf-16") as orig_file:
    line_count = 0
    for row in orig_file.readlines():
        if line_count == 0:
            first_row = row
        else:
            data_rows.append(row)
        line_count += 1
print("read in {} lines".format(str(line_count)))

file_count = ceil(line_count / lines_per_file)
print("number of new files needed: ")
print(repr(file_count))

orig_path = orig_path.replace('.csv', '')

for i in range(1, file_count+1):
    with open(orig_path + '_part' + str(i) + '.csv', 'w', encoding="utf-16") as sub_file:
        sub_file.writelines(str(first_row))
        for j in range(lines_per_file):
            if j * i <= line_count:
                sub_file.writelines(str(data_rows[j*i]))

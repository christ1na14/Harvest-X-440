import xlrd

part_workbook = xlrd.open_workbook("AggieSource - Participation.xlsx")
worksheet = part_workbook.sheet_by_index(0)
part_dict = {}

for i in range(1, worksheet.nrows):
    row = worksheet.row_values(i)
    part_dict[i] = row[1]
#print(part_dict)

rev_part_workbook = {}
for key, value in part_dict.items():
    rev_part_workbook.setdefault(value, set()).add(key)
rev_part_dict = [key for key, values in rev_part_workbook.items() if len(values) > 1]

duplicates = len(rev_part_dict)
total_part = worksheet.nrows-1
uniq_part = total_part - duplicates

print("There are", total_part, "individuals/households.")
print("There are", uniq_part, "unique individuals.")



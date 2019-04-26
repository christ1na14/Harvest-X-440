import xlrd

# use xlrd to open excel file
part_workbook = xlrd.open_workbook("AggieSource - Participation.xlsx")
# define what worksheet to use
worksheet = part_workbook.sheet_by_index(0)
# define a dictionary
part_dict = {}

# for loop to start at the 2nd row of the excel file and add the index (key) and ID (value) and ad to dictionary
for i in range(1, worksheet.nrows):
    row = worksheet.row_values(i)
    part_dict[i] = row[1]
# print(part_dict)

# dictionary to reverse original dictionary to find duplicates
rev_part_workbook = {}
for key, value in part_dict.items():
    rev_part_workbook.setdefault(value, set()).add(key)
rev_part_dict = [key for key, values in rev_part_workbook.items() if len(values) > 1]

# calculations to determine duplicates
duplicates = len(rev_part_dict)
total_part = worksheet.nrows - 1
uniq_part = total_part - duplicates

# print results
print("There are", total_part, "individuals/households.")
print("There are", uniq_part, "unique individuals.")

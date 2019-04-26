import xlrd

############ DEFINE VARIABLES FOR EXCEL ###################
# use xlrd to open excel file
part_workbook = xlrd.open_workbook("AggieSource - Participation.xlsx")
# define what part_worksheet to use
part_worksheet = part_workbook.sheet_by_index(0)
# use xlrd to open excel file
# app_workbook = xlrd.open_workbook("AggieSource - Applications.xlsx")
app_workbook = xlrd.open_workbook("File.xlsx")

# define what part_worksheet to use
app_worksheet = app_workbook.sheet_by_index(0)

############ FOR LOOPS FOR DICTIONARY CREATION ###################
# for loop to start at the 2nd row of the excel file and add the index (key) and ID (value) and add to dictionary
part_dict = {}
for i in range(1, part_worksheet.nrows):
    row = part_worksheet.row_values(i)
    value_row_int = int(row[1])
    value_row = row[1]
    part_dict[i] = value_row_int

# for loop to start at the 2nd row of the excel file and add the index (key) and ID (value) and add to dictionary
app_dict = {}
for i in range(1, app_worksheet.nrows):
    row = app_worksheet.row_values(i)
    app_dict[int(row[0])] = row[1]

############ FOR LOOPS TO REVERSE DICTIONARIES ###################
# dictionary to reverse original dictionary to find duplicates creates a list
rev_part_workbook = {}
for key, value in part_dict.items():
    rev_part_workbook.setdefault(value, set()).add(key)
rev_part_dict = [key for key, values in rev_part_workbook.items() if len(values) > 1]

rev_app_workbook = {}
for key, value in app_dict.items():
    rev_app_workbook.setdefault(value, set()).add(key)
rev_app_dict = [key for key, values in rev_app_workbook.items() if len(values) > 1]

address_dict = {}
for key in rev_part_workbook:
    if key in app_dict:
        address_dict[key] = app_dict[key]

rev_address_workbook = {}
for key, value in address_dict.items():
    rev_address_workbook.setdefault(value, set()).add(key)
rev_address_dict = [key for key, values in rev_address_workbook.items() if len(values) > 1]


############ CALCULATIONS ###################
# calculations to determine duplicates
duplicates = len(rev_part_dict)
total_part = part_worksheet.nrows - 1
ind_part = len(rev_part_workbook)
ind_hh = len(rev_address_workbook)

# print results
print("There are", ind_hh, "individual households.")
print("There are", total_part, "total participants.")
print("There are", ind_part, "unique individuals.")


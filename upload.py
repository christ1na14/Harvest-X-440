#!C:\xampp\htdocs\venv\Scripts\python.exe
print('Content-Type: text/html\n')
print(
    "<head>"
    "<title>Harvest X</title>"
    "<link rel='shortcut icon' href='favicon.ico'>"
    "<link rel='stylesheet' type='text/css' href='css/normalize.css'/>"
    "<link rel='stylesheet' type='text/css' "
    "href='css/demo.css'/><link rel='stylesheet' type='text/css' href='css/component.css'/>"
    "<link rel='stylesheet' href='https://www.w3schools.com/w3css/4/w3.css'>"
    "<link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' "
    "integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'> "
    "</head>")

import os
import cgi
import cgitb
import xlrd
import shutil
from collections import Counter

cgitb.enable()
form = cgi.FieldStorage()
# gets files from submitted form
fileitem = form['fileToUpload']
fileitem2 = form['fileToUpload2']

# checks to see if there are files there
if fileitem.filename and fileitem2.filename:
    # strip leading path from file name to avoid
    # directory traversal attacks
    fn = os.path.basename(fileitem.filename)
    fn2 = os.path.basename(fileitem2.filename)

    dest_dir = 'uploads/'
    # deletes contents of uploads folder before processing new files
    folder = 'uploads/'
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            # elif os.path.isdir(file_path): shutil.rmtree(file_path)
        except Exception as e:
            print(e)
    # writes file to directory
    open(dest_dir + fn, 'wb').write(fileitem.file.read())
    open(dest_dir + fn2, 'wb').write(fileitem2.file.read())

    # renames files
    os.rename(dest_dir + fn, dest_dir + 'App-File.xlsx')
    os.rename(dest_dir + fn2, dest_dir + 'Part-File.xlsx')

    message = "The files '" + fn + '" "' and '" "' + fn2 + "' were uploaded successfully!"

else:
    message = 'No files were uploaded'

############ DEFINE VARIABLES FOR EXCEL ###################
# use xlrd to open excel file
part_workbook = xlrd.open_workbook('uploads/Part-File.xlsx')
# define what part_worksheet to use
part_worksheet = part_workbook.sheet_by_index(0)
# use xlrd to open excel file
app_workbook = xlrd.open_workbook('uploads/App-File.xlsx')

# define what part_worksheet to use
app_worksheet = app_workbook.sheet_by_index(0)

############ FOR LOOPS FOR DICTIONARY CREATION ###################
# for loop to start at the 2nd row of the excel file and add the index (key) and ID (value) and add to dictionary
part_dict = {}
for i in range(1, part_worksheet.nrows):
    row = part_worksheet.row_values(i)
    # value_row_int = int(row[1])
    value_row = row[5]
    date_row = row[0]
    part_dict[i] = value_row
    # part_dict[i] = date_row


# for loop to start at the 2nd row of the excel file and add the index (key) and ID (value) and add to dictionary
app_dict = {}
for i in range(1, app_worksheet.nrows):
    row = app_worksheet.row_values(i)
    # dictionary index should be your key (whatever row you want, has to be unique), the appending varibale will be
    # your value (whatever row)
    app_dict[row[7]] = row[8]

date_dict = {}
for i in range(1, part_worksheet.nrows):
    row = part_worksheet.row_values(i)
    # value_row_int = int(row[1])
    date_row = row[0]
    new_dates = date_row[5:7]
    date_dict[row[5]] = new_dates

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

rev_date_workbook = {}
for key, value in date_dict.items():
    rev_date_workbook.setdefault(value, set()).add(key)
rev_date_dict = [key for key, values in rev_date_workbook.items() if len(values) > 1]

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
total_part = len(part_dict)
ind_part = len(rev_part_workbook)
ind_hh = len(rev_address_workbook)

# print results
print("<div class='topnav'><a href='/Harvest-X/'>Home</a></div>")
print('<div align=''center>')
print('<br>')
print('<h2>Individual Households:<small class="text-muted">', 90, '</small></h2><br>')
print('<h2>Total Participants:<small class="text-muted">', total_part, '</small></h2><br>')
print('<h2>Unique Individuals:<small class="text-muted">', ind_part, '</small></h2><br>')
print('<h2>Unique Individuals By Month:</h2><br>')
print('<h2>February: <small class="text-muted">', 88, '</small></h2><br>')
print('<h2>March: <small class="text-muted">', 29, '</small></h2><br>')
print('<h2>April: <small class="text-muted">', 37, '</small></h2><br>')
print('<h2>Total Individuals By Month:</h2><br>')
print('<h2>February: <small class="text-muted">', 90, '</small></h2><br>')
print('<h2>March: <small class="text-muted">', 29, '</small></h2><br>')
print('<h2>April: <small class="text-muted">', 37, '</small></h2><br>')
print('<h2>Unique House Holds By Month:</h2><br>')
print('<h2>February: <small class="text-muted">', 27, '</small></h2><br>')
print('<h2>March: <small class="text-muted">', 34, '</small></h2><br>')
print('<h2>April: <small class="text-muted">', 29, '</small></h2><br>')


for key, value in rev_date_workbook.items():
    print(key, len([item for item in value if item]))
print('</div>')

# print("Individual Households:", ind_hh)
# print("Total Participants:", total_part)
# print("Unique Individuals:", ind_part)
# print("###########")
# print("Unique Individuals By Month:")
# for key, value in rev_date_workbook.items():
#     print(key, len([item for item in value if item]))
# #

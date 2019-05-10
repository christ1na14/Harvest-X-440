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
    value_row = row[5]
    date_row = row[0]
    part_dict[i] = value_row

month_dict = {}
for i in range(1, part_worksheet.nrows):
    row = part_worksheet.row_values(i)
    # value_row_int = int(row[1])
    month_value_row = row[5]
    month_date_row = row[0]
    month_new_dates = month_date_row[5:7]
    month_dict[i] = month_value_row, month_new_dates

# for loop to start at the 2nd row of the excel file and add the index (key) and ID (value) and add to dictionary
app_dict = {}
for i in range(1, app_worksheet.nrows):
    row = app_worksheet.row_values(i)
    # dictionary index should be your key (whatever row you want, has to be unique), the appending varibale will be
    # your value (whatever row)
    app_dict[row[6]] = row[7]

date_dict = {}
for i in range(1, part_worksheet.nrows):
    row = part_worksheet.row_values(i)
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

rev_month_workbook = {}
for key, value in month_dict.items():
    rev_month_workbook.setdefault(value, set()).add(key)
rev_month_dict = [key for key, values in rev_month_workbook.items() if len(values) > 1]

address_dict = {}
for key in rev_part_workbook:
    if key in app_dict:
        address_dict[key] = app_dict[key]

rev_address_workbook = {}
for key, value in address_dict.items():
    rev_address_workbook.setdefault(value, set()).add(key)
rev_address_dict = [key for key, values in rev_address_workbook.items() if len(values) > 1]

month_hh = {}
for key in date_dict:
    if key in address_dict:
        month_hh[address_dict[key]] = date_dict[key]

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
print('<h2>Individual Households:<small class="text-muted">', ind_hh, '</small></h2><br>')
print('<h2>Total Participants:<small class="text-muted">', total_part, '</small></h2><br>')
print('<h2>Unique Individuals:<small class="text-muted">', ind_part, '</small></h2><br>')
print('<h2>Unique Households By Month:</h2><br>')
jan_hh_count = 0
for key, value in month_hh.items():
    if value == '01':
        jan_hh_count += 1
if jan_hh_count != 0:
    print('<h2><small class="text-muted">January:', jan_hh_count, '</small></h2><br>')
feb_hh_count = 0
for key, value in month_hh.items():
    if value == '02':
        feb_hh_count += 1
if feb_hh_count != 0:
    print('<h2><small class="text-muted">February', feb_hh_count, "</small></h2><br>")
mar_hh_count = 0
for key, value in month_hh.items():
    if value == '03':
        mar_hh_count += 1
if mar_hh_count != 0:
    print('<h2><small class="text-muted">March', mar_hh_count, "</small></h2><br>")
apr_hh_count = 0
for key, value in month_hh.items():
    if value == '04':
        apr_hh_count += 1
if apr_hh_count != 0:
    print('<h2><small class="text-muted">April', apr_hh_count, "</small></h2><br>")
may_hh_count = 0
for key, value in month_hh.items():
    if value == '05':
        may_hh_count += 1
if may_hh_count != 0:
    print('<h2><small class="text-muted">May', may_hh_count, "</small></h2><br>")
jun_hh_count = 0
for key, value in month_hh.items():
    if value == '06':
        jun_hh_count += 1
if jun_hh_count != 0:
    print('<h2><small class="text-muted">June', jun_hh_count, "</small></h2><br>")
jul_hh_count = 0
for key, value in month_hh.items():
    if value == '07':
        jul_hh_count += 1
if jul_hh_count != 0:
    print('<h2><small class="text-muted">July', jul_hh_count, "</small></h2><br>")
aug_hh_count = 0
for key, value in month_hh.items():
    if value == '08':
        aug_hh_count += 1
if aug_hh_count != 0:
    print('<h2><small class="text-muted">August', aug_hh_count, "</small></h2><br>")
sep_hh_count = 0
for key, value in month_hh.items():
    if value == '09':
        sep_hh_count += 1
if sep_hh_count != 0:
    print('<h2><small class="text-muted">September', sep_hh_count, "</small></h2><br>")
oct_hh_count = 0
for key, value in month_hh.items():
    if value == '10':
        oct_hh_count += 1
if oct_hh_count != 0:
    print('<h2><small class="text-muted">October', oct_hh_count, "</small></h2><br>")
nov_hh_count = 0
for key, value in month_hh.items():
    if value == '11':
        nov_hh_count += 1
if nov_hh_count != 0:
    print('<h2><small class="text-muted">November', nov_hh_count, "</small></h2><br>")
dec_hh_count = 0
for key, value in month_hh.items():
    if value == '12':
        dec_hh_count += 1
if dec_hh_count != 0:
    print('<h2><small class="text-muted">December', dec_hh_count, "</small></h2><br>")

print('<h2>Total Individuals By Month:</h2><br>')
jan_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '01':
        jan_count += len([item for item in value if item])
if jan_count != 0:
    print('<h2><small class="text-muted">January:', jan_count, "</small></h2><br>")

feb_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '02':
        feb_count += len([item for item in value if item])
if feb_count != 0:
    print('<h2><small class="text-muted">February:', feb_count, "</small></h2><br>")

mar_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '03':
        mar_count += len([item for item in value if item])
if mar_count != 0:
    print('<h2><small class="text-muted">March:', mar_count, "</small></h2><br>")

apr_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '04':
        apr_count += len([item for item in value if item])
if apr_count != 0:
    print('<h2><small class="text-muted">April:', apr_count, "</small></h2><br>")

may_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '05':
        may_count += len([item for item in value if item])
if may_count != 0:
    print('<h2><small class="text-muted">May:', may_count, "</small></h2><br>")

jun_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '06':
        jun_count += len([item for item in value if item])
if jun_count != 0:
    print('<h2><small class="text-muted">June:', jun_count, "</small></h2><br>")

jul_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '07':
        jul_count += len([item for item in value if item])
if jul_count != 0:
    print('<h2><small class="text-muted">July:', jul_count, "</small></h2><br>")

aug_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '08':
        aug_count += len([item for item in value if item])
if aug_count != 0:
    print('<h2><small class="text-muted">August:', aug_count, "</small></h2><br>")

sep_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '09':
        sep_count += len([item for item in value if item])
if sep_count != 0:
    print('<h2><small class="text-muted">September:', sep_count, "</small></h2><br>")

oct_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '10':
        oct_count += len([item for item in value if item])
if oct_count != 0:
    print('<h2><small class="text-muted">October:', oct_count, "</small></h2><br>")

nov_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '11':
        nov_count += len([item for item in value if item])
if nov_count != 0:
    print('<h2><small class="text-muted">November:', nov_count, "</small></h2><br>")

dec_count = 0
for key, value in rev_month_workbook.items():
    if key[1] == '01':
        dec_count += len([item for item in value if item])
if dec_count != 0:
    print('<h2><small class="text-muted">December:', dec_count, "</small></h2><br>")

print('<h2>Unique House Holds By Month:</h2><br>')

month = {'01': 'January:',
         '02': 'February:',
         '03': 'March:',
         '04': 'April:',
         '05': 'May:',
         '06': 'June:',
         '07': 'July:',
         '08': 'August:',
         '09': 'September:',
         '10': 'October:',
         '11': 'November:',
         '12': 'December:'}
for key, value in rev_date_workbook.items():
    if key in month:
        print('<h2><small class="text-muted">', month[key], len([item for item in value if item]), "</small></h2><br>")

print('</div>')

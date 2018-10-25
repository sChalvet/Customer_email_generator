# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/05/2018

import sys
import xlrd
from datetime import datetime
import Config
import Get_Domo_Data
from Create_Email import create_email
from My_Library import get_hours_worked, get_lead_email, comment_exist
from Pull_from_web import get_comments_and_team_from_web

# Pull start and end date from arguments
try:
    # START_DATE = datetime.strptime(sys.argv[1], '%m/%d/%Y').date()
    # END_DATE = datetime.strptime(sys.argv[2], '%m/%d/%Y').date()
    START_DATE = datetime.strptime('10/22/2018', '%m/%d/%Y').date()
    END_DATE = datetime.strptime('10/25/2018', '%m/%d/%Y').date()
except IndexError:
    raise Warning("Requires START_DATE and END_DATE as arguments in 01/23/2018 format.")
    exit()

# download csv from Domo API
Get_Domo_Data.get_domo_data()

# Internal Team names
consultant_names_dict = {
    Config.USERNAME_1: Config.USER_EMAIL_1,
    Config.USERNAME_2: Config.USER_EMAIL_2,
    Config.USERNAME_3: Config.USER_EMAIL_3,
    Config.USERNAME_4: Config.USER_EMAIL_4
}

# Pull all the projects from the Projects workbook
try:
    workbook = xlrd.open_workbook(Config.PROJECT_FILE_PATH)
    worksheet = workbook.sheet_by_name(Config.SHEET_NAME)
except FileNotFoundError:
    raise Warning("Make sure the PROJECT HOURS file is in the same directory.\n They should be: "
                  "Services+%7C+CORE%3A+Projects+-+Majime and Services+%7C+CORE%3A+Hours+-+Majime")
    exit()

# Column index and names from PROJECT_FILE_PATH workbook
columns_dict = {}
for col_num in range(0, worksheet.ncols):
    columns_dict[worksheet.cell(0, int(col_num)).value] = col_num

# fetch the data from the rows we want in PROJECT file
project_dictionary = {}
project_info = {}
for row_num in range(1, worksheet.nrows):   # for every row in the sheet
    for col in columns_dict:    # for every column we are interested in
        project_info[col] = worksheet.cell(row_num, columns_dict[col]).value    # get the value of that cell
    if project_info["Project Hub Id"] is not None and project_info["Project Hub Id"] != '' and \
            project_info["Project Status"] == "Active":     # if the project ID is not null and project is active
        project_dictionary[int(project_info["Project Hub Id"])] = project_info.copy()    # add project info to dictionary

# fetch the data from the rows we want in HOURS file
for key in project_dictionary:
    project_dictionary[key].update(
        {"Project Hours Worked":
             get_hours_worked(START_DATE, END_DATE, project_dictionary[key].get("Project Hub Id")),
         "Project Lead Email":
             get_lead_email(project_dictionary[key].get("Project Hub Id")),
         "Comment Exist":
             comment_exist(START_DATE, END_DATE, project_dictionary[key].get("Project Last Comment Date"), workbook.datemode)
         })

# Scrape Comments and customer team info from Hub
project_dictionary = get_comments_and_team_from_web(project_dictionary)

# aggregate all the data into Gmail links
email_mailto_array = create_email(project_dictionary, consultant_names_dict, START_DATE, END_DATE)

print(project_dictionary)

# for thing in email_mailto_array:
#     print(thing)

html_header = "<!DOCTYPE html><html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=iso-8859-1\" /></head><body>"
html_footer = "</body></html>"

# TODO add ability to enter hub credentials
# TODO format date to show real date

# Print results to HTML file
filename = "../Project_Email_" + str(START_DATE) + "_to_" + str(END_DATE) + ".html"
f = open(filename, "w")
f.write(html_header)
for email in email_mailto_array:
    f.write(email)
f.write(html_footer)
f.close()

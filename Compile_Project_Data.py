# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/05/2018

import sys
import xlrd
from datetime import datetime
import Config
from Create_Email import create_email
from My_Library import SHEET_NAME, PROJECT_FILE_PATH, get_hours_worked, get_lead_email, comment_exist
from Pull_from_web import get_comments_and_team_from_web

# Pull start and end date from arguments
try:
    START_DATE = datetime.strptime(sys.argv[1], '%m/%d/%Y').date()  # datetime.strptime('10/01/2018', '%m/%d/%Y').date()
    END_DATE = datetime.strptime(sys.argv[2], '%m/%d/%Y').date()  # datetime.strptime('10/05/2018', '%m/%d/%Y').date()
except IndexError:
    raise Warning("Requires START_DATE and END_DATE as arguments in 01/23/2018 format.")
    exit()

# Internal Team names
consultant_names_dict = {
    Config.USERNAME_1: Config.USER_EMAIL_1,
    Config.USERNAME_2: Config.USER_EMAIL_2,
    Config.USERNAME_3: Config.USER_EMAIL_3,
    Config.USERNAME_4: Config.USER_EMAIL_4
}

# Column numbers from PROJECT_FILE_PATH workbook
columns_dict = {
    "COL_PROJECT_LAST_COMMENT_DATE": 7,
    "COL_PERCENT_HOURS_REMAINING": 25,
    "COL_PROJECT_HOURS_BUDGETED": 27,
    "COL_PROJECT_DOMO_INSTANCE": 33,
    "COL_PROJECT_LEAD_NAME": 39,
    "COL_BILLABLE_HOURS_CONSUMED": 43,
    "COL_PROJECT_COMMENTS": 46,
    "COL_PROJECT_HUB_ID": 56,
    "COL_CUSTOMER_TEAM_COUNT": 60,
    "COL_ACCOUNT_NAME": 74,
    "COL_CSM_EMAIL": 92,
    "COL_ACCOUNT_OWNER_EMAIL": 98,
    "COL_PROJECT_STATUS": 100,
    "COL_ALL_HOURS_CONSUMED": 113,
    "COL_HOURS_REMAINING": 118
}

# Pull all the projects from the Projects workbook
try:
    workbook = xlrd.open_workbook(PROJECT_FILE_PATH)
    worksheet = workbook.sheet_by_name(SHEET_NAME)
except FileNotFoundError:
    raise Warning("Make sure the PROJECT HOURS file is in the same directory.\n They should be: "
                  "Services+%7C+CORE%3A+Projects+-+Majime and Services+%7C+CORE%3A+Hours+-+Majime")
    exit()

# fetch the data from the rows we want in PROJECT file
project_dictionary = {}
project_info = {}
for row_num in range(1, worksheet.nrows):   # for every row in the sheet
    for col in columns_dict:    # for every column we are interested in
        project_info[col] = worksheet.cell(row_num, columns_dict[col]).value    # get the value of that cell
    if project_info["COL_PROJECT_HUB_ID"] is not None and project_info["COL_PROJECT_HUB_ID"] != '' and \
            project_info["COL_PROJECT_STATUS"] == "Active":     # if the project ID is not null and project is active
        project_dictionary[int(project_info["COL_PROJECT_HUB_ID"])] = project_info.copy()    # add project info to dictionary

# fetch the data from the rows we want in HOURS file
for key in project_dictionary:
    project_dictionary[key].update(
        {"COL_PROJECT_HOURS_WORKED":
             get_hours_worked(START_DATE, END_DATE, project_dictionary[key].get("COL_PROJECT_HUB_ID")),
         "COL_PROJECT_LEAD_EMAIL":
             get_lead_email(project_dictionary[key].get("COL_PROJECT_HUB_ID")),
         "COMMENT_EXIST":
             comment_exist(START_DATE, END_DATE, project_dictionary[key].get("COL_PROJECT_LAST_COMMENT_DATE"), workbook.datemode)
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

# Print results to HTML file
filename = "Project_Email_" + str(START_DATE) + "_to_" + str(END_DATE) + ".html"
f = open(filename, "w")
f.write(html_header)
for email in email_mailto_array:
    f.write(email)
f.write(html_footer)
f.close()

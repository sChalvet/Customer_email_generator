import xlrd
from datetime import datetime

from My_Library import get_column_data, SHEET_NAME, PROJECT_FILE_PATH, get_hours_worked, get_number_rows
from Credentials import USER_NAME, PASSWORD

START_DATE = datetime.strptime('05/17/2018', '%m/%d/%Y')
END_DATE = datetime.strptime('07/17/2018', '%m/%d/%Y')

# Column numbers from PROJECT_FILE_PATH workbook
COL_PROJECT_LAST_COMMENT_DATE = 7
COL_PERCENT_HOURS_REMAINING = 25
COL_PROJECT_HOURS_BUDGETED = 27
COL_PROJECT_DOMO_INSTANCE = 33
COL_PROJECT_LEAD_NAME = 39
COL_BILLABLE_HOURS_CONSUMED = 43
COL_PROJECT_COMMENTS = 46
COL_PROJECT_HUB_ID = 56
COL_CUSTOMER_TEAM_COUNT = 60
COL_ACCOUNT_NAME = 74
# COL_CUMULATIVE_HOURS_PARTNERS = 89
COL_CSM_EMAIL = 92
COL_ACCOUNT_OWNER_EMAIL = 98
COL_PROJECT_STATUS = 100
# COL_CUMULATIVE_HOURS_BILLABLE = 104
# COL_PROJECT_HOURS_PURCHASED = 107
COL_ALL_HOURS_CONSUMED = 113
COL_HOURS_REMAINING = 118

# Pull all the projects from the Projects workbook
workbook = xlrd.open_workbook(PROJECT_FILE_PATH)
worksheet = workbook.sheet_by_name(SHEET_NAME)

# Get the number of rows (projects) in the Projects workbook
# (using project_domo_instance because it seems to always be filled)
#project_domo_instance = []
#for cell in worksheet.col(COL_PROJECT_DOMO_INSTANCE):
#    if isinstance(cell.value, str):
 #       project_domo_instance.append(cell.value)

# Save the total row count
## numRows = len(project_domo_instance)
numRows = get_number_rows(PROJECT_FILE_PATH, COL_PROJECT_DOMO_INSTANCE)

# fetch the data from the rows we want in PROJECT file
project_domo_instance = get_column_data(COL_PROJECT_DOMO_INSTANCE, numRows, PROJECT_FILE_PATH)
project_last_comment_date = get_column_data(COL_PROJECT_LAST_COMMENT_DATE, numRows, PROJECT_FILE_PATH)
percent_hours_remaining = get_column_data(COL_PERCENT_HOURS_REMAINING, numRows, PROJECT_FILE_PATH)
project_hours_budgeted = get_column_data(COL_PROJECT_HOURS_BUDGETED, numRows, PROJECT_FILE_PATH)
project_lead_name = get_column_data(COL_PROJECT_LEAD_NAME, numRows, PROJECT_FILE_PATH)
billable_hours_consumed = get_column_data(COL_BILLABLE_HOURS_CONSUMED, numRows, PROJECT_FILE_PATH)
project_comments = get_column_data(COL_PROJECT_COMMENTS, numRows, PROJECT_FILE_PATH)
project_hub_id = get_column_data(COL_PROJECT_HUB_ID, numRows, PROJECT_FILE_PATH)
customer_team_count = get_column_data(COL_CUSTOMER_TEAM_COUNT, numRows, PROJECT_FILE_PATH)
account_name = get_column_data(COL_ACCOUNT_NAME, numRows, PROJECT_FILE_PATH)
# cumulative_hours_partners = get_column_data(COL_CUMULATIVE_HOURS_PARTNERS, numRows) # same as cumulative_hours_billable
csm_email = get_column_data(COL_CSM_EMAIL, numRows, PROJECT_FILE_PATH)
account_owner_email = get_column_data(COL_ACCOUNT_OWNER_EMAIL, numRows, PROJECT_FILE_PATH)  # this is AE
project_status = get_column_data(COL_PROJECT_STATUS, numRows, PROJECT_FILE_PATH)
# cumulative_hours_billable = get_column_data(COL_CUMULATIVE_HOURS_BILLABLE, numRows) # all_hours_consumed seems to be more accurate
# project_hours_purchased = get_column_data(COL_PROJECT_HOURS_PURCHASED, numRows) # same as project_hours_budgeted
all_hours_consumed = get_column_data(COL_ALL_HOURS_CONSUMED, numRows, PROJECT_FILE_PATH) # this is sometimes more than cumulative_hours_billable.. which should be use?
hours_remaining = get_column_data(COL_HOURS_REMAINING, numRows, PROJECT_FILE_PATH)

# fetch the data from the rows we want in HOURS file
time_since_last_entry = []
entry_date = []
hours = []
project_id = []
project_lead_email = []
last_update_date = []

project_hours_worked = []
for project_id in project_hub_id:
    project_hours_worked.append(get_hours_worked(START_DATE, END_DATE, project_id))

print(project_hours_worked[0])

x = 1
# email_from = project_lead_email[x] + "@majime.ws"
email_to = "<customer team emails>"
email_cc = csm_email[x] + ", " + account_owner_email[x] + ", kyle@majime.ws, sam@majime.ws"
email_body = "Dear " + account_name[x] + ",\n"
email_body += "\nThe Majime team is committed to beautiful reliable delivery. A central part of this is good " \
              "communication. As such this is your weekly project status report.\n\n"
email_body += "\n<b>Project Comments:</b>\n" \
              "Hours complelted this week: <list worked hours>\n" \
              "<Latest Comment from Hub> OR <default comment if none>\n\n"
email_body += "\n<b>Project Summary:</b>\n" \
              "<tasks completed>/<tasks remaining>\n" \
              + "<b>" + str(all_hours_consumed[x]) + "</b> hours completed of <b>" + str(project_hours_budgeted[x]) + "</b>. " \
              "Hours remaining: <b>" + str(hours_remaining[x]) + "</b>\n"
email_body += "\n<b>Risks:</b>\n" \
              "<Comment from developer>\n"
email_body += "\n----------------------------------------------------------------------------------------\n\n" \
              "Please let us know if you have any comments or questions.\n\n" \
              "You can also find more information in your project hub: https://" + project_domo_instance[x] + "\n\n" \
              "Sincerely,\n" \
              + project_lead_name[x] + ".\n"

# https://hub.domo.com/project/7385/comments # TODO: get customer names and emails
# https://hub.domo.com/project/7385/team  # TODO: get customer names and emails
# TODO: pull hours for this week for each project
# TODO: filter out completed projects
# TODO: check that project_last_comment_date is within range


print(email_body)
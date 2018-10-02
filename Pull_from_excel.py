import xlrd

from My_Library import get_column_data, FILE_PATH
from Credentials import USER_NAME, PASSWORD

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
COL_ACCOUNT_OWNER_NAME = 98
COL_PROJECT_STATUS = 100
# COL_CUMULATIVE_HOURS_BILLABLE = 104
# COL_PROJECT_HOURS_PURCHASED = 107
COL_ALL_HOURS_CONSUMED = 113
COL_HOURS_REMAINING = 118

workbook = xlrd.open_workbook(FILE_PATH)
worksheet = workbook.sheet_by_name('Sheet1')

# Get the number of rows in the sheet (domo instance seems to always be filled)
project_domo_instance = []
for cell in worksheet.col(COL_PROJECT_DOMO_INSTANCE):
    if isinstance(cell.value, str):
        project_domo_instance.append(cell.value)

# Save the total row count
numRows = len(project_domo_instance)

# fetch the data from the rows we want
percent_hours_remaining = get_column_data(COL_PERCENT_HOURS_REMAINING, numRows)
project_hours_budgeted = get_column_data(COL_PROJECT_HOURS_BUDGETED, numRows)
project_lead_name = get_column_data(COL_PROJECT_LEAD_NAME, numRows)
billable_hours_consumed = get_column_data(COL_BILLABLE_HOURS_CONSUMED, numRows)
project_comments = get_column_data(COL_PROJECT_COMMENTS, numRows)
project_hub_id = get_column_data(COL_PROJECT_HUB_ID, numRows)
customer_team_count = get_column_data(COL_CUSTOMER_TEAM_COUNT, numRows)
account_name = get_column_data(COL_ACCOUNT_NAME, numRows)
# cumulative_hours_partners = get_column_data(COL_CUMULATIVE_HOURS_PARTNERS, numRows) # same as cumulative_hours_billable
csm_email = get_column_data(COL_CSM_EMAIL, numRows)
account_owner_name = get_column_data(COL_ACCOUNT_OWNER_NAME, numRows)
project_status = get_column_data(COL_PROJECT_STATUS, numRows)
# cumulative_hours_billable = get_column_data(COL_CUMULATIVE_HOURS_BILLABLE, numRows) # all_hours_consumed seems to be more accurate
# project_hours_purchased = get_column_data(COL_PROJECT_HOURS_PURCHASED, numRows) # same as project_hours_budgeted
all_hours_consumed = get_column_data(COL_ALL_HOURS_CONSUMED, numRows) # this is sometimes more than cumulative_hours_billable.. which should be use?
hours_remaining = get_column_data(COL_HOURS_REMAINING, numRows)

x = 1

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
              "Sincerely,\n" \
              + project_lead_name[x] + ".\n"

print(email_body)
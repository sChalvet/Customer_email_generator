import xlrd
from datetime import datetime

from My_Library import SHEET_NAME, PROJECT_FILE_PATH, get_hours_worked, get_lead_email, comment_exist
from Pull_from_web import get_comments_and_team_from_web

START_DATE = datetime.strptime('09/04/2018', '%m/%d/%Y').date()
END_DATE = datetime.strptime('09/21/2018', '%m/%d/%Y').date()

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
workbook = xlrd.open_workbook(PROJECT_FILE_PATH)
worksheet = workbook.sheet_by_name(SHEET_NAME)

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

project_dictionary = get_comments_and_team_from_web(project_dictionary)

print(project_dictionary)

#for entry in project_hours_worked:
#    print(entry)

# x = 15
# print(project_hub_id[x])
# email_from = project_lead_email[x]
# email_to = "<customer team emails>"
# email_cc = csm_email[x] + ", " + account_owner_email[x] + ", kyle@majime.ws, sam@majime.ws"
# email_body = "Dear " + account_name[x] + ",\n"
# email_body += "\nThe Majime team is committed to beautiful reliable delivery. A central part of this is good " \
#               "communication. As such this is your weekly project status report.\n\n"
# email_body += "\n<b>Project Comments:</b>\n" \
#               "Hours completed this week: " + str(project_hours_worked[x]) + "\n" \
#               "<Latest Comment from Hub> OR <default comment if none>\n\n"
# email_body += "\n<b>Project Summary:</b>\n" \
#               "<tasks completed>/<tasks remaining>\n" \
#               + "<b>" + str(all_hours_consumed[x]) + "</b> hours completed of <b>" + str(project_hours_budgeted[x]) + "</b>. " \
#               "Hours remaining: <b>" + str(hours_remaining[x]) + "</b>\n"
# email_body += "\n<b>Risks:</b>\n" \
#               "<Comment from developer>\n"
# email_body += "\n----------------------------------------------------------------------------------------\n\n" \
#               "Please let us know if you have any comments or questions.\n\n" \
#               "You can also find more information in your project hub: https://" + project_domo_instance[x] + "\n\n" \
#               "Sincerely,\n" \
#               + project_lead_name[x] + ".\n"

# https://hub.domo.com/project/7385/comments # TODO: get comments
# https://hub.domo.com/project/7385/team  # TODO: get customer names and emails


# print(email_body)
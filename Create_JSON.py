# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/30/2018

import json
import datetime


def create_JSON(project_dictionary, consultant_names_dict, start_date, end_date):

    email_data_JSON = {}
    for key in project_dictionary:
        if int(project_dictionary[key].get("Customer Team Count")) > 0:     # !!! If no customer team then no email

            # Getting nested person info within this project
            customer_team = {}
            for person_id in project_dictionary[key].get("Team Info"):
                if project_dictionary[key]['Team Info'][person_id].get("role") == "Major Domo" or \
                        project_dictionary[key]['Team Info'][person_id].get("role") == "Executive Sponsor":
                    customer_team[project_dictionary[key]['Team Info'][person_id].get("role")] = project_dictionary[key]['Team Info'][person_id].get("email")

            domo_team = {}
            if project_dictionary[key].get("CSM Email") != "(not set)":
                domo_team["CSM Email"] = project_dictionary[key].get("CSM Email")              # CSM email
            else:
                domo_team["CSM Email"] = "(not set)"

            if project_dictionary[key].get("Account Owner Email") != "(not set)":
                domo_team["AE Email"] = project_dictionary[key].get("Account Owner Email")    # AE email
            else:
                domo_team["AE Email"] = "(not set)"

            consultant_team = {}
            for consultant in consultant_names_dict:
                consultant_team[consultant] = consultant_names_dict.get(consultant)

            email_data_JSON[key] = {
                "Customer Team": customer_team,
                "Domo Team": domo_team,
                "Consultant Team": consultant_team,
                "Account Name": project_dictionary[key].get("Account Name").replace(", LLC", "").replace(" Inc.", "").replace(" Inc", "").replace(" LLC", "").replace(" Ltd", ""),
                "Start Date": str(start_date),
                "End Date": str(end_date),
                "Hours Worked": str(project_dictionary[key].get("Project Hours Worked")),
                "Hours Remaining": str(project_dictionary[key].get("Hours Remaining")),
                "Project Total": str(project_dictionary[key].get("Project Hours Budgeted")),
                "Comment": project_dictionary[key].get("Comment"),
                "Domo Instance": project_dictionary[key].get("Project Domo Instance"),
                "Lead Name": project_dictionary[key].get("Project Lead Name"),
                "send_email": True,
                "Data Age": str(datetime.datetime.now().strftime('%m/%d/%Y %I%p %Z')),
                "Email Out Date": str((datetime.datetime.now() + datetime.timedelta(days=1, hours=7)).strftime('%m/%d/%Y %I%p %Z'))  # adding 1 day and 7 hours
            }

    return json.loads(json.dumps(email_data_JSON))

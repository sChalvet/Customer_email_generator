# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/26/2018

import json

def create_email_JSON(project_dictionary, consultant_names_dict, start_date, end_date):

    email_mailto_dict = {}
    for key in project_dictionary:
        if int(project_dictionary[key].get("Customer Team Count")) > 0:     # !!! If no customer team then no email

            # Getting nested person info within this project
            field_to = ""
            for person_id in project_dictionary[key].get("Team Info"):
                if project_dictionary[key]['Team Info'][person_id].get("role") == "Major Domo" or \
                        project_dictionary[key]['Team Info'][person_id].get("role") == "Executive Sponsor":
                    field_to += project_dictionary[key]['Team Info'][person_id].get("email") + ","

            field_cc = ""
            if project_dictionary[key].get("CSM Email") != "(not set)":
                field_cc += project_dictionary[key].get("CSM Email") + ","              # CSM email

            if project_dictionary[key].get("Account Owner Email") != "(not set)":
                field_cc += project_dictionary[key].get("Account Owner Email") + ","    # AE email

            field_bcc = ""
            for consultant in consultant_names_dict:
                field_bcc += consultant_names_dict.get(consultant) + ","

            project_dictionary[key]["Account Name"] = project_dictionary[key].get("Account Name")\
                .replace(", LLC", "")\
                .replace(" Inc.", "")\
                .replace(" Inc", "")\
                .replace(" LLC", "")\
                .replace(" Ltd", "")

            field_subject = project_dictionary[key].get("Account Name") + " Domo Status Update From: " \
                            + str(start_date) + " to " + str(end_date)

            # %0D%0A is equivalent of \n
            field_body1 = "Hi Team,%0D%0A" \
                         "%0D%0ABelow is your weekly project status report:" \
                         "%0D%0A" \
                         "%0D%0AHours completed this week: " + str(project_dictionary[key].get("Project Hours Worked")) \
                         + "%0D%0A" + str(project_dictionary[key].get("Hours Remaining")) + " hours remaining of " \
                         + str(project_dictionary[key].get("Project Hours Budgeted")) + " project total." \
                         + "%0D%0A" \
                         "%0D%0AComment:" \
                         "%0D%0A"
            comment = project_dictionary[key].get("Comment")
            field_body2 = "" \
                         "%0D%0A%0D%0A---%0D%0A" \
                         "%0D%0AWe are committed to your success, so please let us know if you have any comments or " \
                         "questions on this email or as the project unfolds." \
                         "%0D%0A%0D%0A" \
                         "You can also find more information in your customer instance: " \
                         "https://" + project_dictionary[key].get("Project Domo Instance") + "" \
                         "%0D%0A%0D%0A" \
                         "Sincerely,%0D%0A" \
                         + project_dictionary[key].get("Project Lead Name") + ".%0D%0A"


            # Create email dict
            # remove automatically appended comma from last entry using s[:-1]
            email_mailto_dict[key] = {
                "field_to": field_to[:-1],
                "field_cc": field_cc[:-1],
                "field_bcc": field_bcc[:-1],
                "field_subject": field_subject,
                "field_body1": field_body1,
                "field_comment": comment,
                "field_body2": field_body2
            }

    return json.loads(json.dumps(email_mailto_dict))

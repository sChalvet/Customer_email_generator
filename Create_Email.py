# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/05/2018


def create_email(project_dictionary, consultant_names_dict, start_date, end_date):

    email_mailto_array = []
    for key in project_dictionary:

        # Getting nested person info within this project
        field_to = ""
        if int(project_dictionary[key].get("COL_CUSTOMER_TEAM_COUNT")) > 0:
            for person_id in project_dictionary[key].get("COL_TEAM_INFO"):
                if project_dictionary[key]['COL_TEAM_INFO'][person_id].get("role") == "Major Domo" or \
                        project_dictionary[key]['COL_TEAM_INFO'][person_id].get("role") == "Executive Sponsor":
                    field_to += project_dictionary[key]['COL_TEAM_INFO'][person_id].get("email") + ","

        field_cc = ""
        if project_dictionary[key].get("COL_CSM_EMAIL") != "(not set)":
            field_cc += project_dictionary[key].get("COL_CSM_EMAIL") + ","              # CSM email

        if project_dictionary[key].get("COL_ACCOUNT_OWNER_EMAIL") != "(not set)":
            field_cc += project_dictionary[key].get("COL_ACCOUNT_OWNER_EMAIL") + ","    # AE email

        field_bcc = ""
        for consultant in consultant_names_dict:
            if consultant != project_dictionary[key].get("COL_PROJECT_LEAD_NAME"):
                field_bcc += consultant_names_dict.get(consultant) + ","

        field_subject = project_dictionary[key].get("COL_ACCOUNT_NAME") + " Domo Status Update From: " \
                        + str(start_date) + " to " + str(end_date)

        field_body = "Dear " + project_dictionary[key].get("COL_ACCOUNT_NAME") + ",%0D%0A" \
                      "%0D%0AThe Majime team is committed to beautiful reliable delivery. A central part of this is " \
                      "good communication. As such this is your weekly project status report." \
                      "%0D%0A" \
                      "%0D%0AProject Status:" \
                      "%0D%0AHours completed this week: " + str(project_dictionary[key].get("COL_PROJECT_HOURS_WORKED")) \
                      + "%0D%0A" \
                      "%0D%0AComment:" \
                      "%0D%0A" + project_dictionary[key].get("COL_COMMENT") + "" \
                      "%0D%0A%0D%0A" \
                      "%0D%0AProject Summary:" \
                      "%0D%0A" + str(project_dictionary[key].get("COL_ALL_HOURS_CONSUMED")) + " hours completed of " \
                      "" + str(project_dictionary[key].get("COL_PROJECT_HOURS_BUDGETED")) + "" \
                      "%0D%0AHours remaining: " + str(project_dictionary[key].get("COL_HOURS_REMAINING")) + "" \
                      "%0D%0A%0D%0A---%0D%0A" \
                      "%0D%0APlease let us know if you have any comments or questions." \
                      "%0D%0A%0D%0A" \
                      "You can also find more information in your project hub: " \
                      "https://" + project_dictionary[key].get("COL_PROJECT_DOMO_INSTANCE") + "" \
                      "%0D%0A%0D%0A" \
                      "%0D%0ASincerely,%0D%0A" \
                      + project_dictionary[key].get("COL_PROJECT_LEAD_NAME") + ".%0D%0A"

        field_body = field_body.replace(" ", "%20")

        # Create Gmail link
        # remove automatically appended comma from last entry using s[:-1]
        mail_to_link = "<a href=\"https://mail.google.com/mail/?view=cm&fs=1&to=" \
                       + field_to[:-1] + \
                       "&cc=" \
                       + field_cc[:-1] + \
                       "&bcc=" \
                       + field_bcc[:-1] + \
                       "" \
                       "&su=" \
                       + field_subject + \
                       "&body=" \
                       + field_body + \
                       "\">" \
                       + project_dictionary[key].get("COL_ACCOUNT_NAME") + \
                       "</a></br>"

        email_mailto_array.append(mail_to_link)

    return email_mailto_array
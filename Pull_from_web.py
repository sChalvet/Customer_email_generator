# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/05/2018

import time
import Config
import requests
from bs4 import BeautifulSoup
import json


def get_comments_and_team_from_web(project_dictionary):
    url = "https://hub.domo.com/login"
    body = {'email': Config.USERNAME, 'password': Config.PASSWORD}

    s = requests.Session()
    login_page = s.get(url)
    soup = BeautifulSoup(login_page.text, 'lxml')

    body["_token"] = soup.find("input", {"name": "_token"}).get('value')
    r = s.post(soup.form['action'], data=body)

    time.sleep(1)

    for key in project_dictionary:
        # Get project Comment
        if project_dictionary[key].get("Comment Exist"):
            url_comment = "https://hub.domo.com/project/" + str(key) + "/comments/all"

            json_comment = json.loads(s.get(url_comment).content)

            time.sleep(1)

            project_dictionary[key].update({"Comment": json_comment['data'][0]['comment']})
        else:
            project_dictionary[key].update({"Comment": "No Comment Found"})

        # Get Team Emails
        if int(project_dictionary[key].get("Customer Team Count")) > 0:
            url_team = "https://hub.domo.com/project/" + str(key) + "/team/customer"

            team_dict = {}

            json_team = json.loads(s.get(url_team).content)

            time.sleep(1)

            # Putting the Team info into a dictionary
            index = 0
            for team_data in json_team['data']:

                # using re to parse the teams Name, email and role from string
                team_dict[index] = {
                    "name": team_data['first_name'] + " " + team_data['last_name'],
                    "email": team_data['email'],
                    "role": team_data['role']
                }
                index += 1

            project_dictionary[key].update({"Team Info": team_dict})
        else:
            project_dictionary[key].update({"Team Info": "No Team Members Found"})

    return project_dictionary


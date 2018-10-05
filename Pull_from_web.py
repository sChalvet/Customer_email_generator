from selenium import webdriver
import time
from bs4 import BeautifulSoup as soup
import Config
import re

def get_comments_and_team_from_web(project_dictionary):
    driver = webdriver.Chrome()
    url = "https://hub.domo.com/login"
    driver.get(url)

    email = driver.find_element_by_name("email")
    email.send_keys(Config.USERNAME)

    password = driver.find_element_by_name("password")
    password.send_keys(Config.PASSWORD)

    driver.find_element_by_id("login-button").click()

    time.sleep(1)

    for key in project_dictionary:
        # Get project Comment
        if project_dictionary[key].get("COMMENT_EXIST"):
            url_comment = "https://hub.domo.com/project/" + str(key) + "/comments"

            driver.get(url_comment)

            time.sleep(1)

            innerHTML = driver.execute_script("return document.body.innerHTML")
            page_soup = soup(innerHTML, "html.parser")

            containers = page_soup.findAll("td", {"colspan": "2"})

            project_dictionary[key].update({"COL_COMMENT": containers[0].text})
        else:
            project_dictionary[key].update({"COL_COMMENT": "No Comment Found"})

        # Get Team Emails
        if int(project_dictionary[key].get("COL_CUSTOMER_TEAM_COUNT")) > 0:
            url_team = "https://hub.domo.com/project/" + str(key) + "/team"

            team_dict = {}

            driver.get(url_team)

            time.sleep(1)

            innerHTML = driver.execute_script("return document.body.innerHTML")
            page_soup = soup(innerHTML, "html.parser")

            customer_table = page_soup.find("table", {"id": "customer"})
            customer_team = customer_table.findAll("a", {"class": "name editable editable-click"})

            # Putting the Team info into a dictionary
            index = 0
            for team_data in customer_team:
                s = team_data.get('data-value')

                # using re to parse the teams Name, email and role from string
                team_dict[index] = {
                    "name": re.search('\"first_name\": \"(.*)\",', s).group(1) + " " +
                            re.search('\"last_name\": \"(.*)\",', s).group(1),
                    "email": re.search('\"email\": \"(.*)\",', s).group(1),
                    "role": re.search('\"role\": \"(.*)\",', s).group(1)
                }
                index += 1

            project_dictionary[key].update({"COL_TEAM_INFO": team_dict})
        else:
            project_dictionary[key].update({"COL_TEAM_INFO": "No Team Members Found"})

    return project_dictionary


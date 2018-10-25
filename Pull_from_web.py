# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/05/2018

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import Config
import re


def get_comments_and_team_from_web(project_dictionary):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1280x1696')

    driver = webdriver.Chrome(chrome_options=chrome_options, executable_path=r'bin\chromedriver.exe')

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
        if project_dictionary[key].get("Comment Exist"):
            url_comment = "https://hub.domo.com/project/" + str(key) + "/comments"

            driver.get(url_comment)

            time.sleep(1)

            container = driver.find_element_by_xpath("//td[@colspan='2']")

            project_dictionary[key].update({"Comment": container.text})
        else:
            project_dictionary[key].update({"Comment": "No Comment Found"})

        # Get Team Emails
        if int(project_dictionary[key].get("Customer Team Count")) > 0:
            url_team = "https://hub.domo.com/project/" + str(key) + "/team"

            team_dict = {}

            driver.get(url_team)

            time.sleep(1)

            customer_table = driver.find_elements_by_xpath(
                "//table[@id='customer']//a[@class='name editable editable-click']")

            # Putting the Team info into a dictionary
            index = 0
            for team_data in customer_table:
                s = team_data.get_attribute("data-value")

                # using re to parse the teams Name, email and role from string
                team_dict[index] = {
                    "name": re.search('\"first_name\": \"(.*)\",', s).group(1) + " " +
                            re.search('\"last_name\": \"(.*)\",', s).group(1),
                    "email": re.search('\"email\": \"(.*)\",', s).group(1),
                    "role": re.search('\"role\": \"(.*)\",', s).group(1)
                }
                index += 1

            project_dictionary[key].update({"Team Info": team_dict})
        else:
            project_dictionary[key].update({"Team Info": "No Team Members Found"})

    driver.quit()

    return project_dictionary


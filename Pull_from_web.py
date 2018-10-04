from selenium import webdriver
import time
from bs4 import BeautifulSoup as soup
import Credentials

def get_comments_and_team_from_web (project_dictionary):
    driver = webdriver.Chrome()
    url = "https://hub.domo.com/login"
    driver.get(url)

    email = driver.find_element_by_name("email")
    email.send_keys(Credentials.USERNAME)

    password = driver.find_element_by_name("password")
    password.send_keys(Credentials.PASSWORD)

    driver.find_element_by_id("login-button").click()

    time.sleep(1)

# TODO Pull customer emails from their page
    for key in project_dictionary:
        if project_dictionary[key].get("COMMENT_EXIST"):
            url = "https://hub.domo.com/project/" + str(key) + "/comments"

            driver.get(url)

            time.sleep(1)

            innerHTML = driver.execute_script("return document.body.innerHTML")
            page_soup = soup(innerHTML, "html.parser")

            containers = page_soup.findAll("td", {"colspan": "2"})

            project_dictionary[key].update({"COL_COMMENT": containers[0].text})
        else:
            project_dictionary[key].update({"COL_COMMENT": "No Comment Found"})

    return project_dictionary


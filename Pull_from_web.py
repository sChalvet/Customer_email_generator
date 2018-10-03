from selenium import webdriver
from bs4 import BeautifulSoup as soup
import requests
from lxml import html

browser = webdriver.Chrome()
url = "https://hub.domo.com/project/6879/comments"
browser.get(url)

innerHTML = browser.execute_script("return document.body.innerHTML")

# get soup to parse it
page_soup = soup(innerHTML, "html.parser")

email_container = page_soup.find("input", {"name": "email"})
password_container = page_soup.find("input", {"name": "password"})
token_container = page_soup.find("input", {"name": "_token"}).get('value')

payload = {
    "username": "",
    "password": "",
    "_token": token_container
}

session_requests = requests.session()

print(token_container)
# grab all the containers (each product) of a specific class
#containers = page_soup.findAll("table", {"class": "table table-condensed table-striped"})

#comment = container.find("td", {"colspan": "2"}).get('value')
#print(comment)
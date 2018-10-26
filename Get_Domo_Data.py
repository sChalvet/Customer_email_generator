# Created by: Samuel Chalvet
# Email: samuelchalvet@gmail.com
# On: 10/24/2018
import requests
import json
import Config
import pandas as pd


def get_domo_data():
    params = (('grant_type', 'client_credentials'), ('scope', 'data'))
    response = requests.get('https://api.domo.com/oauth/token', params=params, auth=(Config.CLIENT_ID, Config.CLIENT_SECRET))
    json_data = json.loads(response.text)

    api_url = 'https://api.domo.com/v1/datasets/'
    headers = {"Authorization": "Bearer " + json_data['access_token'], "Accept": "text/csv"}

    # Get Hours Data
    hours_api = api_url + Config.HOURS_DATA_KEY + '/data?includeHeader=true'
    data = requests.get(hours_api, headers=headers)
    data.raise_for_status()  # ensure we notice bad responses
    file = open(Config.HOURS_FILE_PATH, "w")
    file.write(data.text.replace("\r", ""))             # !!! File was formatted strangely with an extra Character return..
    file.close()

    # Get Project Data
    project_api = api_url + Config.PROJECT_DATA_KEY + '/data?includeHeader=true'
    data = requests.get(project_api, headers=headers)
    data.raise_for_status()  # ensure we notice bad responses
    file = open(Config.PROJECT_FILE_PATH, "w")
    file.write(data.text)
    file.close()

    pd.read_csv("../project.csv", delimiter=",", encoding="ISO-8859-1").to_excel("../project.xlsx", index=False)
    pd.read_csv("../hours.csv", delimiter=",", encoding="ISO-8859-1").to_excel("../hours.xlsx", index=False)
    # /tmp/project.csv
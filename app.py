import json
from bs4 import BeautifulSoup
import pandas
import os


def readFile():
    file = open("inputs/data.json")

    data = json.load(file)

    fileData = data["aaData"]

    file.close()

    return fileData


def getActions(row):
    html = row["Act"]
    soup = BeautifulSoup(html, "html.parser")

    links = soup.find_all("a")

    listOfActions = list(map(lambda link: link.text.strip(), links))

    return ", ".join(listOfActions)


def getLanguage(row):
    html = row["language"]
    soup = BeautifulSoup(html, "html.parser")

    options = soup.find_all("option")
    language = None
    for option in options:
        if option.has_attr("selected"):
            language = option["value"]
            break

    return language


# def followupDateTime(row):
#     html = row["Followup_Modify_Date_Time"]
#     soup = BeautifulSoup(html, "html.parser")
#     print(soup.prettify())


# Return json data for the row
def parseRow(row):
    actions = getActions(row)

    language = getLanguage(row)

    # followupDateTime = followupDateTime()
    # print("followupDateTime", followupDateTime)

    data = {"Actions": actions, "Language": language}
    return data


def generateXls(json):
    df = pandas.DataFrame(json)

    if not os.path.exists("outputs"):
        os.makedirs("outputs")

    output_file = "outputs/output.xlsx"
    df.to_excel(output_file, index=False, engine="openpyxl")


# Read the file and get rows data
rows = readFile()

jsonData = []

for row in rows:
    data = parseRow(row)
    jsonData.append(data)

generateXls(jsonData)

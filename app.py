import json
from bs4 import BeautifulSoup


def readFile():
    file = open("data.json")

    data = json.load(file)

    fileData = data["aaData"]

    file.close()

    return fileData


def getActions(row):
    html = row["Act"]
    soup = BeautifulSoup(html, "html.parser")

    links = soup.find_all("a")

    return list(map(lambda link: link.text.strip(), links))


def getLanguage(row):
    html = row["language"]
    soup = BeautifulSoup(html, "html.parser")
    # print(soup.prettify())

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


def parseRow(row):
    actions = getActions(row)
    print("actions:", actions)

    language = getLanguage(row)
    print("language:", language)

    # followupDateTime = followupDateTime()
    # print("followupDateTime", followupDateTime)


# Read the file and get rows data
rows = readFile()


parseRow(rows[0])

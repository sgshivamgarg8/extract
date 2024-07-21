import json
from bs4 import BeautifulSoup
import pandas
import os


def readFile():
    try:
        inputFilePath = "inputs/data.json"
        file = open(inputFilePath)
        data = json.load(file)

        fileData = data["aaData"]

        file.close()

        return fileData

    except:
        raise Exception(
            f"File not found, please check if file is present at {inputFilePath}"
        )


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


def getFollowupDateTime(row):
    html = row["Followup_Modify_Date_Time"]
    soup = BeautifulSoup(html, "html.parser")
    return soup.text


def getName(row):
    html = row["name"]
    soup = BeautifulSoup(html, "html.parser")
    return soup.text


def getLeadQuality(row):
    html = row["Lead_Quality"]
    soup = BeautifulSoup(html, "html.parser")
    return soup.find("p").text


# Return json data for the row
def parseRow(row):
    data = {
        "Actions": getActions(row),
        "Language": getLanguage(row),
        "Followup / Modify Date/Time": getFollowupDateTime(row),
        "Name": getName(row),
        "Lead Quality": getLeadQuality(row),
    }

    return data


def generateXls(json):
    try:
        df = pandas.DataFrame(json)

        outputFolderName = "outputs"
        outputFileName = "output"
        if not os.path.exists(outputFolderName):
            os.makedirs(outputFolderName)

        output_file = outputFolderName + "/" + outputFileName + ".xlsx"
        df.to_excel(output_file, index=False, engine="openpyxl")

        print(f"Generated file successfully at {output_file}")

    except:
        raise Exception("Failed to generate output file")


if __name__ == "__main__":
    rows = readFile()

    jsonData = []

    for row in rows:
        data = parseRow(row)
        jsonData.append(data)
        break

    generateXls(jsonData)


# if __name__ == "__main__":
#     try:
#         rows = readFile()

#         jsonData = []

#         for row in rows:
#             data = parseRow(row)
#             jsonData.append(data)

#         generateXls(jsonData)

#     except Exception as e:
#         print(e)

#     finally:
#         # To keep terminal window on screen
#         input("Press enter to exit...")

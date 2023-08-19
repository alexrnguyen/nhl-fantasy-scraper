from selenium import webdriver
from selenium.webdriver.common.by import By
import xlsxwriter


def createWorkbook():
    workbook = xlsxwriter.Workbook("nhl-fantasy-rankings-2023.xlsx")
    return workbook


def getSpreadsheet(workbook, source):
    # Source: https://stackoverflow.com/questions/75771237/error-parsing-cert-retrieved-from-aia-as-der-error-couldnt-read-tbscertifi
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    browser = webdriver.Chrome(options=options)

    if source == "NHL.com":
        url = "https://www.nhl.com/news/nhl-fantasy-hockey-top-250-200-rankings-drafts-players-big-board/c-281505474"
        browser.get(url)
        # Find the highest ranked player
        output = browser.find_element(By.XPATH, "//p[contains(text(), '1. ')]").text
    else:
        url = "https://www.espn.com/fantasy/hockey/story/_/id/37963929/top-250-fantasy-nhl-rankings-scorers-goalies-2023-24"
        browser.get(url)
        # Find the highest ranked player (in each half)
        firstHalf = browser.find_element(By.XPATH, "//p[contains(text(), '1. ')]").text
        secondHalf = browser.find_element(
            By.XPATH, "//p[contains(text(), '151. ')]"
        ).text
        output = "{0}\n{1}".format(firstHalf, secondHalf)

    browser.quit()

    worksheet = workbook.add_worksheet(source)
    colNames = ["Rank", "Name", "Position", "Team"]
    worksheet.write_row(0, 0, colNames)

    # Create an array from the output
    rankings = output.split("\n")

    # Insert player info into spreadsheet
    row = 1
    for player in rankings:
        playerInfo = player.split(", ")
        # Split first element into rank and player name components
        playerInfo[0] = playerInfo[0].replace(".", "", 1)
        playerInfo[0] = playerInfo[0].split(" ", 1)
        playerInfo.insert(1, playerInfo[0][1])  # Add player name as its own element
        playerInfo[0] = playerInfo[0][0]  # Ranking
        playerInfo[0] = int(playerInfo[0])  # Remove period after ranking

        # Remove excess information from data (eg. health status, position ranking)
        playerInfo[-1] = playerInfo[-1].split(" ")
        playerInfo[-1] = playerInfo[-1][0]
        if source == "ESPN":
            playerInfo[-1] = playerInfo[-1].upper()  # Uppercase team abbreviations

        worksheet.write_row(row, 0, playerInfo)
        row += 1


if __name__ == "__main__":
    workbook = createWorkbook()
    getSpreadsheet(workbook, "NHL.com")
    getSpreadsheet(workbook, "ESPN")
    workbook.close()

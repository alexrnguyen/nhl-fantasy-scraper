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
        # Remove commas
        for char in player:
            if char == ",":
                player = player.replace(char, "")

        # Split into rank, first/last name, position, and team
        if source == "NHL.com":
            playerInfo = player.split(" ")[:5]
        else:
            playerInfo = player.split(" ")[:-1]
        playerInfo[0] = int(playerInfo[0][:-1])  # Remove period after ranking

        if len(playerInfo) == 5:
            # Combine first and last name into 1 element
            playerInfo[1] = "{0} {1}".format(playerInfo[1], playerInfo[2])
            playerInfo.pop(2)
        else:
            # Edge Case (player has multiple first or last names -> eg. Trevor van Riemsdyk)
            playerInfo[1] = "{0} {1} {2}".format(
                playerInfo[1], playerInfo[2], playerInfo[3]
            )
            playerInfo.pop(2)
            playerInfo.pop(2)

        if source == "ESPN":
            playerInfo[-1] = playerInfo[-1].upper()  # Uppercase team abbreviations

        worksheet.write_row(row, 0, playerInfo)
        row += 1


if __name__ == "__main__":
    workbook = createWorkbook()
    getSpreadsheet(workbook, "NHL.com")
    getSpreadsheet(workbook, "ESPN")
    workbook.close()

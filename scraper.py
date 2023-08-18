from selenium import webdriver
from selenium.webdriver.common.by import By
import xlsxwriter

punctuation = ".,"


def createWorkbook():
    workbook = xlsxwriter.Workbook("nhl-fantasy-rankings-2023.xlsx")
    return workbook


def getNHLPageSpreadsheet(workbook):
    # Source: https://stackoverflow.com/questions/75771237/error-parsing-cert-retrieved-from-aia-as-der-error-couldnt-read-tbscertifi
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # Source: https://www.selenium.dev/selenium/docs/api/py/index.html
    nhl_page_url = "https://www.nhl.com/news/nhl-fantasy-hockey-top-250-200-rankings-drafts-players-big-board/c-281505474"
    browser = webdriver.Chrome(options=options)
    browser.get(nhl_page_url)

    # Find the highest ranked player
    output = browser.find_element(By.XPATH, "//p[contains(text(), '1. ')]").text
    browser.quit()

    worksheet = workbook.add_worksheet("NHL.com")

    colNames = ["Rank", "Name", "Position", "Team"]
    worksheet.write_row(0, 0, colNames)
    # Create an array from the output
    rankings = output.split("\n")

    # Insert player info into spreadsheet
    row = 1
    for player in rankings:
        for char in player:
            if char in punctuation:
                player = player.replace(char, "")
        playerInfo = player.split(" ")[:5]
        playerInfo[1] = "{0} {1}".format(playerInfo[1], playerInfo[2])
        playerInfo.pop(2)
        worksheet.write_row(row, 0, playerInfo)
        row += 1


if __name__ == "__main__":
    workbook = createWorkbook()
    getNHLPageSpreadsheet(workbook)
    workbook.close()

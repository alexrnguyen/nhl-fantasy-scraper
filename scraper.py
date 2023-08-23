from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
import xlrd


# Create a new Excel spreadsheet
def createWorkbook():
    workbook = xlsxwriter.Workbook("nhl-fantasy-rankings-2023.xlsx")
    return workbook


# Create a sheet within an Excel spreadsheet containing the fantasy rankings from a given source
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
    elif source == "Yahoo":
        url = "https://hockey.fantasysports.yahoo.com/hockey/draftanalysis"
        output = ""
        browser.get(url)

        # Get Top 300 players (30 per page)
        for i in range(10):
            # Wait for table to load up
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//div[@class='H(30px) W(100%) Bgc($shade) D(f) Ai(c) Jc(e)']",
                    )
                )
            )
            # Find all players on the current page
            elements = browser.find_elements(By.XPATH, "//div[@data-tst='player-name']")
            for element in elements:
                output += element.text + "\n"

            # Go to next page
            buttons = browser.find_elements(By.XPATH, "//button[@role='presentation']")
            nextPageButton = buttons[1]
            nextPageButton.click()
        output = output.strip()  # Remove new line character after last player
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
    if source == "Yahoo":
        colNames = ["Rank", "Name"]
    else:
        colNames = ["Rank", "Name", "Position", "Team"]
    worksheet.write_row(0, 0, colNames)

    # Create an array from the output
    rankings = output.split("\n")

    # Insert player info into spreadsheet
    row = 1
    for player in rankings:
        # Player information on Yahoo's website is formatted quite differently. Handle Yahoo sheet in else clause
        if source != "Yahoo":
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
            worksheet.autofilter("A1:D1")

        else:
            playerInfo = [row, player]
            worksheet.write_row(row, 0, playerInfo)
            worksheet.autofilter("A1:B1")
        row += 1


def getAverageRankings(workbook):
    rankings = {}
    readOnlyWorkbook = xlrd.open_workbook("nhl-fantasy-rankings-2023.xlsx")

    numSheets = readOnlyWorkbook.nsheets
    for i in range(numSheets):
        worksheet = readOnlyWorkbook.sheet_by_index(i)
        # Do not read contents of average rankings
        if readOnlyWorkbook.sheet_names()[i] == "Average Rankings":
            continue
        for row in range(1, worksheet.nrows):
            if rankings.get(worksheet.cell_value(row, 1)) is None:
                # Create an array associated with the player
                rankings[worksheet.cell_value(row, 1)] = [worksheet.cell_value(row, 0)]
            else:
                # Add the ranking to the existing array for the player
                rankings[worksheet.cell_value(row, 1)].append(
                    worksheet.cell_value(row, 0)
                )

    # Create a dictionary containing each player's average ranking
    averageRankings = {}
    for player in rankings.keys():
        rankingArray = rankings.get(player)
        averageRankings[player] = sum(rankingArray) / len(rankingArray)

    # Add average rankings to a new worksheet
    avgRankingsWorksheet = workbook.add_worksheet("Average Rankings")
    colNames = ["Name", "Average Rank"]
    avgRankingsWorksheet.write_row(0, 0, colNames)
    row = 1
    for player in averageRankings.keys():
        playerInfo = [player, averageRankings.get(player)]
        avgRankingsWorksheet.write_row(row, 0, playerInfo)
        row += 1

    avgRankingsWorksheet.autofilter("A1:B1")


if __name__ == "__main__":
    workbook = createWorkbook()
    getSpreadsheet(workbook, "NHL.com")
    getSpreadsheet(workbook, "ESPN")
    getSpreadsheet(workbook, "Yahoo")
    getAverageRankings(workbook)
    workbook.close()

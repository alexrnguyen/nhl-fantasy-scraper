from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
import xlrd


# Create a new Excel spreadsheet
def create_workbook():
    workbook = xlsxwriter.Workbook("nhl-fantasy-rankings-2023.xlsx")
    return workbook


# Create a sheet within an Excel spreadsheet containing the fantasy rankings from a given source
def get_spreadsheet(workbook, source):
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
            next_page_button = buttons[1]
            next_page_button.click()
        output = output.strip()  # Remove new line character after last player
    else:
        url = "https://www.espn.com/fantasy/hockey/story/_/id/37963929/top-250-fantasy-nhl-rankings-scorers-goalies-2023-24"
        browser.get(url)
        # Find the highest ranked player (in each half)
        first_half = browser.find_element(By.XPATH, "//p[contains(text(), '1. ')]").text
        second_half = browser.find_element(
            By.XPATH, "//p[contains(text(), '151. ')]"
        ).text
        output = "{0}\n{1}".format(first_half, second_half)

    browser.quit()

    worksheet = workbook.add_worksheet(source)
    if source == "Yahoo":
        col_names = ["Rank", "Name"]
    else:
        col_names = ["Rank", "Name", "Position", "Team"]
    worksheet.write_row(0, 0, col_names)

    # Create an array from the output
    rankings = output.split("\n")

    # Insert player info into spreadsheet
    row = 1
    for player in rankings:
        # Player information on Yahoo's website is formatted quite differently. Handle Yahoo sheet in else clause
        if source != "Yahoo":
            player_info = player.split(", ")
            # Split first element into rank and player name components
            player_info[0] = player_info[0].replace(".", "", 1)
            player_info[0] = player_info[0].split(" ", 1)
            player_info.insert(
                1, player_info[0][1]
            )  # Add player name as its own element
            player_info[0] = player_info[0][0]  # Ranking
            player_info[0] = int(player_info[0])  # Remove period after ranking

            # Remove excess information from data (eg. health status, position ranking)
            player_info[-1] = player_info[-1].split(" ")
            player_info[-1] = player_info[-1][0]
            if source == "ESPN":
                player_info[-1] = player_info[
                    -1
                ].upper()  # Uppercase team abbreviations

            worksheet.write_row(row, 0, player_info)

        else:
            player_info = [row, player]
            worksheet.write_row(row, 0, player_info)
        row += 1

    if source != "Yahoo":
        worksheet.autofilter("A1:D301")
    else:
        worksheet.autofilter("A1:B301")
    worksheet.autofit()


def get_average_rankings(workbook):
    rankings = {}
    read_only_workbook = xlrd.open_workbook("nhl-fantasy-rankings-2023.xlsx")

    num_sheets = read_only_workbook.nsheets
    for i in range(num_sheets):
        worksheet = read_only_workbook.sheet_by_index(i)
        # Do not read contents of average rankings
        if read_only_workbook.sheet_names()[i] == "Average Rankings":
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
    average_rankings = {}
    for player in rankings.keys():
        ranking_array = rankings.get(player)
        average_rankings[player] = sum(ranking_array) / len(ranking_array)

    # Add average rankings to a new worksheet
    avg_rankings_worksheet = workbook.add_worksheet("Average Rankings")
    col_names = ["Name", "Average Rank"]
    avg_rankings_worksheet.write_row(0, 0, col_names)
    row = 1
    for player in average_rankings.keys():
        player_info = [player, average_rankings.get(player)]
        avg_rankings_worksheet.write_row(row, 0, player_info)
        row += 1

    avg_rankings_worksheet.autofilter("A1:B501")
    avg_rankings_worksheet.autofit()


if __name__ == "__main__":
    workbook = create_workbook()
    get_spreadsheet(workbook, "NHL.com")
    get_spreadsheet(workbook, "ESPN")
    get_spreadsheet(workbook, "Yahoo")
    get_average_rankings(workbook)
    workbook.close()

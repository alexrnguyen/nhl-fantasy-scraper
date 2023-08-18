from selenium import webdriver
from selenium.webdriver.common.by import By

punctuation = ".,"


def getNHLPageSpreadsheet():
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

    # Create an array from the output
    rankings = output.split("\n")

    # Prepare player information for spreadsheet
    for player in rankings:
        for char in player:
            if char in punctuation:
                player = player.replace(char, "")
        playerInfo = player.split(" ")[:4]
        print(playerInfo)


getNHLPageSpreadsheet()

# nhl-fantasy-scraper
A data scraper that extracts fantasy hockey rankings from NHL.com and ESPN into a spreadsheet. Using the extracted rankings, the program calculates each player's average ranking and writes a sheet containing each player's average ranking. Built using Python, Selenium, and XlsxWriter.

## Sample Spreadsheet

An example of the spreadsheet generated by the data scraper is included as an Excel file named nhl-fantasy-rankings-2023.xlsx. Alternatively, a link to a Google Sheets version of the same spreadsheet can be accessed [here](https://docs.google.com/spreadsheets/d/1Q6Y7awE1mJHTnA3Y-zH3Ev7yHEvHNYV0524eykKdXIU/edit?usp=sharing)

## How to use

1. Clone this repository to create a local repository on your computer

```sh
git clone https://github.com/alexrnguyen/nhl-fantasy-scraper.git
```

2. Install all libraries required to run the data scraper

```sh
pip install -r requirements.txt
```

3. Run the data scraper program to retrieve the latest fantasy hockey rankings from NHL.com and ESPN

```sh
python scraper.py
```

4. The latest fantasy rankings will be available through a spreadsheet named nhl-fantasy-rankings-2024. This file is located in the same directory as the data scraper program.

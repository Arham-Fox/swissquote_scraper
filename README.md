# Swissquote Data Scraper

This project is an automated Python-based scraper for collecting and exporting financial data from [Swissquote](https://www.swissquote.ch/trading-platform/#scanner). It uses Selenium and Pandas to navigate the website, interact with elements, collect fundamental data on securities, and save the data to an Excel file.

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Customization](#customization)
- [Notes](#notes)
- [Disclaimer](#disclaimer)
- [License](#license)

## Overview
The script:

1. Navigates to Swissquote's stock scanner page.
2. Sets the language to German.
3. Adds a new scanner and configures it to show the "Höchste Kapitalisierung der Schweiz" filter.
4. Scrolls down to load the entire table of securities.
5. Extracts table data and collects links for further scraping.
6. Visits each link to scrape fundamental data for each security.
7. Merges, cleans, and exports the data as an Excel file.

## Features
- **Automated Navigation and Data Collection**: Uses Selenium to automate web scraping tasks on Swissquote.
- **Dynamic Data Collection**: Supports scraping both table data and specific data fields from detail pages.
- **Excel Export**: Exports the final dataset to an Excel file with customized column widths for readability.
- **Error Handling**: Robust error handling to ensure graceful failures if elements are not found.

## Requirements
- **Python 3.6+**
- **Google Chrome** and **Chromedriver** compatible with the Chrome version installed
- **Undetected ChromeDriver**: For bypassing bot detection
- **Selenium**: For web scraping
- **Pandas**: For data manipulation and storage
- **XlsxWriter**: For exporting data to Excel

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/Arham-Fox/swissquote_scraper.git
   ```
2. Install the required packages:
   ```bash
   pip install bs4 pandas selenium undetected-chromedriver xlsxwriter
   ```
3. Download Chromedriver and place it in /usr/bin/chromedriver.

## Usage
1. Run the script:

   ```bash
   python3 swissquote_scraper.py
   ```
2. The script will navigate to the Swissquote scanner page, set up filters, load data, and export the resulting data as an Excel file named in the format: YYYYMMDD-HHMMSS-sq-ch-aktien.xlsx.

## Customization
You can customize the script by adjusting the following parameters:

- **URL**: Set the initial URL of the page to be scraped (initial_url).
- **Scrolling and Wait Time**: Adjust the number of page downs and wait time (press_page_down_n_times).
- **Output Filename**: Modify the output filename in the save_dataframe_to_xlsx function as per requirements.

## Notes
- Ensure that Chromedriver matches the installed Chrome version. You may need to update Chromedriver periodically.
- This project uses undetected_chromedriver to bypass bot detection on Swissquote. Running in headless mode may still trigger detection.

## Disclaimer
- Some portions of this README and project guidance were generated with the assistance of ChatGPT by OpenAI.

## License
- This project is licensed under the MIT License.

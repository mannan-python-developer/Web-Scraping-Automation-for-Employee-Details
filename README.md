# Automation Employee Details Scraper

This Python script is designed for web scraping employee details from the Madhya Pradesh Education Portal's employee search page. It utilizes Selenium for web automation, BeautifulSoup for HTML parsing, and OpenPyXL for Excel manipulation.

## Features

- **Automated Data Extraction:** The script automates the process of searching for employee details based on combinations provided in an Excel file.
- **Dynamic Pagination Handling:** It dynamically navigates through paginated search results, ensuring comprehensive data extraction.
- **Data Parsing and Storage:** Utilizes BeautifulSoup to parse HTML content and organizes the extracted employee data into an Excel file.

## Usage

1. **Excel Input:** The script reads combinations of first names and last names from an input Excel file (`combinations.xlsx`).
2. **Web Scraping:** It then performs automated searches on the Madhya Pradesh Education Portal using Selenium, based on the provided combinations.
3. **Data Extraction:** Extracted employee details, including unique ID, name, date of birth, category, designation, subject, posting details, district, contact information, date of joining, and payment authority, are saved in an Excel file (`Employee_details.xlsx`).

## Requirements

- Python 3.x
- Selenium
- BeautifulSoup
- OpenPyXL
- ChromeDriver (for WebDriver)

## Setup

1. Install required dependencies:

   ```bash
   pip install selenium beautifulsoup4 openpyxl

2. Download ChromeDriver and provide the path in the script:
   (webdriver.Chrome('path/to/chromedriver')).

## License
This project is licensed under the MIT License.




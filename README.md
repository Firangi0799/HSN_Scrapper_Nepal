# HSN Code Scraper

This Python script automates the process of searching for Harmonized System Nomenclature (HSN) codes and related product information on the [nnsw.gov.np](https://nnsw.gov.np) website. It reads a list of materials from an Excel file, queries the site for each material, extracts relevant search results, and saves the combined data into a new Excel file.

---

## Features

- Uses Selenium WebDriver with Chrome in headless mode to perform automated searches.
- Randomizes user agents and waits between requests to mimic human behavior and avoid detection.
- Parses search results to extract HSN codes, product descriptions, probabilities, sub-parts, and product names.
- Outputs all gathered data into an Excel file for further analysis.

---

## Requirements

- Python 3.7+
- [Selenium](https://pypi.org/project/selenium/)
- [Pandas](https://pandas.pydata.org/)
- Google Chrome browser installed
- ChromeDriver matching your Chrome version (configured in the script)

---

## Setup

1. **Install Python libraries:**

```pip install selenium pandas```

2. **ChromeDriver:**

- Download ChromeDriver that matches your installed Chrome browser version from [here](https://chromedriver.chromium.org/downloads).
- Update the `service = Service()` path in the script to point to your ChromeDriver executable location.

3. **Input Data:**

- Prepare an Excel file containing a column named `materials` with the list of materials you want to search.
- Update the `materials_df = pd.read_excel(...)` path in the script to point to your Excel file.

4. **Output file:**

- The script saves the final results to a specified Excel file path (`output_path` variable). Adjust this path as needed.

---

## Usage

Run the script using Python:

```python main.py```

The script will:

- Load search terms from the input Excel.
- Perform searches for each term on the target website.
- Extract and parse results.
- Save all collected results into the output Excel file.
- Print progress and status messages to the console.

---

## Code Overview

- **User-Agent Rotation:** Randomizes user-agent strings to reduce detection risk.
- **Headless Chrome:** Chrome runs without a GUI for efficient background automation.
- **Anti-bot measures:** Modifies browser properties to avoid the "webdriver" flag detection.
- **Robust Parsing:** Parses multiple attributes and handles cases with missing or extra data.
- **Timeout Handling:** Continues gracefully if pages fail to load or no results are found.

---

## Notes

- Make sure ChromeDriver version matches your installed Chrome browser.
- Adjust paths for input/output files according to your environment.
- The random delays between requests are intentional to mimic human browsing.
- This script is intended for ethical and legal scraping purposes only.

---

## License

This project is released under the MIT License.

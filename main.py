import re
import time
import random
import pandas as pd
from urllib.parse import urlparse, parse_qs, urlencode
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def is_probability(s):
    try:
        float(s)
        return True
    except:
        return False

# List of user agents to rotate randomly (add more if you want)
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)"
    " Chrome/115.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko)"
    " Version/16.1 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko)"
    " Chrome/114.0.0.0 Safari/537.36",
]

chrome_options = Options()
chrome_options.headless = True
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument(f"user-agent={random.choice(USER_AGENTS)}")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)

service = Service(r"D:\WebScrapping\chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome(service=service, options=chrome_options)

# Prevent detection
driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
    'source': '''
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined})
    '''
})

materials_df = pd.read_excel(r"D:\HSN_Finder\materials.xlsx")  # adjust path
search_terms = materials_df['materials'].dropna().tolist()  # first 5 items

base_url = "https://nnsw.gov.np/hscode-search/search"
all_records = []

try:
    for term in search_terms:
        print(f"Searching for: {term}")
        query_params = {'q': term}
        full_url = f"{base_url}?{urlencode(query_params)}"
        driver.get(full_url)

        # Randomized wait between 4-7 seconds to mimic human behavior
        time.sleep(random.uniform(4,7))

        wait = WebDriverWait(driver, 20)
        parent_selector = "#root > div > div > div.padding-bottom > div > div.container-fluid.Content-module_marginTop__wDS2U > div.App-module_searhEngineApp__dkHcP > div > div.Search-module_content__qJ8TK > div"
        try:
            parent = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, parent_selector)))
            raw_text = parent.text.strip()
            lines = raw_text.split('\n')

            data_lines = lines[3:]
            current_record = {}
            hs_code_pattern = re.compile(r'^\d{6}$')

            for line in data_lines:
                line = line.strip()
                if hs_code_pattern.match(line):
                    if current_record:
                        current_record['Search Query'] = term
                        all_records.append(current_record)
                        current_record = {}

                    current_record['HSN'] = line
                    current_record['All Attributes'] = []
                elif line.lower() == term.lower():
                    continue
                else:
                    current_record.setdefault('All Attributes', []).append(line)

            if current_record:
                current_record['Search Query'] = term
                all_records.append(current_record)

        except TimeoutException:
            print(f"No results or page did not load properly for '{term}'")

    # Parsing logic
    for rec in all_records:
        attr = rec.get('All Attributes', [])
        combined = " | ".join(attr)
        parts = [p.strip() for p in combined.split('|')]

        rec['Probability Match'] = ""
        rec['Description'] = ""
        rec['HSN Sub Part'] = ""
        rec['Product Name'] = ""

        if len(parts) == 0:
            pass
        elif len(parts) == 1:
            rec['Product Name'] = parts[0]
        else:
            if is_probability(parts[-1]):
                rec['Probability Match'] = parts[-1]
                if len(parts) >= 2:
                    rec['Description'] = parts[-2]
                if len(parts) >= 3 and parts[-3].isdigit():
                    rec['HSN Sub Part'] = parts[-3]
                    rec['Product Name'] = " | ".join(parts[:-3])
                else:
                    rec['Product Name'] = " | ".join(parts[:-2])
            else:
                rec['Product Name'] = combined

        if 'All Attributes' in rec:
            del rec['All Attributes']

    df = pd.DataFrame(all_records)
    output_path = r"D:\HSN_Finder\Final-1.xlsx"
    df.to_excel(output_path, index=False)
    print(f"Saved all search results to {output_path}")

finally:
    driver.quit()

import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import os

# Initializing
root = tk.Tk()
root.withdraw()

# Prompt user to select the URL list file
url_list_path = filedialog.askopenfilename(title="Select the URL list file", filetypes=[("Text Files", "*.txt")])
if not url_list_path:
    print("No file selected. Exiting.")
    exit()

# Create 'scraped' directory if it doesn't exist
base_dir = os.path.dirname(url_list_path)
scraped_dir = os.path.join(base_dir, 'scraped')
os.makedirs(scraped_dir, exist_ok=True)

# Load URLs
with open(url_list_path, 'r') as file:
    url_list = [line.strip() for line in file.readlines() if line.strip()]

# Set up WebDriver using webdriver-manager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

for index, url in enumerate(url_list):
    print(f"Scraping URL: {url}")
    driver.get(url)
    time.sleep(2)

    # Scroll to load full page
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    html_source = driver.page_source
    soup = BeautifulSoup(html_source, 'html.parser')

    data = []
    first_headword = None

    for entry in soup.find_all(class_='headword'):
        headword = entry.get_text(strip=True)
        print(headword)

        if not first_headword:
            first_headword = re.sub(r'[\\/*?:<>|]', "", headword)

        # Get etymology text
        etymology_section = entry.find_next(class_='etymology')
        if etymology_section:
            etymology_text = ' '.join(etymology_section.stripped_strings)
        else:
            etymology_text = ''

        # Get all meanings
        meaning_entries = entry.find_all_next(class_='item-content')
        last_meaning_text = None

        for meaning_entry in meaning_entries:
            definition_element = meaning_entry.find(class_='definition')
            if definition_element:
                meaning_text = ' '.join(definition_element.stripped_strings)
            else:
                meaning_text = ''

            grammar = meaning_entry.find_previous(class_='grammar').get_text(strip=True) if meaning_entry.find_previous(class_='grammar') else ''
            daterange = meaning_entry.find_previous(class_='daterange').get_text(strip=True) if meaning_entry.find_previous(class_='daterange') else ''
            item_enumerator = meaning_entry.find_previous(class_='item-enumerator').get_text(strip=True) if meaning_entry.find_previous(class_='item-enumerator') else ''

            # Avoid repeating same meaning
            new_meaning = meaning_text != last_meaning_text
            last_meaning_text = meaning_text if new_meaning else ''

            quotation_container = meaning_entry.find_next(class_='quotation-container')
            if quotation_container:
                quotations = quotation_container.find_all(class_='quotation')
                for quote in quotations:
                    quote_date = quote.find(class_='quotation-date').get_text(strip=True) if quote.find(class_='quotation-date') else ''
                    quote_text = quote.find(class_='quotation-text').get_text(strip=True) if quote.find(class_='quotation-text') else ''
                    citation = quote.find(class_='citation').get_text(strip=True) if quote.find(class_='citation') else ''

                    data.append([
                        headword,
                        etymology_text,
                        item_enumerator,
                        daterange,
                        grammar,
                        meaning_text if new_meaning else '',
                        quote_date,
                        quote_text,
                        citation
                    ])
                    # After first row, avoid repeating etymology
                    etymology_text = ''
                    headword = ''
            else:
                data.append([
                    headword,
                    etymology_text,
                    item_enumerator,
                    daterange,
                    grammar,
                    meaning_text if new_meaning else '',
                    '',
                    '',
                    ''
                ])
                etymology_text = ''
                headword = ''

    # Write to Excel
    columns = [
        'Headword',
        'Etymology',
        'Item Enumerator',
        'Date Range',
        'Grammar',
        'Meaning',
        'Quotation Date',
        'Quotation Text',
        'Citation'
    ]
    df = pd.DataFrame(data, columns=columns)

    filename = f"{first_headword}_{index + 1}.xlsx" if df.shape[0] > 0 else f"extracted_data_{index + 1}.xlsx"
    filepath = os.path.join(scraped_dir, filename)
    df.to_excel(filepath, index=False)
    print(f"Data for URL {url} has been exported to '{filepath}'.")

driver.quit()
print("Data scraping complete for all URLs.")

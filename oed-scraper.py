import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time
import re

# Initializing
root = tk.Tk()
root.withdraw()  # to hide the root window

# Get the URL for the OED page that is to be scraped
url = simpledialog.askstring("Input", "Paste OED URL here:")

# Set up Chrome WebDriver
chrome_service = Service('/usr/local/bin/chromedriver')  # path to your chromedriver, edit as needed
driver = webdriver.Chrome(service=chrome_service)

# Open provided URL in Chrome
driver.get(url)
time.sleep(5)  # lets the page load (adjust this based on your connection speed and page complexity)

# Scroll the page to load all content
last_height = driver.execute_script("return document.body.scrollHeight")
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

# Get the page source code and close the browser
html_source = driver.page_source
driver.quit()

# Parse the source code using BeautifulSoup
soup = BeautifulSoup(html_source, 'html.parser')

# Extract data
data = []
first_headword = None  # this will use the headword to create the file with the same name

for entry in soup.find_all(class_='headword'):
    headword = entry.get_text(strip=True)

    if not first_headword:
        # Clean the headword to make it a valid file name
        first_headword = re.sub(r'[\\/*?:<>|]', "", headword)  # removes most unauthorized special characters

    # Get all meanings under this headword
    meaning_entries = entry.find_all_next(class_='item-body')
    for meaning_entry in meaning_entries:
        meaning_text = meaning_entry.find(class_='definition').get_text(strip=True)
        grammar = meaning_entry.find_previous(class_='grammar').get_text(strip=True) if meaning_entry.find_previous(class_='grammar') else ''
        daterange = meaning_entry.find_previous(class_='daterange').get_text(strip=True) if meaning_entry.find_previous(class_='daterange') else ''

        # Extract 'item-enumerator' (the numbering of the meanings)
        item_enumerator = meaning_entry.find_previous(class_='item-enumerator').get_text(strip=True) if meaning_entry.find_previous(class_='item-enumerator') else ''

        # Find quotations related to this meaning, scoped within the meaning_entry
        quotation_blocks = meaning_entry.find_all(class_='quotation-block-wrapper')
        if quotation_blocks:
            for quote_block in quotation_blocks:
                # Extract date
                quote_date = quote_block.find(class_='quotation-date').get_text(strip=True) if quote_block.find(class_='quotation-date') else ''
                
                # Extract quotation text and citation separately
                quote_text = quote_block.find(class_='quotation-text').get_text(strip=True) if quote_block.find(class_='quotation-text') else ''
                citation = quote_block.find(class_='citation').get_text(strip=True) if quote_block.find(class_='citation') else ''

                # Add data as a row for each quotation
                data.append([headword, item_enumerator, daterange, grammar, meaning_text, quote_date, quote_text, citation])
        else:
            # Add row without quotation if there are no quotations for this meaning
            data.append([headword, item_enumerator, daterange, grammar, meaning_text, '', '', ''])

# Create DataFrame from the extracted data
columns = ['Headword', 'Item Enumerator', 'Date Range', 'Grammar', 'Meaning', 'Quotation Date', 'Quotation Text', 'Citation']
df = pd.DataFrame(data, columns=columns)

# Use the headword as the file name
if first_headword:
    filename = f"{first_headword}.xlsx"
else:
    filename = "extracted_data.xlsx"  # in case there is not a correct name to use

# Export data to Excel file
df.to_excel(filename, index=False)

print(f"Data scraping done. The file '{filename}' has been generated.")

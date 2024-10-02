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
time.sleep(15)  # lets the page load (adjust this based on your connection speed and page complexity)

# Scroll the page to load all content
last_height = driver.execute_script("return document.body.scrollHeight")
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)
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
    print(headword)

    if not first_headword:
        # Clean the headword to make it a valid file name
        first_headword = re.sub(r'[\\/*?:<>|]', "", headword)  # removes most unauthorized special characters

    # Get all meanings under this headword
    meaning_entries = entry.find_all_next(class_='item-content')
    print(meaning_entries)
    for meaning_entry in meaning_entries:
        # Updated: Extracting text with spaces between sub-elements in the 'definition' class
        definition_element = meaning_entry.find(class_='definition')
        if definition_element:
            meaning_text = ' '.join(definition_element.stripped_strings)
        else:
            meaning_text = ''

        grammar = meaning_entry.find_previous(class_='grammar').get_text(strip=True) if meaning_entry.find_previous(class_='grammar') else ''
        daterange = meaning_entry.find_previous(class_='daterange').get_text(strip=True) if meaning_entry.find_previous(class_='daterange') else ''
        print(meaning_text, '', grammar, '', daterange)

        # Extract 'item-enumerator' (the numbering of the meanings)
        item_enumerator = meaning_entry.find_previous(class_='item-enumerator').get_text(strip=True) if meaning_entry.find_previous(class_='item-enumerator') else ''

        # Find quotations related to this meaning, scoped within the meaning_entry
        quotation_container = meaning_entry.find_next(class_='quotation-container')
        if quotation_container:
            # Iterate over all 'quotation' class items within the 'quotation-container'
            quotations = quotation_container.find_all(class_='quotation')
            for quote in quotations:
                # Extract date
                quote_date = quote.find(class_='quotation-date').get_text(strip=True) if quote.find(class_='quotation-date') else ''
                
                # Extract quotation text and citation separately
                quote_text = quote.find(class_='quotation-text').get_text(strip=True) if quote.find(class_='quotation-text') else ''
                citation = quote.find(class_='citation').get_text(strip=True) if quote.find(class_='citation') else ''
                print(quote_date, '', quote_text)

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

import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import os

# Initializing
root = tk.Tk()
root.withdraw()  # to hide the root window

# Prompt user to select the url_list.txt file
url_list_path = filedialog.askopenfilename(title="Select the URL list file", filetypes=[("Text Files", "*.txt")])
if not url_list_path:
    print("No file selected. Exiting.")
    exit()

# Get the directory of the selected file and create 'scraped' subdirectory
base_dir = os.path.dirname(url_list_path)
scraped_dir = os.path.join(base_dir, 'scraped')

# Create the 'scraped' directory if it doesn't exist
os.makedirs(scraped_dir, exist_ok=True)

# Load URLs from the selected file
with open(url_list_path, 'r') as file:
    url_list = [line.strip() for line in file.readlines() if line.strip()]

# Set up Chrome WebDriver
chrome_service = Service('/usr/local/bin/chromedriver')  # path to your chromedriver, edit as needed
driver = webdriver.Chrome(service=chrome_service)

for index, url in enumerate(url_list):
    print(f"Scraping URL: {url}")
    
    # Open the URL in Chrome
    driver.get(url)
    time.sleep(2)  # lets the page load (adjust this based on your connection speed and page complexity)

    # Scroll the page to load all content
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    # Get the page source code
    html_source = driver.page_source

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
        for meaning_entry in meaning_entries:
            # Updated: Extracting text with spaces between sub-elements in the 'definition' class
            definition_element = meaning_entry.find(class_='definition')
            if definition_element:
                meaning_text = ' '.join(definition_element.stripped_strings)
            else:
                meaning_text = ''

            grammar = meaning_entry.find_previous(class_='grammar').get_text(strip=True) if meaning_entry.find_previous(class_='grammar') else ''
            daterange = meaning_entry.find_previous(class_='daterange').get_text(strip=True) if meaning_entry.find_previous(class_='daterange') else ''
            
            # Extract 'item-enumerator'
            item_enumerator = meaning_entry.find_previous(class_='item-enumerator').get_text(strip=True) if meaning_entry.find_previous(class_='item-enumerator') else ''

            # Find quotations related to this meaning
            quotation_container = meaning_entry.find_next(class_='quotation-container')
            if quotation_container:
                quotations = quotation_container.find_all(class_='quotation')
                for quote in quotations:
                    quote_date = quote.find(class_='quotation-date').get_text(strip=True) if quote.find(class_='quotation-date') else ''
                    quote_text = quote.find(class_='quotation-text').get_text(strip=True) if quote.find(class_='quotation-text') else ''
                    citation = quote.find(class_='citation').get_text(strip=True) if quote.find(class_='citation') else ''
                    
                    # Add data as a row for each quotation
                    data.append([headword, item_enumerator, daterange, grammar, meaning_text, quote_date, quote_text, citation])
            else:
                # Add row without quotation if there are no quotations for this meaning
                data.append([headword, item_enumerator, daterange, grammar, meaning_text, '', '', ''])

    # Create DataFrame from the extracted data
    columns = ['Headword', 'Item Enumerator', 'Date Range', 'Grammar', 'Meaning', 'Quotation Date', 'Quotation Text', 'Citation']
    df = pd.DataFrame(data, columns=columns)

    # Use the first headword of this URL as the file name, appending index for uniqueness
    if df.shape[0] > 0:
        filename = f"{first_headword}_{index + 1}.xlsx"  # Adding index to make filenames unique
    else:
        filename = f"extracted_data_{index + 1}.xlsx"  # Fallback filename in case of no data

    # Save the Excel file in the 'scraped' directory
    filepath = os.path.join(scraped_dir, filename)
    df.to_excel(filepath, index=False)
    print(f"Data for URL {url} has been exported to '{filepath}'.")

# Close the browser
driver.quit()

print("Data scraping complete for all URLs.")

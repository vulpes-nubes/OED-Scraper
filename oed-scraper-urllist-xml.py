import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import re
import os
import xml.etree.ElementTree as ET
from xml.dom import minidom

# Initialize and hide Tkinter root window
root = tk.Tk()
root.withdraw()

# Prompt user to select the URL list file
url_list_path = filedialog.askopenfilename(title="Select the URL list file", filetypes=[("Text Files", "*.txt")])
if not url_list_path:
    print("No file selected. Exiting.")
    exit()

# Create 'scraped' subdirectory if it doesn't exist
base_dir = os.path.dirname(url_list_path)
scraped_dir = os.path.join(base_dir, 'scraped_xml')
os.makedirs(scraped_dir, exist_ok=True)

# Load URLs
with open(url_list_path, 'r') as file:
    url_list = [line.strip() for line in file.readlines() if line.strip()]

# Setup WebDriver using webdriver-manager
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Scrape each URL
for index, url in enumerate(url_list):
    print(f"Scraping URL: {url}")
    driver.get(url)
    time.sleep(2)

    # Scroll to load all content
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

    for entry in soup.find_all(class_='headword'):
        headword = entry.get_text(strip=True)
        print(f"Processing headword: {headword}")

        safe_headword = re.sub(r'[\\/*?:<>|]', "", headword)

        # Extract etymology
        etymology_section = entry.find_next(class_='etymology')
        etymology_text = ''

        if etymology_section:
            summary = etymology_section.find(class_='etymology-summary')
            if summary:
                etymology_text += ' '.join(summary.stripped_strings)

            other_parts = [
                tag for tag in etymology_section.find_all(recursive=False)
                if 'etymology-summary' not in (tag.get('class') or [])
            ]
            extra_text = ' '.join(tag.get_text(strip=True, separator=' ') for tag in other_parts)
            if extra_text and extra_text not in etymology_text:
                etymology_text += ' ' + extra_text

            etymology_text = etymology_text.strip()

        # XML root for this entry
        entry_elem = ET.Element('entry')
        ET.SubElement(entry_elem, 'headword').text = headword
        ET.SubElement(entry_elem, 'etymology').text = etymology_text

        # Meanings and quotations
        meaning_entries = entry.find_all_next(class_='item-content')
        last_meaning_text = None

        for meaning_entry in meaning_entries:
            definition_element = meaning_entry.find(class_='definition')
            meaning_text = ' '.join(definition_element.stripped_strings) if definition_element else ''
            is_new_meaning = meaning_text != last_meaning_text
            last_meaning_text = meaning_text if is_new_meaning else last_meaning_text

            if is_new_meaning:
                meaning_elem = ET.SubElement(entry_elem, 'meaning')
                item_enum = meaning_entry.find_previous(class_='item-enumerator')
                daterange = meaning_entry.find_previous(class_='daterange')
                grammar = meaning_entry.find_previous(class_='grammar')

                ET.SubElement(meaning_elem, 'item_enumerator').text = item_enum.get_text(strip=True) if item_enum else ''
                ET.SubElement(meaning_elem, 'daterange').text = daterange.get_text(strip=True) if daterange else ''
                ET.SubElement(meaning_elem, 'grammar').text = grammar.get_text(strip=True) if grammar else ''
                ET.SubElement(meaning_elem, 'definition').text = meaning_text
            else:
                meaning_elem = entry_elem.findall('meaning')[-1]

            # Quotations
            quotation_container = meaning_entry.find_next(class_='quotation-container')
            if quotation_container:
                for quote in quotation_container.find_all(class_='quotation'):
                    quote_elem = ET.SubElement(meaning_elem, 'quotation')
                    date = quote.find(class_='quotation-date')
                    text = quote.find(class_='quotation-text')
                    citation = quote.find(class_='citation')

                    ET.SubElement(quote_elem, 'date').text = date.get_text(strip=True) if date else ''
                    ET.SubElement(quote_elem, 'text').text = text.get_text(strip=True) if text else ''
                    ET.SubElement(quote_elem, 'citation').text = citation.get_text(strip=True) if citation else ''

        # Pretty-print and save XML
        rough_string = ET.tostring(entry_elem, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        pretty_xml = reparsed.toprettyxml(indent="  ")

        filename = f"{safe_headword}_{index + 1}.xml"
        filepath = os.path.join(scraped_dir, filename)

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)

        print(f"Saved XML: {filepath}")

# Clean up
driver.quit()
print("Finished scraping all URLs.")

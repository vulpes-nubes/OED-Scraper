import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
import time
from threading import Thread
from pathlib import Path


class OEDScraperApp:
    def __init__(self, master):
        self.master = master
        master.title("OED Scraper")

        self.url_label = tk.Label(master, text="Select URL list file:")
        self.url_label.grid(row=0, column=0, sticky="w")
        self.url_entry = tk.Entry(master, width=50)
        self.url_entry.grid(row=0, column=1)
        self.url_button = tk.Button(master, text="Browse", command=self.browse_file)
        self.url_button.grid(row=0, column=2)

        self.export_label = tk.Label(master, text="Select export format:")
        self.export_label.grid(row=1, column=0, sticky="w")
        self.export_var = tk.StringVar(value="excel")
        self.export_options = ["excel", "xml", "tei-xml", "pdf"]
        for i, f in enumerate(self.export_options):
            tk.Radiobutton(master, text=f.upper(), variable=self.export_var, value=f).grid(row=1, column=i+1)

        self.single_file_var = tk.BooleanVar()
        self.single_file_checkbox = tk.Checkbutton(master, text="Export single XML/TEI-XML file", variable=self.single_file_var)
        self.single_file_checkbox.grid(row=2, column=1, columnspan=2, sticky="w")

        self.extract_label = tk.Label(master, text="Select fields to extract:")
        self.extract_label.grid(row=3, column=0, sticky="w")
        self.fields = ["Headword", "URL", "Etymology", "Item Enumerator", "Date Range", "Grammar", "Meaning", "Quotation Date", "Quotation Text", "Citation"]
        self.field_vars = {field: tk.BooleanVar(value=True) for field in self.fields}
        for i, field in enumerate(self.fields):
            tk.Checkbutton(master, text=field, variable=self.field_vars[field]).grid(row=4+i//3, column=i%3+1, sticky="w")

        self.scrape_button = tk.Button(master, text="Scrape OED", command=self.start_scraping_thread)
        self.scrape_button.grid(row=8, column=1, pady=10)

        self.progress = ttk.Progressbar(master, orient="horizontal", mode="determinate", length=400)
        self.progress.grid(row=9, column=0, columnspan=3, pady=(0, 10))

        self.log_text = tk.Text(master, height=20, width=80)
        self.log_text.grid(row=10, column=0, columnspan=3, pady=10)

    def log(self, message):
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        print(message)

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if filepath:
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, filepath)

    def start_scraping_thread(self):
        thread = Thread(target=self.scrape_urls)
        thread.start()

    def ensure_single_page(self, driver):
        try:
            button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "tabbed-view-switch-nontabbed"))
            )
            button.click()
            self.log("✔️ Forced Single Page view.")
        except TimeoutException:
            self.log("ℹ️ Single Page button not found—may already be on single page.")

    def expand_etymology_show_more(self, driver):
        try:
            etymology_section = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "etymology"))
            )
            show_more_button = etymology_section.find_element(By.CLASS_NAME, "quotations-button")
            if show_more_button.is_displayed():
                show_more_button.click()
                self.log("✔️ Clicked 'Show more' in Etymology section.")
        except Exception:
            self.log("ℹ️ No 'Show more' button found in Etymology.")

    def scrape_urls(self):
        url_list_path = self.url_entry.get()
        if not os.path.isfile(url_list_path):
            messagebox.showerror("Error", "Invalid file path")
            return

        with open(url_list_path, 'r') as f:
            urls = [line.strip() for line in f if line.strip()]

        self.log(f"Loaded {len(urls)} URLs.")
        self.progress["maximum"] = len(urls)
        self.progress["value"] = 0

        export_format = self.export_var.get()
        base_dir = os.path.dirname(url_list_path)
        scraped_dir = os.path.join(base_dir, 'scraped')
        os.makedirs(scraped_dir, exist_ok=True)

        chrome_service = Service('/usr/local/bin/chromedriver')
        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(service=chrome_service, options=options)

        for index, url in enumerate(urls):
            self.log(f"Scraping URL {index + 1}/{len(urls)}: {url}")
            driver.get(url)
            self.ensure_single_page(driver)
            self.expand_etymology_show_more(driver)
            time.sleep(2)

            soup = BeautifulSoup(driver.page_source, 'html.parser')
            headword = soup.find(class_='headword').get_text(strip=True) if soup.find(class_='headword') else 'Unknown'

            etymology = ''
            if self.field_vars['Etymology'].get():
                etym_section = soup.find(class_='etymology')
                if etym_section:
                    etym_summary = etym_section.find(class_='etymology-summary')
                    if etym_summary:
                        etymology = ' '.join(etym_summary.stripped_strings)
                    else:
                        etymology = ' '.join(etym_section.stripped_strings)

            meaning_entries = soup.find_all(class_='item-content')
            data = []
            meanings_seen = {}  # Dictionary to store meanings and their associated quotations

            for meaning_entry in meaning_entries:
                meaning_text = meaning_entry.find(class_='definition')
                meaning = ' '.join(meaning_text.stripped_strings) if meaning_text else ''
                if not meaning.strip() or meaning in meanings_seen:
                    continue  # Skip empty or duplicate meanings

                # Initialize data for this meaning (will group quotes here)
                meaning_data = {
                    'Headword': headword if self.field_vars['Headword'].get() else '',
                    'Meaning': meaning if self.field_vars['Meaning'].get() else '',
                    'Etymology': etymology if self.field_vars['Etymology'].get() else '',
                    'Item Enumerator': '',
                    'Date Range': '',
                    'Grammar': '',
                    'Quotation Date': '',
                    'Quotation Text': '',
                    'Citation': ''
                }

                # Collect quotations for this meaning
                quotes = []
                item_enum = meaning_entry.find_previous(class_='item-enumerator')
                daterange = meaning_entry.find_previous(class_='daterange')
                grammar = meaning_entry.find_previous(class_='grammar')
                quote_container = meaning_entry.find_next(class_='quotation-container')
                if quote_container:
                    quote_elements = quote_container.find_all(class_='quotation')
                    for quote in quote_elements:
                        quote_date = quote.find(class_='quotation-date').get_text(strip=True) if quote.find(class_='quotation-date') else ''
                        quote_text = quote.find(class_='quotation-text').get_text(strip=True) if quote.find(class_='quotation-text') else ''
                        citation = quote.find(class_='citation').get_text(strip=True) if quote.find(class_='citation') else ''
                        
                        quotes.append((quote_date, quote_text, citation))

                # Only add this meaning if it's new (not already in the dictionary)
                if meaning not in meanings_seen:
                    # If the meaning has quotes, include them
                    for quote in quotes:
                        if self.field_vars['Quotation Date'].get():
                            meaning_data['Quotation Date'] = quote[0]
                        if self.field_vars['Quotation Text'].get():
                            meaning_data['Quotation Text'] = quote[1]
                        if self.field_vars['Citation'].get():
                            meaning_data['Citation'] = quote[2]
                    
                    # Add the meaning and its associated quotations to the data list
                    data.append(meaning_data)
                    meanings_seen[meaning] = True  # Mark this meaning as seen

            # Export data to Excel for this URL
            if data:
                df = pd.DataFrame(data)
                filename = f"{headword}_{index + 1}.xlsx"
                file_path = os.path.join(scraped_dir, filename)
                df.to_excel(file_path, index=False)
                self.log(f"Data for URL {url} exported to {file_path}")
            
            self.progress["value"] = index + 1
            self.master.update_idletasks()

        driver.quit()

        self.log("Scraping and export complete.")


if __name__ == '__main__':
    root = tk.Tk()
    app = OEDScraperApp(root)
    root.mainloop()

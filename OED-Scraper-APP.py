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
import xml.etree.ElementTree as ET

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
        for i, f in enumerate(["excel", "xml", "tei-xml"]):
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

        self.scrape_button = tk.Button(master, text="Scrape OED", command=self.scrape)
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

    def ensure_single_page(self, driver):
        try:
            button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "tabbed-view-switch-nontabbed"))
            )
            button.click()
            self.log("\u2714\ufe0f Forced Single Page view.")
        except TimeoutException:
            self.log("\u2139\ufe0f Single Page button not found—may already be on single page.")

    def expand_etymology_show_more(self, driver):
        try:
            etymology_section = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "etymology"))
            )
            show_more_button = etymology_section.find_element(By.CLASS_NAME, "quotations-button")
            if show_more_button.is_displayed():
                show_more_button.click()
                self.log("\u2714\ufe0f Clicked 'Show more' in Etymology section.")
        except Exception:
            self.log("\u2139\ufe0f No 'Show more' button found in Etymology.")

    def deduplicate_rows(self, data):
        seen = set()
        unique_data = []
        for row in data:
            key = tuple(row.items())
            if key not in seen:
                seen.add(key)
                unique_data.append(row)
        duplicates_removed = len(data) - len(unique_data)
        return unique_data, duplicates_removed

    def scrape(self):
        url_list_path = self.url_entry.get()
        if not os.path.isfile(url_list_path):
            messagebox.showerror("Error", "Invalid file path")
            return

        with open(url_list_path, 'r') as f:
            urls = [line.strip() for line in f if line.strip()]

        self.log(f"Loaded {len(urls)} URLs.")
        self.progress["maximum"] = len(urls)
        self.progress["value"] = 0

        chrome_service = Service('/usr/local/bin/chromedriver')
        driver = webdriver.Chrome(service=chrome_service)
        all_data = []

        for index, url in enumerate(urls):
            self.log(f"Scraping URL {index + 1}/{len(urls)}: {url}")
            driver.get(url)
            self.ensure_single_page(driver)
            self.expand_etymology_show_more(driver)
            time.sleep(2)

            last_height = driver.execute_script("return document.body.scrollHeight")
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height

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
            for meaning_entry in meaning_entries:
                meaning_text = meaning_entry.find(class_='definition')
                meaning = ' '.join(meaning_text.stripped_strings) if meaning_text else ''
                if not meaning.strip():
                    continue

                item_enum = meaning_entry.find_previous(class_='item-enumerator')
                daterange = meaning_entry.find_previous(class_='daterange')
                grammar = meaning_entry.find_previous(class_='grammar')
                quote_container = meaning_entry.find_next(class_='quotation-container')
                quotes = quote_container.find_all(class_='quotation') if quote_container else [None]

                for quote in quotes:
                    row = {
                        'Headword': headword,
                        'URL': url,
                        'Etymology': etymology,
                        'Item Enumerator': item_enum.get_text(strip=True) if item_enum else '',
                        'Date Range': daterange.get_text(strip=True) if daterange else '',
                        'Grammar': grammar.get_text(strip=True) if grammar else '',
                        'Meaning': meaning,
                        'Quotation Date': quote.find(class_='quotation-date').get_text(strip=True) if quote and quote.find(class_='quotation-date') else '',
                        'Quotation Text': quote.find(class_='quotation-text').get_text(strip=True) if quote and quote.find(class_='quotation-text') else '',
                        'Citation': quote.find(class_='citation').get_text(strip=True) if quote and quote.find(class_='citation') else ''
                    }
                    all_data.append(row)

            self.progress["value"] = index + 1
            self.master.update_idletasks()

        driver.quit()

        self.log("Deduplicating entries...")
        unique_data, duplicates_removed = self.deduplicate_rows(all_data)

        selected_fields = [field for field, var in self.field_vars.items() if var.get()]
        export_format = self.export_var.get()
        base_dir = os.path.dirname(url_list_path)

        self.log(f"Total unique entries: {len(unique_data)}")
        self.log(f"Duplicates removed: {duplicates_removed}")

        if export_format == "excel":
            df = pd.DataFrame(unique_data)
            df = df[selected_fields]
            df.sort_values(by=["Headword"], inplace=True)
            output_path = os.path.join(base_dir, "scraped_output.xlsx")
            df.to_excel(output_path, index=False)
            self.log(f"Data exported to Excel: {output_path}")

        elif export_format in ("xml", "tei-xml"):
            root_tag = "TEI" if export_format == "tei-xml" else "Entries"
            root = ET.Element(root_tag)
            grouped_data = {}
            for row in unique_data:
                headword = row.get("Headword", "Unknown")
                grouped_data.setdefault(headword, []).append(row)

            for headword, entries in grouped_data.items():
                group = ET.SubElement(root, "headword_group")
                headword_elem = ET.SubElement(group, "headword")
                headword_elem.text = headword
                for row in entries:
                    entry = ET.SubElement(group, "entry")
                    for field in selected_fields:
                        tag_name = field.lower().replace(" ", "_")
                        value = row.get(field, '')
                        if value:
                            ET.SubElement(entry, tag_name).text = value

            output_path = os.path.join(base_dir, "scraped_output.xml")
            import xml.dom.minidom

            rough_string = ET.tostring(root, 'utf-8')
            reparsed = xml.dom.minidom.parseString(rough_string)
            pretty_xml = reparsed.toprettyxml(indent="  ")

            with open(output_path, "w", encoding="utf-8") as f:
                lines = pretty_xml.splitlines()
                spaced = []
                for i, line in enumerate(lines):
                    spaced.append(line)
                    if line.strip() == "</entry>" and i + 1 < len(lines):
                        spaced.append("")  # add a blank line after each entry
                f.write("\n".join(spaced))
            self.log(f"Data exported to XML: {output_path}")

        self.log("Scraping and export complete.")

if __name__ == '__main__':
    root = tk.Tk()
    app = OEDScraperApp(root)
    root.mainloop()

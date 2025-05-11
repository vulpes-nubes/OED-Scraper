import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
import time
import xml.etree.ElementTree as ET
from collections import defaultdict

class OEDScraperApp:
    def __init__(self, master):
        self.master = master
        master.title("OED Scraper")

        # UI elements
        self.url_label = tk.Label(master, text="Select URL list file:")
        self.url_label.grid(row=0, column=0, sticky="w")
        self.url_entry = tk.Entry(master, width=50)
        self.url_entry.grid(row=0, column=1)
        self.url_button = tk.Button(master, text="Browse", command=self.browse_file)
        self.url_button.grid(row=0, column=2)

        self.export_label = tk.Label(master, text="Select export format:")
        self.export_label.grid(row=1, column=0, sticky="w")
        self.export_var = tk.StringVar(value="excel")
        formats = ["excel", "xml", "tei-xml"]
        for i, f in enumerate(formats):
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

            meanings_seen = set()
            meaning_entries = soup.find_all(class_='item-content')
            for meaning_entry in meaning_entries:
                meaning_text = meaning_entry.find(class_='definition')
                meaning = ' '.join(meaning_text.stripped_strings) if meaning_text else ''
                if meaning in meanings_seen:
                    continue
                meanings_seen.add(meaning)

                item_enum = meaning_entry.find_previous(class_='item-enumerator')
                daterange = meaning_entry.find_previous(class_='daterange')
                grammar = meaning_entry.find_previous(class_='grammar')
                quote_container = meaning_entry.find_next(class_='quotation-container')
                if quote_container:
                    for quote in quote_container.find_all(class_='quotation'):
                        row = {
                            'Headword': headword,
                            'URL': url,
                            'Etymology': etymology,
                            'Item Enumerator': item_enum.get_text(strip=True) if item_enum else '',
                            'Date Range': daterange.get_text(strip=True) if daterange else '',
                            'Grammar': grammar.get_text(strip=True) if grammar else '',
                            'Meaning': meaning,
                            'Quotation Date': quote.find(class_='quotation-date').get_text(strip=True) if quote.find(class_='quotation-date') else '',
                            'Quotation Text': quote.find(class_='quotation-text').get_text(strip=True) if quote.find(class_='quotation-text') else '',
                            'Citation': quote.find(class_='citation').get_text(strip=True) if quote.find(class_='citation') else ''
                        }
                        all_data.append(row)

            self.progress["value"] = index + 1
            self.master.update_idletasks()

        driver.quit()

        unique_data = [dict(t) for t in {tuple(d.items()) for d in all_data}]

        selected_fields = [field for field, var in self.field_vars.items() if var.get()]
        export_format = self.export_var.get()
        base_dir = os.path.dirname(url_list_path)

        self.log(f"Total unique entries: {len(unique_data)}")

        if export_format == "excel":
            df = pd.DataFrame(unique_data)
            df = df[selected_fields]
            output_path = os.path.join(base_dir, "scraped_output.xlsx")
            df.to_excel(output_path, index=False)
            self.log(f"Data exported to Excel: {output_path}")

        elif export_format in ("xml", "tei-xml"):
            root_tag = "TEI" if export_format == "tei-xml" else "Entries"
            root = ET.Element(root_tag)
            for row in unique_data:
                entry = ET.SubElement(root, "entry")
                for field in selected_fields:
                    ET.SubElement(entry, field.lower().replace(" ", "_")).text = row.get(field, '')

            tree = ET.ElementTree(root)
            output_path = os.path.join(base_dir, "scraped_output.xml")
            tree.write(output_path, encoding="utf-8", xml_declaration=True)
            self.log(f"Data exported to XML: {output_path}")

        self.log("Scraping and export complete.")

if __name__ == '__main__':
    root = tk.Tk()
    app = OEDScraperApp(root)
    root.mainloop()

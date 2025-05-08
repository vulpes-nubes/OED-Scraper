import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import os
import time
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom
import pandas as pd

class OEDScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OED Scraper Tool")

        # Variables
        self.url_file_path = tk.StringVar()
        self.export_format = tk.StringVar(value="Excel")
        self.single_file = tk.BooleanVar()
        self.fields = {
            'Headword': tk.BooleanVar(value=True),
            'URL': tk.BooleanVar(),
            'Etymology': tk.BooleanVar(value=True),
            'Meaning': tk.BooleanVar(value=True),
            'Enumerator': tk.BooleanVar(),
            'Daterange': tk.BooleanVar(),
            'Grammar': tk.BooleanVar(),
            'Definition': tk.BooleanVar(value=True),
            'Quotation Date': tk.BooleanVar(),
            'Quotation Text': tk.BooleanVar(),
            'Citation': tk.BooleanVar(),
        }

        self.build_gui()

    def build_gui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="URL List File:").grid(row=0, column=0, sticky='w')
        ttk.Entry(frame, textvariable=self.url_file_path, width=40).grid(row=0, column=1, sticky='we')
        ttk.Button(frame, text="Browse", command=self.browse_file).grid(row=0, column=2)

        ttk.Label(frame, text="Export Format:").grid(row=1, column=0, sticky='w')
        for i, fmt in enumerate(["Excel", "XML", "TEI-XML"]):
            ttk.Radiobutton(frame, text=fmt, variable=self.export_format, value=fmt, command=self.toggle_single_file).grid(row=1, column=1+i, sticky='w')

        self.single_file_check = ttk.Checkbutton(frame, text="Export to a single file (XML only)", variable=self.single_file)
        self.single_file_check.grid(row=2, column=1, columnspan=3, sticky='w')

        ttk.Label(frame, text="Fields to Extract:").grid(row=3, column=0, sticky='nw')
        fields_frame = ttk.Frame(frame)
        fields_frame.grid(row=3, column=1, columnspan=3, sticky='w')
        for i, (label, var) in enumerate(self.fields.items()):
            ttk.Checkbutton(fields_frame, text=label, variable=var).grid(row=i//3, column=i%3, sticky='w')

        ttk.Button(frame, text="Scrape OED", command=self.scrape).grid(row=4, column=0, columnspan=3, pady=10)

        self.log = tk.Text(frame, height=10, state='disabled')
        self.log.grid(row=5, column=0, columnspan=3, sticky='nsew')
        frame.rowconfigure(5, weight=1)
        frame.columnconfigure(1, weight=1)

    def log_message(self, message):
        self.log.config(state='normal')
        self.log.insert('end', message + '\n')
        self.log.see('end')
        self.log.config(state='disabled')

    def browse_file(self):
        path = filedialog.askopenfilename(title="Select URL List File", filetypes=[("Text Files", "*.txt")])
        if path:
            self.url_file_path.set(path)

    def toggle_single_file(self):
        if self.export_format.get() in ["XML", "TEI-XML"]:
            self.single_file_check.state(['!disabled'])
        else:
            self.single_file.set(False)
            self.single_file_check.state(['disabled'])

    def scrape(self):
        if not self.url_file_path.get():
            messagebox.showerror("Error", "Please select a URL list file.")
            return

        selected_fields = [k for k, v in self.fields.items() if v.get()]
        if not selected_fields:
            messagebox.showerror("Error", "Select at least one field to extract.")
            return

        self.log_message("Starting OED scraping...")

        with open(self.url_file_path.get(), 'r') as file:
            url_list = [line.strip() for line in file.readlines() if line.strip()]

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

        all_data = []
        for index, url in enumerate(url_list):
            self.log_message(f"Scraping URL: {url}")
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

            headword_tag = soup.find(class_='headword')
            headword = headword_tag.get_text(strip=True) if headword_tag else ''
            etymology = ''
            ety_section = soup.find(class_='etymology')
            if ety_section:
                etymology = ' '.join(ety_section.stripped_strings)
                etymology_summary = ety_section.find(class_='etymology-summary')
                if etymology_summary:
                    etymology += ' ' + ' '.join(etymology_summary.stripped_strings)

            data_rows = []
            meaning_entries = soup.find_all(class_='item-content')
            for meaning_entry in meaning_entries:
                meaning_text = ' '.join(meaning_entry.find(class_='definition').stripped_strings) if meaning_entry.find(class_='definition') else ''
                grammar = meaning_entry.find_previous(class_='grammar').get_text(strip=True) if meaning_entry.find_previous(class_='grammar') else ''
                daterange = meaning_entry.find_previous(class_='daterange').get_text(strip=True) if meaning_entry.find_previous(class_='daterange') else ''
                enumerator = meaning_entry.find_previous(class_='item-enumerator').get_text(strip=True) if meaning_entry.find_previous(class_='item-enumerator') else ''
                quotations = meaning_entry.find_next(class_='quotation-container')
                if quotations:
                    for quote in quotations.find_all(class_='quotation'):
                        quote_date = quote.find(class_='quotation-date').get_text(strip=True) if quote.find(class_='quotation-date') else ''
                        quote_text = quote.find(class_='quotation-text').get_text(strip=True) if quote.find(class_='quotation-text') else ''
                        citation = quote.find(class_='citation').get_text(strip=True) if quote.find(class_='citation') else ''
                        row = {
                            'Headword': headword,
                            'URL': url,
                            'Etymology': etymology,
                            'Meaning': meaning_text,
                            'Enumerator': enumerator,
                            'Daterange': daterange,
                            'Grammar': grammar,
                            'Definition': meaning_text,
                            'Quotation Date': quote_date,
                            'Quotation Text': quote_text,
                            'Citation': citation
                        }
                        data_rows.append(row)
                else:
                    row = {
                        'Headword': headword,
                        'URL': url,
                        'Etymology': etymology,
                        'Meaning': meaning_text,
                        'Enumerator': enumerator,
                        'Daterange': daterange,
                        'Grammar': grammar,
                        'Definition': meaning_text,
                        'Quotation Date': '',
                        'Quotation Text': '',
                        'Citation': ''
                    }
                    data_rows.append(row)
            all_data.extend(data_rows)

        driver.quit()

        export_format = self.export_format.get()
        base_dir = os.path.dirname(self.url_file_path.get())
        export_dir = os.path.join(base_dir, 'scraped')
        os.makedirs(export_dir, exist_ok=True)

        if export_format == 'Excel':
            df = pd.DataFrame(all_data)
            df = df[[f for f in self.fields if self.fields[f].get()]]
            export_path = os.path.join(export_dir, 'scraped_data.xlsx')
            df.to_excel(export_path, index=False)
            self.log_message(f"Excel file saved to: {export_path}")
        else:
            root_element = ET.Element('entries')
            for row in all_data:
                entry = ET.SubElement(root_element, 'entry')
                for key in self.fields:
                    if self.fields[key].get():
                        child = ET.SubElement(entry, key.replace(' ', '_').lower())
                        child.text = row[key]

            tree = ET.ElementTree(root_element)
            rough_string = ET.tostring(root_element, 'utf-8')
            reparsed = minidom.parseString(rough_string)
            pretty_xml = reparsed.toprettyxml(indent="  ")

            if self.export_format.get() == 'TEI-XML':
                root_element.tag = 'TEI'

            file_name = 'scraped_data.xml' if self.single_file.get() else f'scraped_data_{int(time.time())}.xml'
            export_path = os.path.join(export_dir, file_name)
            with open(export_path, 'w', encoding='utf-8') as f:
                f.write(pretty_xml)
            self.log_message(f"XML file saved to: {export_path}")

        self.log_message("Scraping and export complete.")

if __name__ == '__main__':
    root = tk.Tk()
    app = OEDScraperApp(root)
    root.mainloop()

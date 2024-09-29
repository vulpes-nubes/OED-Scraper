#All code is open source and under creative commons atribution licence, please respect that
import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time
import re

#Initialising
root = tk.Tk()
root.withdraw() #to hide the root window

#Get the URL for the OED page that is to be scraped
url = simpledialog.askstring("Input", "Paste OED URL here:")

#Startup the Firefox WebDriver
firefox_service = Service('/usr/local/bin/geckodriver') #path to your geckdriver, edit as needed
driver = webdriver.Firefox(service=firefox_service)

#Open provided URL in firefox
driver.get(url)
time.sleep(15) #lets the page load (adjust this based on your connection speed and page complexity)

#Load the full page by scrolling it
last_height = driver.execute_script("return document.body.scrollHeight")
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

#Get the page source code and close the browser
html_source = driver.page_source
driver.quit()

#parse the source code using beautifulsoup
soup = BeautifulSoup(html_source, 'html.parser')

#extract data with given CSS classes
data = []
first_headword = None #this will use the headword to create the file with the same name

for entry in soup.find_all(class_='headword'):
    headword = entry.get_text(strip=True)

    if not first_headword: 
        #clean the headword to make it a valid file name
        first_headword = re.sub(r'[\\/*?:<>|]', "", headword) #removes most unauthorised special characters

    daterange = entry.find_next(class_='daterange').get_text(strip=True) if entry.find_next(class_='daterange') else ''
    grammar = entry.find_next(class_='grammar').get_text(strip=True) if entry.find_next(class_='grammar') else ''
    definition = entry.find_next(class_='definition').get_text(strip=True) if entry.find_next(class_='definition') else ''

    quotations = entry.find_next(class_='quotation-block-wrapper')
    if quotations:
        quote_date = quotations.find(class_='quotation-date').get_text(strip=True) if quotations.find(class_='quotation-date') else ''
        quote_body = quotations.find(class_='quotation-body').get_text(strip=True) if quotations.find(class_='quotation-body') else ''
    else:
        quote_date, quote_body = '', ''

    #add data as a row 
    data.append([headword, daterange, grammar, definition, quote_date, quote_body])

#take that data into Pandas
columns = ['Headword', 'Meaning Date Range', 'Grammar', 'Meaning', 'Quotation date', 'Quotation body']
df = pd.DataFrame(data, columns=columns)

#Use headword as file name
if first_headword:
    filename = f"{first_headword}.xlsx"
else:
    filename = "extracted_data.xlsx" #in case there is not a correct name to use

#export data to Excel file
df.to_excel(filename, index=False)

print(f"Data scraping done. The file '{filename}' has been generated.")
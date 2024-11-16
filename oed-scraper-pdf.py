import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# Function to save webpage as PDF using print-to-PDF functionality in Chrome
def save_as_pdf(url, output_path):
    # Set up Chrome options for headless mode and print to PDF functionality
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode (no UI)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    
    # Set the PDF output path
    prefs = {
        "printing.print_to_pdf": True,
        "savefile.default_directory": os.path.dirname(output_path)
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # Initialize Chrome WebDriver
    driver = webdriver.Chrome(options=chrome_options)

    # Open the URL
    driver.get(url)

    # Wait for the page to fully load
    time.sleep(1)

    # Trigger the print dialog and save the page as PDF
    driver.execute_script('window.print();')

    # Wait for the PDF to be generated and saved
    time.sleep(3)

    # Close the browser
    driver.quit()

    print(f"PDF saved to {output_path}")

def process_urls_from_file(input_file):
    # Read URLs from the input text file
    with open(input_file, 'r') as file:
        urls = file.readlines()

    # Loop through each URL and save as PDF
    for i, url in enumerate(urls):
        url = url.strip()  # Remove any extra spaces or newlines
        if url:
            # Generate a unique filename for each URL
            output_pdf = f"output_page_{i + 1}.pdf"
            save_as_pdf(url, output_pdf)

if __name__ == "__main__":
    input_file = "urls.txt"  # Name of the text file with URLs
    process_urls_from_file(input_file)

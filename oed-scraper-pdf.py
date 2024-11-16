import os
import time
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def get_txt_file_path():
    """
    Prompt the user to select a TXT file containing the list of URLs.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Ensure the dialog appears on top
    file_path = filedialog.askopenfilename(
        title="Select the TXT file containing URLs",
        filetypes=[("Text Files", "*.txt")]
    )
    if not file_path:
        print("No file selected. Exiting.")
        exit()
    return file_path

def setup_driver():
    """
    Set up the Selenium WebDriver for Chrome with headless mode and PDF capabilities.
    """
    options = Options()
    options.add_argument("--headless")  # Run Chrome in headless mode
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--kiosk-printing")  # Allow for direct PDF printing

    # Return the WebDriver with managed ChromeDriver
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def download_page_as_pdf(driver, url, output_folder, file_index):
    """
    Open a URL in the browser, ensure it is fully loaded, and save the page as a PDF.
    """
    try:
        print(f"Processing URL {file_index + 1}: {url}")
        driver.get(url)
        time.sleep(0.5)  # Wait for the page to load

        # Set the PDF filename
        pdf_filename = os.path.join(output_folder, f"page_{file_index + 1}.pdf")

        # Configure Chrome for PDF printing
        print_settings = {
            "printing.print_preview_sticky_settings.appState": f"""{{
                "recentDestinations": [{{"id": "Save as PDF", "origin": "local", "account": ""}}],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }}""",
            "savefile.default_directory": output_folder
        }
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": output_folder})
        driver.execute_script('window.print();')  # Trigger PDF download
        time.sleep(2)  # Allow time for the PDF to be saved
        print(f"Saved: {pdf_filename}")
    except Exception as e:
        print(f"Failed to process URL {url}: {e}")

def main():
    # Step 1: Ask the user for the TXT file containing URLs
    url_list_path = get_txt_file_path()

    # Step 2: Load the URLs from the file
    with open(url_list_path, 'r') as file:
        url_list = [line.strip() for line in file if line.strip()]

    if not url_list:
        print("The selected file is empty. Exiting.")
        exit()

    # Step 3: Create the output folder for PDFs
    output_folder = "OEDPDF"
    os.makedirs(output_folder, exist_ok=True)

    # Step 4: Set up the Selenium WebDriver
    driver = setup_driver()

    # Step 5: Process each URL sequentially
    for index, url in enumerate(url_list):
        download_page_as_pdf(driver, url, output_folder, index)

    # Step 6: Close the WebDriver
    driver.quit()
    print("All URLs have been processed. PDFs saved in:", output_folder)

if __name__ == "__main__":
    main()

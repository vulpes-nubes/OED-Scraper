import os
import time
import base64
import re
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

def sanitize_filename(filename):
    """
    Sanitize the filename by removing invalid characters for most filesystems.
    """
    return re.sub(r'[<>:"/\\|?*\t\n]', '', filename)

def setup_driver(headless=True):
    """
    Set up the Selenium WebDriver for Chrome.
    If `headless` is True, the browser runs in headless mode.
    """
    options = Options()
    if headless:
        options.add_argument("--headless")  # Run Chrome in headless mode
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")

    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def download_page_as_pdf(driver, url, output_folder):
    """
    Open a URL in the browser, ensure it is fully loaded, and save the page as a PDF.
    """
    try:
        print(f"Processing URL: {url}")
        driver.get(url)
        time.sleep(0.5)  # Allow the page to load fully

        # Get the page title to use as the filename
        page_title = driver.title
        sanitized_title = sanitize_filename(page_title)

        # Fallback filename if the title is empty or invalid
        if not sanitized_title:
            sanitized_title = "webpage"

        pdf_filename = os.path.join(output_folder, f"{sanitized_title}.pdf")

        # Use Chrome DevTools Protocol to get PDF data
        pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
        pdf_content = base64.b64decode(pdf_data['data'])  # Decode the base64-encoded PDF data

        # Write the PDF content to a file
        with open(pdf_filename, 'wb') as pdf_file:
            pdf_file.write(pdf_content)

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
    output_folder = "/home/gray221/Documents/OED PDF"
    os.makedirs(output_folder, exist_ok=True)

    # Step 4: Process the first URL without headless mode
    print("Launching the browser for the first URL to handle manual interactions (e.g., cookies).")
    driver = setup_driver(headless=False)
    try:
        print(f"Opening the first URL: {url_list[0]}")
        driver.get(url_list[0])
        print("Please handle any cookie popups or manual interactions in the browser.")
        input("Press Enter here in the terminal once you're done handling the cookies...")

        # Download the first page as a PDF
        download_page_as_pdf(driver, url_list[0], output_folder)
    finally:
        driver.quit()

    # Step 5: Process the remaining URLs in headless mode
    print("Processing the remaining URLs in headless mode.")
    driver = setup_driver(headless=True)

    for url in url_list[1:]:
        download_page_as_pdf(driver, url, output_folder)

    # Step 6: Close the WebDriver
    driver.quit()
    print("All URLs have been processed. PDFs saved in:", output_folder)

if __name__ == "__main__":
    main()

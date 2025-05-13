import os
import time
import base64
import re
import tkinter as tk
from tkinter import filedialog
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

OUTPUT_FOLDER = "/home/gray221/Documents/New OED Scrape/PDFs-new"
MAX_WORKERS = 8  # You can increase this depending on system capability

def get_txt_file_path():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(
        title="Select the TXT file containing URLs",
        filetypes=[("Text Files", "*.txt")]
    )
    if not file_path:
        print("❌ No file selected. Exiting.")
        exit()
    return file_path

def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*\t\n]', '', filename)

def setup_driver(headless=True):
    options = Options()
    if headless:
        options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def ensure_single_page(driver):
    try:
        button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "tabbed-view-switch-nontabbed"))
        )
        button.click()
        print("🔁 Switched to Single Page view.")
    except TimeoutException:
        print("ℹ️ Already in Single Page view or button not found.")

def expand_etymology_show_more(driver):
    try:
        etymology_section = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "etymology"))
        )
        show_more_button = etymology_section.find_element(By.CLASS_NAME, "quotations-button")
        if show_more_button.is_displayed():
            show_more_button.click()
            print("🔽 Expanded Etymology section.")
    except Exception:
        print("ℹ️ 'Show more' not available in Etymology section.")

def download_page_as_pdf(url, index):
    print(f"[{index}] ⏳ Starting: {url}")
    driver = setup_driver(headless=True)
    try:
        driver.get(url)
        time.sleep(0.5)
        ensure_single_page(driver)
        expand_etymology_show_more(driver)
        time.sleep(0.5)

        headword = driver.title.strip()
        sanitized_title = sanitize_filename(headword) or f"webpage_{index}"
        pdf_filename = os.path.join(OUTPUT_FOLDER, f"{sanitized_title}.pdf")

        pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
        pdf_content = base64.b64decode(pdf_data['data'])

        with open(pdf_filename, 'wb') as f:
            f.write(pdf_content)

        print(f"[{index}] ✅ Saved: {pdf_filename}")
    except Exception as e:
        print(f"[{index}] ❌ Error for {url}: {e}")
    finally:
        driver.quit()

def main():
    print("📂 Select the TXT file with URLs.")
    url_list_path = get_txt_file_path()

    with open(url_list_path, 'r') as file:
        url_list = [line.strip() for line in file if line.strip()]

    if not url_list:
        print("❌ The file is empty. Exiting.")
        return

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    # Manual first-run to handle cookies
    print("🧭 Opening the first URL in visible browser. Handle cookies or login if needed.")
    driver = setup_driver(headless=False)
    try:
        driver.get(url_list[0])
        print(f"🌐 Opened: {url_list[0]}")
        input("👉 Press ENTER here when you're done with cookies/popups...")
        download_page_as_pdf(url_list[0], 0)
    finally:
        driver.quit()

    # Multithreaded processing
    print(f"🚀 Processing remaining {len(url_list)-1} URLs using {MAX_WORKERS} threads...\n")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(download_page_as_pdf, url, i): (url, i)
            for i, url in enumerate(url_list[1:], start=1)
        }
        for future in as_completed(futures):
            _, i = futures[future]
            try:
                future.result()
            except Exception as e:
                print(f"[{i}] ❌ Unexpected error: {e}")

    print("\n🏁 All tasks complete. PDFs saved in:", OUTPUT_FOLDER)

if __name__ == "__main__":
    main()

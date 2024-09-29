from selenium import webdriver
from selenium.webdriver.chrome.service import Service

chromedriver_path = '/usr/local/bin/chromedriver'
chrome_service = Service(chromedriver_path)

driver = webdriver.Chrome(service=chrome_service)

driver.get("https://www.google.com")

print(driver.title)

driver.quit()

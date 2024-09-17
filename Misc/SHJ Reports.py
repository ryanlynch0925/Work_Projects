from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

chrome_driver_path = r"C:\Users\DavidLynch\OneDrive - Tidal Wave Autospa\Documents\Scripts\chromedriver.exe"

# Set up the Chrome driver
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service)

try:
    # Find and click the specified element
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div[2]/div[1]/section/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div[3]/div/div/div/div/div/div[2]/div/div[1]/ul/li[13]/div/div/div/div/div/div[1]/div[4]/div[2]"))
    )
    element.click()
    print("Clicked the element!")
    
    # Wait for a change on the page after clicking the element
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Expected Change Text/Element After Click')]"))
    )
    
    print("The click action was successful!")
except Exception as e:
    print(f"Error: {e}")
    print("The click action might not have been successful.")

# Continue your interactions as needed



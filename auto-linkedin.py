from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time

chromedriver_path = 'path_to_chromedriver'

def login_linkedin(email, password):
    driver.get('https://www.linkedin.com/login')
    driver.find_element(By.NAME, 'session_key').send_keys(email)
    driver.find_element(By.NAME, 'session_password').send_keys(password)
    driver.find_element(By.XPATH, "//button[contains(text(), 'Sign in')]").click() 

def search_linkedin_profile(name):
    search_query = f'{name} organization_name' 
    search_box = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CLASS_NAME, "search-global-typeahead__input")))
    search_box.clear()
    search_box.send_keys(search_query)
    search_box.send_keys(Keys.RETURN)

    time.sleep(2)

    try:
        no_results_element = driver.find_element(By.XPATH, "//h2[text()='No results found']")
        print("No results found for:", name)
        return
    except NoSuchElementException:
        pass

    try:
        no_access_element = driver.find_element(By.XPATH, "//h2[@id='out-of-network-modal-header' and text()='You don’t have access to this profile']")
        print("No access to profile for:", name)
        got_it_button = driver.find_element(By.XPATH, "//button[contains(., 'Got it')]")
        time.sleep(2)
        got_it_button.click()
        time.sleep(4)
        return
    except NoSuchElementException:
        pass
    
    first_search_result = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '//span[contains(@class, "entity-result__title-text")]//a[contains(@class, "app-aware-link")]')))
    first_search_result.click()
    
    time.sleep(2)

    current_url = driver.current_url

    if "linkedin.com/jobs" in current_url:
        driver.get("https://www.linkedin.com/feed/")
        print("No results found for:", name)
        print("Current page is LinkedIn Jobs. Redirecting to LinkedIn feed.")
    elif "linkedin.com/in" in current_url:
        profile_url = driver.current_url
        print("Profile URL:", profile_url)

excel_file = 'load_file.xlsx'
df = pd.read_excel(excel_file)

service = Service(executable_path=chromedriver_path)
options = Options()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()

login_linkedin('email', 'password')

for index, row in df.iterrows():
    name = row['NAME']

    if pd.isna(name) or name == "":
        print(f"Name at index {index} are empty or not valid.")
        continue
    
    search_linkedin_profile(name)
    profile_url = driver.current_url
    df.at[index, 'LINK'] = profile_url

df.to_excel('save_to_file.xlsx', index=False)

driver.quit()
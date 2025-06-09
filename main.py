from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from seleniumbase import Driver 
from selenium.common.exceptions import NoSuchElementException
import time
import os
from dotenv import load_dotenv
import pyotp

load_dotenv()

def get_otp(secret_key):
    totp = pyotp.TOTP(secret_key)
    return totp.now()
# ------------------------------------------------------------------
# Selenium and Browser Options

# chrome_options = Options()
# chrome_options.add_argument("--disable-blink-features=AutomationControlled")
#Driver Setup
driver = Driver(uc=True)
driver.maximize_window()
wait = WebDriverWait(driver, 60)
shortWait = WebDriverWait(driver, 15)


try:
    web_url = "https://outlook.live.com/mail/0/"
    driver.uc_open_with_reconnect(web_url, 4)

    print("Opening Website...")
            
    signIn_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="c-shellmenu_custom_outline_newtab_signin_bhvr100_right"]')))
    signIn_btn.click()
    print("Moving to signIn steps")

        # --- wait until a second window handle appears, then switch ---
    wait.until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(5)

    email_box = wait.until(EC.element_to_be_clickable((By.ID, "i0116")))
    email_box.clear()
    email_box.send_keys(os.getenv('EMAIL'), Keys.ENTER)
    print("Email Entered.")

    password_box = wait.until(EC.element_to_be_clickable((By.ID, "passwordEntry")))
    # password_box = wait.until(EC.element_to_be_clickable((By.ID, "i0118")))
    password_box.clear()
    password_box.send_keys(os.getenv('PASSWORD'), Keys.ENTER)
    print("Password Entered.")

    try:
        otc_field = shortWait.until(EC.element_to_be_clickable((By.ID, 'otc-confirmation-input')))
    except TimeoutException:
        otc_field = shortWait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="floatingLabelInput5"]')))
        
    otc_field.clear()
    otc_field.send_keys(get_otp(os.getenv('SECRET_KEY')), Keys.ENTER)
    print("OTP Entered.")

    no_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="view"]/div/div[5]/button[2]')))
    no_btn.click()
    print("Selected No.")

    try:
        wait.until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="New mail"]')))
        print("Logged In to Outlook Successfully.")
    except TimeoutException:
        print("Login Failed or took too long.")
except TimeoutException:
    print("Login Timeout")
    driver.quit()
    

try:
    time.sleep(10)
    search_bar = wait.until(EC.element_to_be_clickable((By.ID, 'topSearchInput')))
    search_bar.clear()
    search_bar.send_keys(os.getenv('SEARCH_CONTENT'))
    search_bar.send_keys(os.getenv('SEARCH_CONTENT'), Keys.ENTER)
    print("Searched for content.")

    time.sleep(6)
    #Select the first one div
    first_search = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="listbox"]//div[@data-focusable-row="true"][1]')))
    first_search.click()
    print("Selected first search result.")
except TimeoutException:
    print("Error during finding email.")
    driver.quit()

try:
    time.sleep(3)
    link = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, '//div[@aria-label="Message body"]//a[1]')               # exact text
            # or By.PARTIAL_LINK_TEXT, 'Click'         # partial
        )
    )

    link.click()
    print("Clicked on the link.")
except NoSuchElementException:
    print("Link not found.")


try:
    # time.sleep(15)
    #Shift to the next page
    wait.until(lambda d: len(d.window_handles) > 2)
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(20)

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    print("Scrolling Done.")
    first_link = wait.until(EC.element_to_be_clickable((
    By.XPATH, '/html/body/div[1]/main/section/div/div[2]/div/div/table/tbody/tr[3]/td[5]/a'
    )))
    first_link.click()
    print("Waiting for file to download...")
    time.sleep(5)
    print("Process Done")
    driver.quit()
except TimeoutException:
    print("Error after clicking on the link.")
    driver.quit()
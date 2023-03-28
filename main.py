from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time
import pyautogui

workbook = openpyxl.load_workbook('contact.xlsx')
sheet = workbook['Sheet1']

message = "II PUC regular classes with CET and NEET coaching will start from April 1st week.. Faculty from reputed Institutions will be handling classes.. For more information contact Insight Academy Santhekatte -9880782393"
image_path="C:\Everything\hehe.jpg"
image_path_2="C:\Everything\hehe2.jpg"
image_path_3="C:\Everything\hehe3.jpg"

driver = webdriver.Chrome('C:/Everything/chromedriver.exe')
driver.get('https://web.whatsapp.com/')

input('Scan the QR code and press enter to continue...')

for row in sheet.iter_rows(min_row=2, values_only=True):

    name, phone = row
    search_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="3"]')
    search_box.clear()
    search_box.send_keys(f'{phone}')
    driver.implicitly_wait(5)
    time.sleep(1)
    try:
        driver.find_element_by_xpath(f'//span[@title="{name}"]')
    except:
        pyautogui.press('backspace', presses=10)
        print(f"{name}'s chat not found, skipping...")
        continue

    search_box.send_keys(Keys.RETURN)
    time.sleep(1)

    message_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="10"]')
    message_box.send_keys(f'{message}')
    driver.implicitly_wait(5)
    message_box.send_keys(Keys.RETURN)

    attach_button = driver.find_element_by_xpath('//div[@title="Attach"]')
    attach_button.click()
    time.sleep(1)
    image_input = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_input.send_keys(image_path)
    time.sleep(3)
    send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
    send_button.click()
    time.sleep(1)

    attach_button = driver.find_element_by_xpath('//div[@title="Attach"]')
    attach_button.click()
    time.sleep(1)
    image_input = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_input.send_keys(image_path_2)
    time.sleep(3)
    send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
    send_button.click()
    time.sleep(1)

    attach_button = driver.find_element_by_xpath('//div[@title="Attach"]')
    attach_button.click()
    time.sleep(1)
    image_input = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_input.send_keys(image_path_3)
    time.sleep(3)
    send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
    send_button.click()
    time.sleep(1)

    search_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="3"]')
    search_box.clear()
    time.sleep(1)

print('Messages sent successfully!')
input()
driver.quit()
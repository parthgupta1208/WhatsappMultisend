from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import openpyxl
import time

# Load the Excel file and select the sheet with the contacts
workbook = openpyxl.load_workbook('contact.xlsx')
sheet = workbook['Sheet1']

# Get the WhatsApp message to send
message = "II PUC regular classes with CET and NEET coaching will start from April 1st week.. Faculty from reputed Institutions will be handling classes.. For more information contact Insight Academy Santhekatte -9880782393"
image_path="C:\Everything\hehe.jpg"
image_path_2="C:\Everything\hehe2.jpg"
image_path_3="C:\Everything\hehe3.jpg"

# Set up the WebDriver
driver = webdriver.Chrome('C:/Everything/chromedriver.exe')
driver.get('https://web.whatsapp.com/')

# Wait for the user to scan the QR code
input('Scan the QR code and press enter to continue...')

# Iterate over the rows in the sheet and send the message to each contact
for row in sheet.iter_rows(min_row=2, values_only=True):
    # Extract the contact information
    name, phone = row

    # Search for the contact and open the chat
    search_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="3"]')
    search_box.clear()
    search_box.send_keys(f'{phone}')
    driver.implicitly_wait(5) # Wait for the contact to load
    search_box.send_keys(Keys.RETURN)
    time.sleep(1)

    # Send the message
    message_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="10"]')
    message_box.send_keys(f'{message}')
    driver.implicitly_wait(5) # Wait for the message box to load
    message_box.send_keys(Keys.RETURN)
    time.sleep(1)

    attach_button = driver.find_element_by_xpath('//div[@title="Attach"]')
    attach_button.click()
    time.sleep(1) # Wait for the dialog box to open
    image_input = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_input.send_keys(image_path)
    time.sleep(3)
    send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
    send_button.click()
    time.sleep(1)

    attach_button = driver.find_element_by_xpath('//div[@title="Attach"]')
    attach_button.click()
    time.sleep(1) # Wait for the dialog box to open
    image_input = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_input.send_keys(image_path_2)
    time.sleep(3)
    send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
    send_button.click()
    time.sleep(1)

    attach_button = driver.find_element_by_xpath('//div[@title="Attach"]')
    attach_button.click()
    time.sleep(1) # Wait for the dialog box to open
    image_input = driver.find_element_by_xpath('//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
    image_input.send_keys(image_path_3)
    time.sleep(3)
    send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
    send_button.click()
    time.sleep(1)

    # Clear the search box to prepare for the next contact
    search_box = driver.find_element_by_xpath('//div[@contenteditable="true"][@data-tab="3"]')
    search_box.clear()
    time.sleep(1) # Wait for the search box to clear

print('Messages sent successfully!')
input()
# Close the WebDriver
driver.quit()
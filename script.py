# Program to send bulk messages through WhatsApp web from an excel sheet without saving contact numbers
# Author @inforkgodara

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import re

excel_data = pandas.read_excel('Recipients data.xlsx', sheet_name='Recipients')

count = 0

options = Options()
options.add_argument('--no-sandbox') 
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

driver.get('https://web.whatsapp.com')
input("Press ENTER after login into Whatsapp Web and your chats are visiable.")
for index, column in enumerate(excel_data['Contact'].tolist()):
    try:
        contactStr = excel_data['Contact'][count]
        if (isfloat(contactStr)):
            contactStr = str(int(contactStr))
        contactStr = re.sub('[^0-9]','', contactStr)
        if (contactStr.startswith('0')):
            contactStr = contactStr.replace('0','+972',1)
        if (contactStr.startswith('972')):
            contactStr = "+" + contactStr
        if (contactStr.startswith('5')):
            contactStr = '+972' + contactStr
        url = 'https://web.whatsapp.com/send?phone=' + str(contactStr) + '&text=' + excel_data['Message'][index]
        sent = False
        # It tries 3 times to send a message in case if there any error occurred
        driver.get(url)
        try:
            click_btn_child = WebDriverWait(driver, 35).until(EC.visibility_of_element_located((By.XPATH, '//span[@data-testid="send"]')))
            click_btn = click_btn_child.find_element("xpath", ('..'))
        except Exception as e:
            print("Sorry message could not sent to " + str(contactStr))
        else:
            sleep(2)
            click_btn.click()
            sent = True
            sleep(5)
            print('Message sent to: ' + str(contactStr))
        count = count + 1
    except Exception as e:
        print('Failed to send message to ' + str(contactStr) + str(e))
driver.quit()
print("The script executed successfully.")
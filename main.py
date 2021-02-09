import os #This is for the find function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


def find(name, path): # This is so the program can search for where the webdriver file is on the user's computer
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)
        #else:
         #   raise Exception("chromedriver.exe could not be found")


# driver = webdriver.Chrome('C:/Users/EvanS/Downloads/chromedriver_win32/chromedriver.exe')
# This ^^^ is so I don't forget where I have the webdriver on my computer


driver = webdriver.Chrome(find("chromedriver.exe", "C:/Users"))

# Reads in my login credentials. This will have to be changed later.
with open("C:/Users/EvanS/Documents/College/Capstone/credentials.txt") as f:
    name = f.readline()
    word = f.readline()
    f.close()

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

driver.get('https://bb.nsuok.edu')
driver.maximize_window()

driver.find_element_by_id('username').send_keys(name)
driver.find_element_by_id('password').send_keys(word)
driver.find_element_by_name('submit').click()

try:
    WebDriverWait(driver, .001).until(EC.url_matches('https://bb.nsuok.edu/ultra'))
except TimeoutError:
    print('took too  long')

driver.implicitly_wait(20)
driver.find_element_by_link_text("Courses").click()
WebDriverWait(driver, 10).until(EC.url_matches('https://bb.nsuok.edu/ultra/course'))
driver.implicitly_wait(60)

driver.find_element_by_link_text("Bekkering-Capstone-Sandbox").click()

#driver.implicitly_wait(60)
#driver.find_element_by_link_text("Content").click()

driver.implicitly_wait(60)

element = WebDriverWait(driver, .001).until(EC.presence_of_element_located((By.XPATH, '//*[@id="paletteItem:_1135693_1"]/a')))
element.click()

#driver.find_element_by_xpath('//*[@id="paletteItem:_1135693_1"]/a').click()




import os #This is for the find function
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


def find(name, path): # This is so the program can search for where the webdriver file is on the user's computer
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)


# driver = webdriver.Chrome('C:/Users/EvanS/Downloads/chromedriver_win32/chromedriver.exe')
# This ^^^ is so I don't forget where I have the webdriver on my computer


driver = webdriver.Chrome(find("chromedriver.exe", "C:/Users"))


#driver.get("https://www.nsuok.edu/MyNSU/default.aspx")


#driver.find_element_by_link_text("BLACKBOARD").click()

# Reads in my login credentials. This will have to be changed later.
with open("C:/Users/EvanS/Documents/College/Capstone/credentials.txt") as f:
    name = f.readline()
    word = f.readline()
    f.close()


#from selenium import webdriver
#from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.common.keys import Keys
#from credentials import bb_password, bb_username
#browser = webdriver.Chrome()
driver.get('https://bb.nsuok.edu')
driver.maximize_window()
driver.find_element_by_id('username').send_keys(name)
driver.find_element_by_id('password').send_keys(word)
driver.find_element_by_name('submit').click()
try:
    WebDriverWait(driver, .001).until(EC.url_matches('https://bb.nsuok.edu/ultra'))
except TimeoutError:
    print('took too  long')

#WebDriverWait(driver, 20)
driver.implicitly_wait(20)
#driver.find_element_by_name('Courses').click()
driver.find_element_by_link_text("Courses").click()
WebDriverWait(driver, 10).until(EC.url_matches('https://bb.nsuok.edu/ultra/course'))
driver.implicitly_wait(60)

driver.find_element_by_link_text("Bekkering-Capstone-Sandbox").click()


#driver.find_element_by_xpath('//*[@id="course-link-_51855_1"]').click()
# ^^ This line works every time
#WebDriverWait(driver, 10).until(EC.presence_of_element_located('BEKKERING-CAPSTONE-SANDBOX'))
#driver.find_element_by_name('Organizations').click()
#WebDriverWait(driver, 10).until(EC.url_matches('https://bb.nsuok.edu/ultra/logout'))

# driver.implicitly_wait(30)

#card = driver.find_element_by_class_name("card-body")

#usernameField_Xpath = '//*[@id="username"]'



#userNameField = card.find_element_by_xpath(usernameField_Xpath)
#.send_keys(
# username)
#driver.find_element_by_id("password").send_keys(password)

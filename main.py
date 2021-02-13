import os  # This is for the find function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
import excelTest


def find(name, path):  # This is so the program can search for where the webdriver file is on the user's computer
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)


def nav(  # reminder to future me to pass input here to determine which option was chosen and thus which function to
        # call next from nav, taking care of the problem of driver being a local variable
         ):
    driver = webdriver.Chrome(find("chromedriver.exe", "C:/Users"))
    # driver = webdriver.Chrome('C:/Users/EvanS/Downloads/chromedriver_win32/chromedriver.exe')
    # This ^^^ is so I don't forget where I have the webdriver on my computer

    # Reads in my login credentials. This will have to be changed later.
    with open("C:/Users/EvanS/Documents/College/Capstone/credentials.txt") as f:
        name = f.readline()
        word = f.readline()
        f.close()

    driver.get('https://bb.nsuok.edu')
    driver.maximize_window()

    driver.find_element_by_id('username').send_keys(name)
    driver.find_element_by_id('password').send_keys(word)
    driver.find_element_by_name('submit').click()

    try:
        WebDriverWait(driver, .001).until(ec.url_matches('https://bb.nsuok.edu/ultra'))
    except TimeoutError:
        print('took too  long')

    driver.implicitly_wait(20)
    driver.find_element_by_link_text("Courses").click()
    WebDriverWait(driver, 10).until(ec.url_matches('https://bb.nsuok.edu/ultra/course'))
    driver.implicitly_wait(60)

    driver.find_element_by_link_text("Bekkering-Capstone-Sandbox").click()

    # driver.implicitly_wait(60)
    # driver.find_element_by_link_text("Content").click()

    driver.implicitly_wait(60)

    element = WebDriverWait(driver, .001).until(ec.presence_of_element_located((By.XPATH, '//*[@id="paletteItem:_1135693_1"]/a')))
    element.click()

    # driver.find_element_by_xpath('//*[@id="paletteItem:_1135693_1"]/a').click()


def main():
    option = ""
    while True:
        print("1: Create Test File\n")
        print("2: Upload Test\n")
        print("3: Upload Test Pool\n")
        print("4: Create Test From Existing Pools\n")
        print("q: Quit")
        option = input("Select Option:")

        if option == "1":
            excelTest.create_file()
        if option == "2":
            nav()
        if option == "3":
            nav()
        if option == "4":
            nav()
        if option == "q":
            break


if __name__ == "__main__":
    main()






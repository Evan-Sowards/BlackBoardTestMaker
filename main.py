import os  # This is for the find function
import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
import excelTest
import question


def find(name, path):  # This is so the program can search for where the webdriver file is on the user's computer
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)


def nav(num):
    # username = input("NSU Username: ")
    # I'll uncomment these for the final version, but for I'm not entering my credentials over and over and over again.
    # password = input("NSU Password: ")

    # filepath = input("Test Filepath (USE / NOT \): ")
    filepath = "C:/Users/EvanS/Desktop/test.xlsx"

    driver = webdriver.Chrome(find("chromedriver.exe", "C:/Users"))
    # driver = webdriver.Chrome('C:/Users/EvanS/Downloads/chromedriver_win32/chromedriver.exe')
    # This ^^^ is so I don't forget where I have the webdriver on my computer

    # Reads in my login credentials. This will have to be changed later.
    with open("C:/Users/EvanS/Documents/College/Capstone/credentials.txt") as f:
        username = f.readline()
        password = f.readline()
        f.close()

    driver.get('https://bb.nsuok.edu')
    driver.maximize_window()

    driver.find_element_by_id('username').send_keys(username)
    driver.find_element_by_id('password').send_keys(password)
    driver.find_element_by_name('submit').click()

    username = ""  # Clears the username and password immediately after usage
    password = ""

    try:
        WebDriverWait(driver, .001).until(ec.url_matches('https://bb.nsuok.edu/ultra'))
    except TimeoutError:
        print('took too  long')

    driver.implicitly_wait(20)
    # driver.find_element_by_link_text("Courses").click()
    # WebDriverWait(driver, 10).until(ec.url_matches('https://bb.nsuok.edu/ultra/course'))
    driver.get('https://bb.nsuok.edu/ultra/course')
    driver.implicitly_wait(60)

    course_name = excelTest.getCourseName(filepath)


    driver.find_element_by_link_text(course_name).click()

    if num == 2:
        createTest(driver, filepath)
    if num == 3:
        createPool(driver, filepath)


def createTest(driver, filepath):

    driver.get('https://bb.nsuok.edu/webapps/blackboard/content/listContentEditable.jsp?content_id=_4169187_1&course_id=_51855_1&mode=reset')

    driver.implicitly_wait(20)

    # driver.find_element_by_xpath('//*[@id="evaMenu_actionButton"]/option[text("Test")="option_text"]').click()
    # driver.find_element_by_xpath('//*[@id="content-handler-resource/x-bb-asmt-test-link"]').click()
    # select.select_by_visible_text('Test')

    driver.get('https://bb.nsuok.edu/webapps/assessment/do/content/assessment?action=ADD&course_id=_51855_1&content_id=_4169187_1&assessmentType=Test')

    driver.implicitly_wait(20)

    driver.find_element_by_xpath('//*[@id="stepcontent1"]/ol/li[2]/div[2]/a').click()

    driver.find_element_by_xpath('//*[@id="assessment_name_input"]').send_keys(excelTest.getTestName(filepath))

    driver.implicitly_wait(20)

    driver.find_element_by_xpath('//*[@id="bottom_Submit"]').click()


def createPool(driver, filepath):

    driver.get('https://bb.nsuok.edu/webapps/blackboard/landingPage.jsp?navItem=cp_test_survey_pool&course_id=_51855_1&sortItems=false')

    driver.implicitly_wait(20)

    driver.get('https://bb.nsuok.edu/webapps/assessment/do/authoring/viewAssessmentManager?assessmentType=Pool&course_id=_51855_1')

    driver.implicitly_wait(20)

    driver.find_element_by_xpath('//*[@id="nav"]/li[1]/a').click()

    driver.find_element_by_xpath('//*[@id="assessment_name_input"]').send_keys(excelTest.getTestName(filepath))

    driver.implicitly_wait(20)

    driver.find_element_by_xpath('//*[@id="bottom_Submit"]').click()

def main():
    option = ""
    pool = list()

    while True:
        print("1: Create Test File\n")
        print("2: Upload Test\n")
        print("3: Upload Test Pool\n")
        print("4: Create Test From Existing Pools\n")
        print("5: Reading questions temporary test function\n")
        print("q: Quit")
        option = input("Select Option:")

        if option == "1":
            excelTest.create_file()
        if option == "2":
            nav(2)
        if option == "3":
            nav(3)
        if option == "4":
            nav(4)
        if option == "5":
            excelTest.load_questions(pool)
            for x in range(len(pool)):
                pool[x].print_question()

        if option == "q":
            sys.exit()


if __name__ == "__main__":
    main()






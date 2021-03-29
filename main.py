import os  # This is for the find function
import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
import excelTest
import question
import time


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

    # note: 5 works well, use variable
    try:
        WebDriverWait(driver, 5).until(ec.url_matches('https://bb.nsuok.edu/ultra'))
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

    # Putting in questions

    pool = list()

    excelTest.load_questions(pool, filepath)

    size = len(pool)

    actions = ActionChains(driver)

    for i in range(size):
        if pool[i].num_right == 0:
            driver.find_element_by_xpath('//*[@id="nav"]/li[1]').click()
            driver.find_element_by_link_text('Short Answer').click()
            iframe = driver.find_element_by_xpath('//*[@id="questionText.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].text)
            driver.switch_to.parent_frame()
            # Setting question text

            html = driver.find_element_by_tag_name('html')
            for k in range(20):
                html.send_keys(Keys.ARROW_DOWN)

            iframe = driver.find_element_by_xpath('//*[@id="answerText.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].correct)
            driver.switch_to.parent_frame()
            # Setting Answer Feedback

            driver.find_element_by_xpath('//*[@id="bottom_submitButtonRow"]/input[3]').click()

        elif pool[i].num_right == 1:
            driver.find_element_by_xpath('//*[@id="nav"]/li[1]').click()
            driver.find_element_by_link_text('Multiple Choice').click()

            total = pool[i].num_right + pool[i].num_wrong
            total = total - 3
            num_string = '//*[@id="answerList.answerCountId"]/option[' + str(total) + ']'
            driver.find_element_by_xpath(num_string).click()
            # Setting number of question options
            # time.sleep(3)

            iframe = driver.find_element_by_xpath('//*[@id="questionText.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].text)
            driver.switch_to.parent_frame()
            # Setting question text

            driver.find_element_by_xpath('//*[@id="numberTypeAsStringId"]/option[5]').click()
            # option for how questions are labeled (a, b, c; 1, 2, 3; j, ii, iii; etc.)

            driver.find_element_by_xpath('//*[@id="randomOrderId"]').click()
            # option for randomization of answer order

            for j in range(total + 3):
                html = driver.find_element_by_tag_name('html')
                for k in range(11):
                    html.send_keys(Keys.ARROW_DOWN)

                iframe = driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].textForm.text_ifr"]')
                driver.switch_to.frame(iframe)
                element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
                element.click()

                if j == 0:
                    element.send_keys(pool[i].r1)
                if j == 1:
                    element.send_keys(pool[i].w1)
                if j == 2:
                    element.send_keys(pool[i].w2)
                if j == 3:
                    element.send_keys(pool[i].w3)
                if j == 4:
                    element.send_keys(pool[i].w4)
                if j == 5:
                    element.send_keys(pool[i].w5)
                if j == 6:
                    element.send_keys(pool[i].w6)
                if j == 7:
                    element.send_keys(pool[i].w7)
                if j == 8:
                    element.send_keys(pool[i].w8)

                driver.switch_to.parent_frame()
                # Putting answer choices

            html = driver.find_element_by_tag_name('html')
            for k in range(15):
                html.send_keys(Keys.ARROW_DOWN)
            iframe = driver.find_element_by_xpath('//*[@id="correctFeedbackForm.textForm.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].correct)
            driver.switch_to.parent_frame()
            # Correct Feedback

            iframe = driver.find_element_by_xpath('//*[@id="incorrectFeedbackForm.textForm.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].incorrect)
            driver.switch_to.parent_frame()
            # Incorrect Feedback

            driver.find_element_by_xpath('//*[@id="bottom_submitButtonRow"]/input[3]').click()

        else:
            driver.find_element_by_xpath('//*[@id="nav"]/li[1]').click()
            driver.find_element_by_link_text('Multiple Answer').click()

            total = pool[i].num_right + pool[i].num_wrong
            total = total - 3
            num_string = '//*[@id="fAnswerListAnswerCount"]/option[' + str(total) + ']'
            driver.find_element_by_xpath(num_string).click()
            # Setting number of question options

            driver.find_element_by_xpath('//*[@id="fNumberTypeAsString"]/option[5]').click()
            # option for how questions are labeled (a, b, c; 1, 2, 3; j, ii, iii; etc.)

            iframe = driver.find_element_by_xpath('//*[@id="questionText.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].text)
            driver.switch_to.parent_frame()
            # Setting question text

            # driver.find_element_by_xpath('//*[@id="fPartialCredit"]').click()
            # Option for partial credit; this will be off by default, but can be enabled from the options file
            # Also, reminder to myself that I need to make the options file

            html = driver.find_element_by_tag_name('html')
            for k in range(12):
                html.send_keys(Keys.ARROW_DOWN)
            time.sleep(1)

            if not driver.find_element_by_xpath('//*[@id="fRandomOrder"]').get_attribute('checked'):
                driver.find_element_by_xpath('//*[@id="fRandomOrder"]').click()
                # option for randomization of answer order
                # time.sleep(5)

            right = pool[i].num_right
            wrong = pool[i].num_wrong

            for j in range(total + 3):
                html = driver.find_element_by_tag_name('html')
                for k in range(10):
                    html.send_keys(Keys.ARROW_DOWN)

                iframe = driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].textForm.text_ifr"]')
                driver.switch_to.frame(iframe)
                element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
                element.click()

                if pool[i].r1 != "None":
                    element.send_keys(pool[i].r1)
                    driver.switch_to.parent_frame()
                    if not driver.find_element_by_xpath(
                            '//*[@id="answerList.answer[' + str(j) + '].correct"]').get_attribute('checked'):
                        driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r1 = "None"
                    continue

                if pool[i].r2 != "None":
                    element.send_keys(pool[i].r2)
                    driver.switch_to.parent_frame()
                    driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r2 = "None"
                    continue

                if pool[i].r3 != "None":
                    element.send_keys(pool[i].r3)
                    driver.switch_to.parent_frame()
                    driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r3 = "None"
                    continue

                if pool[i].r4 != "None":
                    element.send_keys(pool[i].r1)
                    driver.switch_to.parent_frame()
                    driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r4 = "None"
                    continue

                if pool[i].r5 != "None":
                    element.send_keys(pool[i].r5)
                    driver.switch_to.parent_frame()
                    driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r5 = "None"
                    continue

                if pool[i].r6 != "None":
                    element.send_keys(pool[i].r6)
                    driver.switch_to.parent_frame()
                    driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r6 = "None"
                    continue

                if pool[i].r7 != "None":
                    element.send_keys(pool[i].r7)
                    driver.switch_to.parent_frame()
                    driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r7 = "None"
                    continue

                if pool[i].r8 != "None":
                    element.send_keys(pool[i].r8)
                    driver.switch_to.parent_frame()
                    driver.find_element_by_xpath('//*[@id="answerList.answer[' + str(j) + '].correct"]').click()
                    pool[i].r8 = "None"
                    continue

                if pool[i].w1 != "None":
                    element.send_keys(pool[i].w1)
                    driver.switch_to.parent_frame()
                    pool[i].w1 = "None"
                    continue

                if pool[i].w2 != "None":
                    element.send_keys(pool[i].w2)
                    driver.switch_to.parent_frame()
                    pool[i].w2 = "None"
                    continue

                if pool[i].w3 != "None":
                    element.send_keys(pool[i].w3)
                    driver.switch_to.parent_frame()
                    pool[i].w3 = "None"
                    continue

                if pool[i].w4 != "None":
                    element.send_keys(pool[i].w4)
                    driver.switch_to.parent_frame()
                    pool[i].w4 = "None"
                    continue

                if pool[i].w5 != "None":
                    element.send_keys(pool[i].w5)
                    driver.switch_to.parent_frame()
                    pool[i].w5 = "None"
                    continue

                if pool[i].w6 != "None":
                    element.send_keys(pool[i].w6)
                    driver.switch_to.parent_frame()
                    pool[i].w6 = "None"
                    continue

                if pool[i].w7 != "None":
                    element.send_keys(pool[i].w7)
                    driver.switch_to.parent_frame()
                    pool[i].w7 = "None"
                    continue

                if pool[i].w8 != "None":
                    element.send_keys(pool[i].w8)
                    driver.switch_to.parent_frame()
                    pool[i].w8 = "None"
                    continue
                # Entering Answers

            html = driver.find_element_by_tag_name('html')
            for k in range(15):
                html.send_keys(Keys.ARROW_DOWN)
            iframe = driver.find_element_by_xpath('//*[@id="correctFeedbackForm.textForm.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].correct)
            driver.switch_to.parent_frame()
            # Correct Feedback

            iframe = driver.find_element_by_xpath('//*[@id="incorrectFeedbackForm.textForm.text_ifr"]')
            driver.switch_to.frame(iframe)
            element = driver.find_element_by_xpath('//*[@id="tinymce"]/p')
            element.click()
            element.send_keys(pool[i].incorrect)
            driver.switch_to.parent_frame()
            # Incorrect Feedback

            driver.find_element_by_xpath('//*[@id="bottom_submitButtonRow"]/input[3]').click()


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

    while option != "q":
        print("1: Create Test File\n")
        print("2: Upload Test\n")
        print("3: Upload Test Pool\n")
        print("q: Quit")
        option = input("Select Option:")

        if option == "1":
            excelTest.create_file()
        if option == "2":
            nav(2)
        if option == "3":
            nav(3)
        if option == "5":
           excelTest.load_questions(pool, "C:/Users/EvanS/Desktop/test.xlsx")
           for x in range(len(pool)):
                pool[x].print_question()


if __name__ == "__main__":
    main()






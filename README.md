# Blackboard Automatic Test Uploader
## Table of Contents
1. [Introduction](#Introduction)
2. [Requirements](#Requirements)
3. [Set Up](#Set-up)
4. [Instructions](#Instructions)
5. [Options](#Options)
6. [License](#License)

### Introduction

Blackboard Automatic Test Uploader is a Python program on Windows computers that uses Selenium WebDriver to quickly upload tests and test pools to Blackboard. It was made for professors at my school.


### Requirements

There are a few things that need to be installed before you can run Blackboard Automatic Test Uploader:

* [Selenium WebDriver](https://www.geeksforgeeks.org/how-to-install-selenium-in-python/)
* [Chromedriver](https://sites.google.com/a/chromium.org/chromedriver/downloads)
* [Ruby](https://rubyinstaller.org/)
* [Python](https://www.python.org/downloads/)
* [Openpyxl](https://openpyxl.readthedocs.io/en/stable/#installation)

### Set up

You'll need some way of running python files, such as [VS Code](https://code.visualstudio.com/download), [PyCharm](https://www.jetbrains.com/pycharm/download/#section=windows), or the Windows Command Prompt. You'll need to make sure you have the chromedriver.exe file somewhere in the C://Users path (e.g., your Downloads folder). Once you have all of the required software listed in the Requirements section, download the code for Blackboard Automatic Test Uploader and run bbAutoTestMaker.py.


### Instructions

When you run bbAutoTestMaker.py, you will see a menu:

![menu](https://github.com/Evan-Sowards/BlackBoardTestMaker/blob/a34343ba2b69f01b7c2d528aeab5686d572e327b/pics/menu.PNG?raw=true)



#### 1: Create a file

This creates a blank, pre-formatted .xlsx file for you to fill in. You will be asked to name the file. You do not need to include the ".xlsx" extension to your file name.

Course Name and the name of the test must be listed in the excel test file. The course name must match exactly with the corresponding course on Blackboard. These must be on row 2.

Each row holds the data for a question. If the "Question Text" cell is blank for a row, the program will not read any questions after that row.

There are three types of questions that the program supports: multiple choice, multiple answer (select all that apply), and short answer. There are 8 "Right Answer" columns and 8 "Wrong Answer" columns. If there is only one right answer on a row, then the program will list that as a multiple choice question. If there is more than one right answer on a row, then the program will list that as a multiple answer question. If there are no right answers listed, then the program will list that as a short answer question.

Multiple choice and multiple answer questions need to have a minimum of 4 possible answers. For a multiple choice question, this means one right answer and at least 3 wrong answers. For multiple answer questions there needs to be at least two right answers and a total of 4 possible answers. (So if there's two right answers, there needs to be 2 wrong answers. If there's 4 right answers, you can anything from 0 to 8 wrong answers).

For short answer questions, question feedback needs to be put under "Correct Feedback".

#### 2 and 3: Upload a test/pool

Both of these will ask you to do the same thing. The only difference is option 2 creates a test, while option 3 creates a pool of questions.

You will first be asked to enter your NSU username and password. Then you will be asked to enter the name of the file that you want to upload as a test (or pool). You WILL need to include the .xlsx extension at the end of your filename. After that, the program will look for all of the necessary files (which might take a moment), open Chrome, and then begin the process of creating a test (or pool) on Blackboard.


### Options

There are three options that can be changed in the included options.xlsx file. If a number that doesn't correspond to a valid setting is entered (e.g., 9 for answer numbering) the program will reset the options. The defaults (in order) are 5, 1, 0.

![options](https://github.com/Evan-Sowards/BlackBoardTestMaker/blob/34918a85628fa41bb5fffd920a4e1bbf0e50902c/pics/options.PNG?raw=true)

#### Answer Numbering:

1. None
2. Arabic Numerals (1, 2, 3)
3. Roman Numberals (I, II, III)
4. Uppercase Letters (A, B, C)
5. Lowercase Letters (a, b, c)

#### Randomness:

  **0:** Off
  
  **1:** On

#### Partial Credit:

  **0:** Off

  **1:** On


**NOTE:** Randomness must be turned on in the options file, or the right answer will always be at the top.


### License

This project is licensed under the MIT License. See the LICENSE.md file for details.

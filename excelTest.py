import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
import question


def create_file():

    wb = openpyxl.Workbook()
    filepath = "C:/Users/Public/Desktop/test.xlsx"
    wb.save(filepath)

    wb1 = openpyxl.load_workbook(filepath)

    mySheet = wb1['Sheet']

    mySheet['A1'] = 'Course Name'
    mySheet['B1'] = 'Test Name'
    mySheet['C1'] = 'Question Text'
    mySheet['D1'] = 'Right Answer 1'
    mySheet['E1'] = 'Right Answer 2'
    mySheet['F1'] = 'Right Answer 3'
    mySheet['G1'] = 'Right Answer 4'
    mySheet['H1'] = 'Right Answer 5'
    mySheet['I1'] = 'Right Answer 6'
    mySheet['J1'] = 'Right Answer 7'
    mySheet['K1'] = 'Right Answer 8'
    mySheet['L1'] = 'Wrong Answer 1'
    mySheet['M1'] = 'Wrong Answer 2'
    mySheet['N1'] = 'Wrong Answer 3'
    mySheet['O1'] = 'Wrong Answer 4'
    mySheet['P1'] = 'Wrong Answer 5'
    mySheet['Q1'] = 'Wrong Answer 6'
    mySheet['R1'] = 'Wrong Answer 7'
    mySheet['S1'] = 'Wrong Answer 8'
    mySheet['T1'] = 'Correct Feedback'
    mySheet['U1'] = 'Incorrect Feedback'

    mySheet.column_dimensions['A'].width = 18
    mySheet.column_dimensions['B'].width = 16
    mySheet.column_dimensions['C'].width = 19
    mySheet.column_dimensions['D'].width = 19
    mySheet.column_dimensions['E'].width = 19
    mySheet.column_dimensions['F'].width = 19
    mySheet.column_dimensions['G'].width = 19
    mySheet.column_dimensions['H'].width = 19
    mySheet.column_dimensions['I'].width = 19
    mySheet.column_dimensions['J'].width = 19
    mySheet.column_dimensions['K'].width = 19
    mySheet.column_dimensions['L'].width = 20
    mySheet.column_dimensions['M'].width = 20
    mySheet.column_dimensions['N'].width = 20
    mySheet.column_dimensions['O'].width = 20
    mySheet.column_dimensions['P'].width = 20
    mySheet.column_dimensions['Q'].width = 20
    mySheet.column_dimensions['R'].width = 20
    mySheet.column_dimensions['S'].width = 20
    mySheet.column_dimensions['T'].width = 22
    mySheet.column_dimensions['U'].width = 24

    thin = Side(border_style="thick", color="303030")
    letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    for x in letters:
        column = x + "1"
        mySheet[column].border = Border(top=thin, left=thin, right=thin, bottom=thin)
        mySheet[column].font = Font(size=14, bold=True)
        mySheet[column].alignment = Alignment(horizontal="center")
        mySheet[column].fill = PatternFill(start_color="00ffcc", fill_type="solid")
        # Thanks to user "crifan" on Stack Overflow for explaining how to use PatternFill
        # https://stackoverflow.com/questions/35918504/adding-a-background-color-to-cell-openpyxl

    wb1.save(filepath)


def load_questions(pool, filepath):

    wb1 = openpyxl.load_workbook(filepath)

    mySheet = wb1['Sheet']
    letters = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
   # i = 2
   # end = 0
   # while True:
   #     for x in letters:
   #         cell = x + str(i)
   #         if str(mySheet[cell].value) == "None":
   #             end = 1
   #             break
   #         else:
   #             question.text = str(mySheet[cell].value)
   #             question.r1 = str(mySheet[cell].value)
   #             # pool.append(question)

    #    if end == 1:
    #        break

     #   else:
      #      i = i + 1

    i = 2
    end = 0
    while True:
        cell = "C" + str(i)
        if str(mySheet[cell].value) == "None":
            end = 1
            break
        else:
            something = question.Question()
            something.text = str(mySheet[cell].value)
            cell = "D" + str(i)
            something.r1 = str(mySheet[cell].value)
            cell = "E" + str(i)
            something.r2 = str(mySheet[cell].value)
            cell = "F" + str(i)
            something.r3 = str(mySheet[cell].value)
            cell = "G" + str(i)
            something.r4 = str(mySheet[cell].value)
            cell = "H" + str(i)
            something.r5 = str(mySheet[cell].value)
            cell = "I" + str(i)
            something.r6 = str(mySheet[cell].value)
            cell = "J" + str(i)
            something.r7 = str(mySheet[cell].value)
            cell = "K" + str(i)
            something.r8 = str(mySheet[cell].value)
            cell = "L" + str(i)
            something.w1 = str(mySheet[cell].value)
            cell = "M" + str(i)
            something.w2 = str(mySheet[cell].value)
            cell = "N" + str(i)
            something.w3 = str(mySheet[cell].value)
            cell = "O" + str(i)
            something.w4 = str(mySheet[cell].value)
            cell = "P" + str(i)
            something.w5 = str(mySheet[cell].value)
            cell = "Q" + str(i)
            something.w6 = str(mySheet[cell].value)
            cell = "R" + str(i)
            something.w7 = str(mySheet[cell].value)
            cell = "S" + str(i)
            something.w8 = str(mySheet[cell].value)
            cell = "T" + str(i)
            something.correct = str(mySheet[cell].value)
            cell = "U" + str(i)
            something.incorrect = str(mySheet[cell].value)
            something.num_right = 0
            if something.r1 != "None":
                something.num_right += 1
            if something.r2 != "None":
                something.num_right += 1
            if something.r3 != "None":
                something.num_right += 1
            if something.r4 != "None":
                something.num_right += 1
            if something.r5 != "None":
                something.num_right += 1
            if something.r6 != "None":
                something.num_right += 1
            if something.r7 != "None":
                something.num_right += 1
            if something.r8 != "None":
                something.num_right += 1

            something.num_wrong = 0
            if something.w1 != "None":
                something.num_wrong += 1
            if something.w2 != "None":
                something.num_wrong += 1
            if something.w3 != "None":
                something.num_wrong += 1
            if something.w4 != "None":
                something.num_wrong += 1
            if something.w5 != "None":
                something.num_wrong += 1
            if something.w6 != "None":
                something.num_wrong += 1
            if something.w7 != "None":
                something.num_wrong += 1
            if something.w8 != "None":
                something.num_wrong += 1

            pool.append(something)

        if end == 1:
            break
        else:
            i = i + 1


    # question.text = str(mySheet['C3'].value)
    # question.r1 = str(mySheet['D3'].value)
    #pool.append(question)
    #print("Hello world\n")

def getCourseName(filepath):
    wb1 = openpyxl.load_workbook(filepath)

    mySheet = wb1['Sheet']

    return str(mySheet['A2'].value)


def getTestName(filepath):
    wb1 = openpyxl.load_workbook(filepath)

    mySheet = wb1['Sheet']

    return str(mySheet['B2'].value)


import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter


def create_file():

    wb = openpyxl.Workbook()
    filepath = "C:/Users/EvanS/Desktop/test.xlsx"
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
        # Thanks to user "crifan" on Stack Overflow for the right side of the assignment operator on the above line
        # https://stackoverflow.com/questions/35918504/adding-a-background-color-to-cell-openpyxl

    wb1.save(filepath)

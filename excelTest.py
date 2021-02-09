import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, cell_style, Color
from openpyxl.utils import get_column_letter
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

mySheet.column_dimensions['A'].width = 12
mySheet.column_dimensions['B'].width = 10
mySheet.column_dimensions['C'].width = 13
mySheet.column_dimensions['D'].width = 13
mySheet.column_dimensions['E'].width = 13
mySheet.column_dimensions['F'].width = 13
mySheet.column_dimensions['G'].width = 13
mySheet.column_dimensions['H'].width = 13
mySheet.column_dimensions['I'].width = 13
mySheet.column_dimensions['J'].width = 13
mySheet.column_dimensions['K'].width = 13
mySheet.column_dimensions['L'].width = 14
mySheet.column_dimensions['M'].width = 14
mySheet.column_dimensions['N'].width = 14
mySheet.column_dimensions['O'].width = 14
mySheet.column_dimensions['P'].width = 14
mySheet.column_dimensions['Q'].width = 14
mySheet.column_dimensions['R'].width = 14
mySheet.column_dimensions['S'].width = 14
mySheet.column_dimensions['T'].width = 16
mySheet.column_dimensions['U'].width = 18

mySheet['A1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['B1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['C1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['D1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['E1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['F1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['G1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['H1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['I1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['J1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['K1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['L1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['M1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['N1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['O1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['P1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['Q1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['R1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['S1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['T1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
mySheet['U1'].fill = PatternFill(start_color="ff0000", fill_type="solid")
# Thanks to user "crifan" on Stack Overflow for the right side of the equals sign of these ^ lines of code
# https://stackoverflow.com/questions/35918504/adding-a-background-color-to-cell-openpyxl


wb1.save(filepath)






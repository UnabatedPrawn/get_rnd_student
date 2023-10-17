import random
import openpyxl

wb = openpyxl.load_workbook('rnd student selection.xlsx')
def get_grp(sheets):
   d = list(sheets)
   print('Who are we _t e r r o r i z i n g_ today?')
   for index, value in enumerate(d):
       print(index, value)
       x = int(input())
       return x

def get_student(var):
    name = list(var)
    return random.choice(name)

def what_do():
    print('What next?')
    print('Options:', '1: re-roll', '2: exit')
    x = int(input())
    return x

sheets = [wb.sheetnames]
grp = get_grp(sheets)
wb._active_sheet_index = grp
sheet = wb.active

students = []
for col in sheet['A']:
    students.append(col.value)

print('.')
print('..')
print('...')
print(get_student(students))
print('...')
print('..')
print('.')

x = what_do()
while x == 1:
    print('.')
    print('..')
    print('...')
    print(get_student(students))
    print('...')
    print('..')
    print('.')
    x = what_do()
if x == 2:
    exit()
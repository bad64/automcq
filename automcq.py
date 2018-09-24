from random import randint
import openpyxl, sys

if len(sys.argv) != 2 or sys.argv[1][-4:] != "xlsx":
    print("Feed me a workbook !")
else:
    wb = openpyxl.load_workbook(filename = sys.argv[1])

    sheets = wb.sheetnames
    ws = wb[sheets[0]]

    possible_values = ['A', 'B', 'C', 'D']
    cells = [ 'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21', 'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28', 'B29', 'B30', 'B31', 'B32', 'B33', 'B34', 'B35', 'B36', 'B37', 'B38', 'B39', 'B40', 'B41' ]
    
    for i in cells:
        ws[i] = possible_values[randint(0, 3)]

    print("Donezo. Now run !")
    wb.save(sys.argv[1])

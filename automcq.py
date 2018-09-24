#!/usr/bin/python3

from random import randint
import openpyxl, sys, platform, os.path

#Get OS
host = platform.system()

#Flags
blankmode = False
verbose = False
go = True
file = None

#Parsing command line args
if len(sys.argv) < 2:
    go = False

for i in range(len(sys.argv)):
    if sys.argv[i][-4:] == "xlsx":
        file = sys.argv[i]
    elif sys.argv[i] == "--blank":
        blankmode = True
    elif sys.argv[i] == "--verbose":
        verbose = True
    elif sys.argv[i][0] == '-' and sys.argv[i][1] != '-':
        if "b" in sys.argv[i]:
            blankmode = True
        if "v" in sys.argv[i]:
            verbose = True
    elif sys.argv[i] == "help":
        print("AutoMCQ v1.01 by Bad64")
        print("Usage: automcq [switches] [xlsx file]")
        print("    -b or --blank: Only fills in blank cells (do not overwrite filled cells)")
        print("    -v or --verbose: Prints everything to the console")
        go = False

#Check if file exists
if file is not None:
    if not os.path.isfile(file):
        file = None
    else:
        if verbose:
            print("Operating on file", file, ":")

#Main logic
if go and file:
    wb = openpyxl.load_workbook(filename = file)

    sheets = wb.sheetnames
    ws = wb[sheets[0]]

    possible_values = ['A', 'B', 'C', 'D']
    cells = [ 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21', 'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28', 'B29', 'B30', 'B31', 'B32', 'B33', 'B34', 'B35', 'B36', 'B37', 'B38', 'B39', 'B40', 'B41' ]
    
    for i in cells:
        if blankmode:
            if not ws[i].value:
                if not verbose:
                    ws[i] = possible_values[randint(0, 3)]
                else:
                    char = possible_values[randint(0, 3)]
                    if host == "Linux":
                        print("\033[92mWriting", char, "to cell", i)
                    else:
                        print("Writing", char, "to cell", i)
                    ws[i] = char
            else:
                if verbose:
                    print("\033[93mSkipping cell", i)
        else:
            if not verbose:
                ws[i] = possible_values[randint(0, 3)]
            else:
                char = possible_values[randint(0, 3)]
                if host == "Linux":
                    print("\033[92mWriting", char, "to cell", i)
                else:
                    print("Writing", char, "to cell", i)
                ws[i] = char

    if host == "Linux":
        print("\033[0mDonezo. Now get out of here !")
    else:
        print("Donezo. Now get out of here !")
    wb.save(file)
elif go and not file:
    print("You must supply a valid xlsx file !")
elif not go:
    print("Feed me a workbook ! Or type \"automcq help\" to get help.")

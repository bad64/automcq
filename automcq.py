#!/usr/bin/python3

import openpyxl, sys, platform, random, os.path

#Get OS
host = platform.system()

#Flags
blankmode = False
verbose = False
go = True
file = None
saveAs = None

#Parsing command line args
if len(sys.argv) < 2:
    go = False

for i in range(len(sys.argv)):
    if sys.argv[i][-4:] == "xlsx" and (sys.argv[i-1] != "-o" or sys.argv[i-1] != "--output"):
        if not file:
            file = sys.argv[i]
    elif sys.argv[i] == "--blank":
        blankmode = True
    elif sys.argv[i] == "--verbose":
        verbose = True
    elif sys.argv[i] == "-o" or sys.argv[i] == "--output":
        saveAs = sys.argv[i+1]
        i += 1
    elif sys.argv[i][0] == '-' and sys.argv[i][1] != '-':
        if "b" in sys.argv[i]:
            blankmode = True
        if "v" in sys.argv[i]:
            verbose = True
    elif sys.argv[i] == "help":
        print("AutoMCQ v1.07 by Bad64")
        print("Usage: automcq [switches] [xlsx file]")
        print("    -b or --blank: Only fills in blank cells (do not overwrite filled cells)")
        print("    -v or --verbose: Prints everything to the console")
        print("    -o <file> or --output <file>: Outputs the new filled workbook to file")
        go = False

#Check if file exists
if file is not None:
    if not os.path.isfile(file):
        file = None
    else:
        if verbose:
            print("Operating on file", file, ":")

#Validating output file
if saveAs and saveAs[-5:] != ".xlsx":
    saveAs += ".xlsx"

#Main logic
if saveAs is not None:
    print("Writing to file", saveAs)
    
if go and file:
    wb = openpyxl.load_workbook(filename = file)

    sheets = wb.sheetnames
    ws = wb[sheets[0]]

    random.SystemRandom()

    possible_values = ['A', 'B', 'C', 'D']
    cells = [ 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21', 'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28', 'B29', 'B30', 'B31', 'B32', 'B33', 'B34', 'B35', 'B36', 'B37', 'B38', 'B39', 'B40', 'B41' ]
    
    for i in range(len(cells)):
        if blankmode:
            if not ws[cells[i]].value:
                if not verbose:
                    ws[cells[i]] = possible_values[random.randint(0, 3)]
                else:
                    char = possible_values[random.randint(0, 3)]
                    if host == "Linux":
                        print("\033[92mAnswering", char, "to question", i+1)
                    else:
                        print("Answering", char, "to question", i+1)
                    ws[cells[i]] = char
            else:
                if verbose:
                    if host == "Linux":
                        print("\033[93mSkipping question", i+1)
                    else:
                        print("Skipping question", i+1)
        else:
            if not verbose:
                ws[cells[i]] = possible_values[random.randint(0, 3)]
            else:
                char = possible_values[random.randint(0, 3)]
                if host == "Linux":
                    print("\033[92mAnswering", char, "to question", i+1)
                else:
                    print("Answering", char, "to cell", i+1)
                ws[cells[i]] = char

    if host == "Linux":
        print("\033[0mDonezo. Now get out of here !")
    else:
        print("Donezo. Now get out of here !")

    if saveAs:
        wb.save(saveAs)
    else:
        wb.save(file)
        
elif go and not file:
    print("You must supply a valid xlsx file !")
elif not go:
    print("Feed me a workbook ! Or type \"automcq help\" to get help.")

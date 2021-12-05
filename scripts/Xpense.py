""" This script should uphold the following format and purpose:

1. Ask the user what they they are doing with their expenses tracking
    - Adding expenses to an existing sheet
    - Creating a new expenses sheet
    - Deleting a current expenses sheet
    - Examining expenses across multiple (or a single) sheet(s) (should just be a printout of expenses, totals for each month, and combined expenses for each month)
        - In addition, we can optionally show a percentage breakdown for each category of expenses
2. Ask the user for input as to which excel spreadsheet they are using to input expense information (give a list of current sheets)
    - if the user inputs a sheet that isn't there, a new sheet is created
    - once the user inputs an excel sheet, we run a loop until the user says that they're finished inputting items. There will be separate loops for costs and expenses to allow continuity and quick input
    - give an exit input at all stages of the for loop (memo, date, amount) such as -exit
3. All arguments should be able to be inputted via command line
    - memo arguments for costs should be put in the first empty space of column A
    - date arguments for costs should be put in the first empty space of column B under B3
    - Amount arguments for costs should be put in the first empty space of column C under C3
    - Same for Income for columns D, E and F, respectively
    - After inputting each costs/expense, the program should display "Income", "Money spent", and "Balance (Income-money spent)"
"""
from openpyxl import load_workbook

# Xpense tracker
action = input("""Welcome to Xpense tracker! What would you like to do today? Options are as follows: \n 
                -Add expenses to an existing sheet [AE] \n 
                -Remove expenses from an existing sheet [RE] \n
                -Create a new expenses sheet [NS] \n 
                -Delete a current expenses sheet [DS] \n 
                -Examine expenses across multiple sheets [EE]: \n""")

wbPath = "C:\\Users\\keess\\Desktop\\2021_Budget_spreadsheet.xlsx"
# loading workbook to insert new sheets
wb = load_workbook(wbPath)

if (action == "NS"):
    print("Current sheet names: ")
    print(wb.sheetnames)
    sheetName = input("What would you like to name your new sheet?: ")
    target = wb['template']
    wb.copy_worksheet(target)
    wb_sheet = wb["template Copy"]
    wb_sheet.title = sheetName

    askPay = input("Are your paychecks for the same amount as last month? [y] or [n] ")

    if (askPay == "y"):
        payVal = input("How much money do you take home from each paycheck? ")
        payVal = float(payVal)
        numChecks = input("How many paychecks are you expecting to receive this month?")
        numChecks = float(numChecks)
        MonthNum = input("What is the number corresponding to this month (1-12):    ")
        Year = input("What is the year?:    ")
        for (i in numChecks):
            Amount_cell = "F" + i+4
            Memo_cell = "D" + i+4
            Date_cell = "E" + i+4
            wb_sheet[Amount_cell] = payVal
            wb_sheet[Memo_cell] = "Salary"
            wb_sheet[Date_cell] = Year + "-" + MonthNum + "-" + "1"

    wb.save(wbPath)
    
    

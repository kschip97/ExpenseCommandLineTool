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
from openpyxl.styles import NamedStyle
import easygui
import sys

date_style = NamedStyle(name='datetime', number_format='YYYY-MM-DD')
# Creating main() function as framework for making code more readable

def main():

    # User input determines what function is called depending on the task being done
    action = easygui.buttonbox("Welcome to Xpense tracker! What would you like to do today?", 'Action', 
    ('Add Expense/Income', 'Remove Expense/Income', 'Make New Expense Sheet', 'Delete Expense Sheet', 'Compare Multiple Sheets', 'Create New Expense Workbook'))

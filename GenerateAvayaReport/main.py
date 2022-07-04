from calendar import calendar
import datetime
import calendar

from getFiles import getFiles
from editMacro import editMacroFile
from runMacro import runMacro

def main():

    # Get current date
    year = datetime.datetime.today().year
    month = datetime.datetime.today().month
    day = datetime.datetime.today().day

    # Convert month from integer to name
    monthName = calendar.month_name[month]

    # Add leading zero to month
    if len(str(month)) == 1:
        month = f'0{month}'

    # Run main functions
    # Saves all files within the working directory as a list
    entries = getFiles(year, month, day)

    # Edits the macro run and saves its path to run in the next line
    path = editMacroFile(year, month, day, monthName, entries)

    # Running the VBA macro
    runMacro(path)

# Run main program
main()

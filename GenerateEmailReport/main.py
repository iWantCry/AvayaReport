from generateTable import generateTable
from askDir import askDir

import datetime
import calendar
import os

def main():
    # Main program

    #Get current date
    month = datetime.datetime.today().month
    day = datetime.datetime.today().day

    monthNameAbbr = calendar.month_abbr[month]

    # Add leading zero to month
    if len(str(month)) == 1:
        month = f'0{month}'

    # Ask directory from user
    path = askDir()

    # Generate table
    generateTable(path, day, monthNameAbbr)


main()
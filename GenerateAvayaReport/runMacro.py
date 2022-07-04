import os, os.path
import win32com.client

def runMacro(path):
    # Run VBA macro to generate report

    # File to run macro
    file = "Avaya_Report latest.xlsm"

    # If working path does exist, run macro
    if os.path.exists(path):
        xl=win32com.client.Dispatch("Excel.Application")

        xl.Workbooks.Open(path)

        xl.Application.Run('\'' + file + '\'' + "!Module1.Open_Avaya_Report")

        xl.Application.Save()

        xl.Application.Quit()
        del xl
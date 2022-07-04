import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import datetime      

from askDir import askDir

def getFiles(year, month, day):
    # Ask directory from user
    path = askDir()

    # Check if all required files for macro are present
    check = []

    entries = os.listdir(path)

    for entry in entries:
        if entry.startswith(f"Agent All_{year}-{month}-{day}"):
            check.append("True1")

        elif entry.startswith(f"Avaya Reports - {year}-{month}-{day-1}"):
            check.append("True2")

        elif entry.startswith("avaya"):
            check.append("True3")
            
        elif entry.startswith("Avaya_Report latest"):
            check.append("True4")

        elif entry.startswith(f"Contact_Details_{datetime.datetime.today().month}-{day-1}-{year}_{datetime.datetime.today().month}-{day-1}-{year}"):
            check.append("True5")

        elif entry.startswith("CRVReport"):
            check.append("True6")

        elif entry.startswith(f"Topics by All hours_{year}-{month}-{day}"):
            check.append("True7")

        elif entry.startswith(f"Topics by Days_{year}-{month}-{day}"):
            check.append("True8")

    # Check = True (All files present)
    if len(check) == 8:
        print("Files ready for macro")
        return entries

    # Check = False (Some files are not present/invalid)
    else:
        print("Error file(s) not present!")

        # Tkinter GUI for invalid files
        ws = tk.Tk()
        ws.geometry("300x200")
        ws.title("Error!")

        w = tk.Label(ws, text="File(s) not found!", font="50")
        w.pack(side=tk.TOP, pady=20)

        tk.Button(ws, text="Ok", command=ws.quit).pack(side=tk.BOTTOM, pady=20)

        ws.mainloop()

        exit()
        
        

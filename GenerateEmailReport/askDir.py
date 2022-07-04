import tkinter as tk
from tkinter import filedialog

def askDir():
    # Asking user for directory to generate table in
    root = tk.Tk()

    # Hide main window
    root.withdraw()

    path = filedialog.askopenfilename(initialdir="/", title="Select avaya workbook", filetypes=(("Excel Workbook", "*.xlsx"), ("All files", "*.*")))

    # Destroy main window
    root.destroy()

    return path
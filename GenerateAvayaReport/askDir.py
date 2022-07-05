import tkinter as tk
from tkinter import filedialog

def askDir():
    # Asking user for directory to generate report in
    root = tk.Tk()

    # Hide main window
    root.withdraw()

    path = filedialog.askdirectory(initialdir="/", title="Select folder for Avaya Report Generation")

    # Destroy main window
    root.destroy()

    return path
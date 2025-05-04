"""
Main entry point for the TRAK File Builder application.
"""
import tkinter as tk
from gui import FileConverterApp

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()
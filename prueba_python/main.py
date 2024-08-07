import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os

from folder_selector_app import folder_selector_app


if __name__ == "__main__":
    root = tk.Tk()
    app = folder_selector_app(root)
    root.mainloop()
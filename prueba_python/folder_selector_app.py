import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os
from file_reader import file_reader 
from excel_consolidator import excel_consolidator
from file_organizer import file_organizer

class folder_selector_app:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Consolidator 1.1.0")

        self.frame = ttk.Frame(root, padding="10 30 10 10")
        self.frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.geometry("700x200")
        self.folder_path = tk.StringVar()
        self.file_path = tk.StringVar()
        self.root.resizable(False, False) 
        self.create_menu()
        self.create_widgets()
        self.progress['maximum'] = 100

    def create_widgets(self):
        self.folder_label = ttk.Label(self.frame, text="Folder Path:")
        self.folder_label.grid(column=1, row=2, sticky=tk.W)

        self.folder_entry = ttk.Entry(self.frame, width=70, textvariable=self.folder_path, state="readonly")
        self.folder_entry.grid(column=2, row=2, sticky=(tk.W, tk.E))

        self.browse_button = ttk.Button(self.frame, text="Select Folder", command=self.browse_folder)
        self.browse_button.grid(column=3, row=2, sticky=tk.W)

        self.file_label = ttk.Label(self.frame, text="Master File Path:")
        self.file_label.grid(column=1, row=3, sticky=tk.W)

        self.file_entry = ttk.Entry(self.frame, width=50, textvariable=self.file_path, state="readonly")
        self.file_entry.grid(column=2, row=3, sticky=(tk.W, tk.E))

        self.browse_file_button = ttk.Button(self.frame, text="Select Master File", command=self.browse_file)
        self.browse_file_button.grid(column=3, row=3, sticky=tk.W)

        self.execute_button = ttk.Button(self.frame, text="Consolidate", command=self.execute_action)
        self.execute_button.grid(column=3, row=4, sticky=tk.W)

        self.progress = ttk.Progressbar(self.frame, length=100, mode='determinate')
        self.progress.grid(column=2, row=5, sticky=tk.W)

        for child in self.frame.winfo_children(): 
            child.grid_configure(padx=5, pady=5)

    def create_menu(self):
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        about_menu = tk.Menu(menu_bar, tearoff=0)
        about_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Help", command=self.show_help)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        menu_bar.add_cascade(label="About", menu=about_menu)

    def show_about(self):
        messagebox.showinfo("About", "This is an excel file consolidation application.\nVersion 1.1.0 - Americas Digital Hub")

    def show_help(self):
        messagebox.showinfo("Help", "To select a folder, click the 'Select Folder' button, then to select a File Master, then click the Consolidate button'.")

    def browse_folder(self):
        self.update_progress(0)
        folder_path = filedialog.askdirectory()
        self.folder_path.set(folder_path)

    def browse_file(self):
        self.update_progress(0)
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        self.file_path.set(file_path)
        
    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()

    def execute_action(self):
        self.update_progress(0)
        folder_path = self.folder_path.get()
        file_path = self.file_path.get()
        if not folder_path or not file_path:
            messagebox.showwarning("Warning", "Please select a folder and master file first.")
            return
        
        file_reader_1 = file_reader(folder_path)
        excel_files = file_reader_1.list_files()
        
        if excel_files:
           excel_consolidator_1 = excel_consolidator(excel_files, file_path, self)
           excel_consolidator_1.consolidate_sheets()
           file_organizer_1 = file_organizer(folder_path)
           file_organizer_1.move_files() 
           messagebox.showinfo("Completed", "Consolidation completed successfully.")
        else:
            messagebox.showinfo("Excel Files", "No Excel files were found in the folder.")


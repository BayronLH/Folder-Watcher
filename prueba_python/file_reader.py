
import os
from tkinter import messagebox

class file_reader:
    def __init__(self, folder_path):
        self.folder_path = folder_path

    def list_files(self):
        try:
            files = os.listdir(self.folder_path)
            excel_files = [file for file in files if file.endswith(('.xls', '.xlsx','.xlsm','.xlsb'))]
            files_complete = [os.path.join(self.folder_path, f) for f in excel_files]
            return files_complete
        except Exception as e:
            messagebox.showerror("Error", f"Could not read folder: {e}")
            return []
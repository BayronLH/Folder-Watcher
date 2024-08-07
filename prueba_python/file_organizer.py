import os
import shutil
from tkinter import filedialog, Tk, messagebox

class file_organizer:
    def __init__(self, source_folder):
        self.source_folder = source_folder
        self.excel_folder = source_folder+ "/Processed"
        self.non_excel_folder = source_folder+ "/Not_aplicable"

    def move_files(self):
        try:  
            os.makedirs(self.excel_folder, exist_ok=True)
            os.makedirs(self.non_excel_folder, exist_ok=True)
            files = os.listdir(self.source_folder)

            for file in files:
                file_path = os.path.join(self.source_folder, file)

                if os.path.isfile(file_path):
                   
                    if file.lower().endswith(('.xls', '.xlsx','.xlsm','.xlsb')):
                      
                        shutil.move(file_path, os.path.join(self.excel_folder, file))
                    else:
                        shutil.move(file_path, os.path.join(self.non_excel_folder, file))

            return "Files moved successfully."
        except Exception as e:
            return f"Error during file movement: {e}"
import xlwings as xw
from tkinter import messagebox
import os

class excel_consolidator:
    def __init__(self, excel_files, master_file, app):
        self.excel_files = excel_files
        self.master_file = master_file
        self.app = app

    def consolidate_sheets(self):
        
        app = xw.App(visible=False)  # type: ignore
        try:
            
            output_wb = xw.Book(self.master_file)
            counter = 1
            counter_progress = 0
            progress_total = len(self.excel_files)

            for file in  self.excel_files:
              counter_progress += 1
              self.app.update_progress(counter_progress/progress_total*100)
             
              existing_sheet_names = [sheet.name for sheet in output_wb.sheets]
             
              input_wb = xw.Book(file)
              input_sheets = input_wb.sheets
              
              for sheet in input_sheets:
                original_name = sheet.name
                new_name = original_name
               
                while new_name in existing_sheet_names:
                    file_name_base  = os.path.basename(file)
                    file_name, _ = os.path.splitext(file_name_base)
                    new_name = f"{original_name}_{file_name}_{counter}"
                    counter += 1

                sheet.copy(after=output_wb.sheets[-1])
                output_wb.sheets[-1].name = new_name
                
              output_wb.save()
              input_wb.close()
            output_wb.close()
            self.update_progress(100)
            
            return "Consolidation completed successfully."
        except Exception as e:
            return f"Error during consolidation: {e}"
        finally:
            app.quit()
          
           
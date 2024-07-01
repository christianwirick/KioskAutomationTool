import tkinter as tk
from tkinter import filedialog, simpledialog
import openpyxl
import os
from collections import defaultdict
class ExcelProcessor:
   def __init__(self, main_file, lookup_file):
       self.main_file = main_file
       self.lookup_file = lookup_file
       self.main_wb = None
       self.lookup_wb = None
       self.segment_data = defaultdict(list)
   def load_workbooks(self):
       try:
           self.main_wb = openpyxl.load_workbook(self.main_file)
           self.lookup_wb = openpyxl.load_workbook(self.lookup_file)
       except Exception as e:
           print(f"Error loading workbooks: {e}")
           raise
   def prepare_lookup_data(self):
       try:
           lookup_sheet = self.lookup_wb.active
           self.mgt_data = {row[0].value: row[1].value for row in lookup_sheet.iter_rows(min_row=2)}
       except Exception as e:
           print(f"Error preparing lookup data: {e}")
           raise
   def process_data(self):
       try:
           sheet = self.main_wb.active
           sheet.insert_cols(4, 4)  # Insert columns for new data
           for row in range(2, sheet.max_row + 1):
               left_chars = sheet[f'C{row}'].value[:2] if sheet[f'C{row}'].value else ''
               sheet[f'D{row}'] = left_chars
               sheet[f'E{row}'] = self.mgt_data.get(left_chars, '')  # VLOOKUP equivalent
               right_chars = sheet[f'C{row}'].value[-9:] if sheet[f'C{row}'].value else ''
               sheet[f'F{row}'] = right_chars
               sheet[f'G{row}'] = f'{sheet[f'E{row}'].value}{sheet[f'F{row}'].value}'
               segment_name = sheet[f'H{row}'].value  # Segment name in column H
               if segment_name:
                   self.segment_data[segment_name].append(sheet[f'G{row}'].value)
           print(f"Processed {sheet.max_row - 1} rows.")
       except Exception as e:
           print(f"Error processing data: {e}")
           raise
   def export_data(self):
       choice = simpledialog.askstring("Choose Export Option", "Type 'ALL' for a single file or 'BY SEGMENT' for separate files")
       if choice.upper() == "ALL":
           self.export_all_data()
       elif choice.upper() == "BY SEGMENT":
           self.export_by_segment()
       else:
           print("Invalid choice. Please type 'ALL' or 'BY SEGMENT'.")
   def export_all_data(self):
       try:
           os.makedirs('MGT', exist_ok=True)
           all_data_wb = openpyxl.Workbook()
           all_data_sheet = all_data_wb.active
           row = 1
           for segment, data in self.segment_data.items():
               for item in data:
                   all_data_sheet[f'A{row}'].value = int(float(item))  # Convert to integer
                   all_data_sheet[f'A{row}'].number_format = "0"  # Set number format
                   row += 1
           all_data_filename = 'MGT/All_Data.xlsx'
           all_data_wb.save(all_data_filename)
           print("Exported all data to a single Excel file.")
       except Exception as e:
           print(f"Error exporting all data: {e}")
           raise
   def export_by_segment(self):
       try:
           os.makedirs('MGT', exist_ok=True)
           for segment, data in self.segment_data.items():
               segment_wb = openpyxl.Workbook()
               segment_sheet = segment_wb.active
               row = 1
               for item in data:
                   segment_sheet[f'A{row}'].value = int(float(item))  # Convert to integer
                   segment_sheet[f'A{row}'].number_format = "0"  # Set number format
                   row += 1
               segment_filename = f'MGT/{segment}.xlsx'
               segment_wb.save(segment_filename)
           print("Exported data by segment into individual Excel files.")
       except Exception as e:
           print(f"Error exporting by segment: {e}")
           raise
def select_excel_file(title):
   root = tk.Tk()
   root.withdraw()
   filename = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx;*.xls")])
   root.destroy()
   return filename
def main():
   main_file = select_excel_file("Select the main Excel file")
   if not main_file:
       print("No main file selected.")
       return
   lookup_file = select_excel_file("Select the MGT Prefix Codes Excel file")
   if not lookup_file:
       print("No lookup file selected.")
       return
   processor = ExcelProcessor(main_file, lookup_file)
   processor.load_workbooks()
   processor.prepare_lookup_data()
   processor.process_data()
   processor.export_data()
   print("Processing completed successfully.")
if __name__ == "__main__":
   main()
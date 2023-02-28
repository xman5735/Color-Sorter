import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

class ExcelSearchGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel Extractor")
        self.master.geometry("400x200")
        
        self.file_path = ""
        self.selected_row = ""
        self.search_value = ""
        
        self.file_label = tk.Label(self.master, text="No file selected.")
        self.file_label.pack()
        
        self.select_button = tk.Button(self.master, text="Select File", command=self.select_file)
        self.select_button.pack(pady=(10, 0))
        
        self.row_label = tk.Label(self.master, text="Select Row:")
        self.row_label.pack()
        
        self.row_entry = tk.Entry(self.master)
        self.row_entry.pack(pady=(10, 0))
        
        self.search_label = tk.Label(self.master, text="Enter Value to Search:")
        self.search_label.pack()
        
        self.search_entry = tk.Entry(self.master)
        self.search_entry.pack(pady=(10, 0))
        
        self.submit_button = tk.Button(self.master, text="Submit", command=self.submit)
        self.submit_button.pack(pady=(10, 0))
        
    def select_file(self):
        self.file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        if self.file_path:
            self.file_label.config(text=f"Selected file: {self.file_path}")
            
    def submit(self):
        if not self.file_path:
            tk.messagebox.showerror("Error", "Please select an Excel file first.")
            return
        
        self.selected_row = self.row_entry.get().strip()
        if not self.selected_row:
            tk.messagebox.showerror("Error", "Please enter a row number to search.")
            return
        
        self.search_value = self.search_entry.get().strip()
        if not self.search_value:
            tk.messagebox.showerror("Error", "Please enter a value to search.")
            return
        
        try:
            df = pd.read_excel(self.file_path, header=None, sheet_name=0)
            row_num = int(self.selected_row) - 1
            target_row = df.iloc[row_num]
            
            matches = []
            for i, value in target_row.iteritems():
                if str(value) == self.search_value:
                    matches.append(df[i])
            
            if not matches:
                tk.messagebox.showinfo("No Matches Found", "No columns contain the specified value.")
                return
            
            new_file_name = os.path.splitext(os.path.basename(self.file_path))[0] + "_" + self.search_value + "_extract.xlsx"
            new_file_path = os.path.join(os.path.expanduser("~"), "Desktop", new_file_name)
            concat_df = pd.concat(matches, axis=1)
            concat_df.to_excel(new_file_path, header=False, index=False)
            
            tk.messagebox.showinfo("Success", f"Matching columns saved to {new_file_path}")
            
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {e}")
        
root = tk.Tk()
gui = ExcelSearchGUI(root)
root.mainloop()

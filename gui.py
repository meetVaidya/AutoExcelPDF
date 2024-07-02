import time
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os

class AutoExcelPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoExcelPDF")
        self.root.geometry("800x450")
        self.root.configure(bg='black')

        self.file_path = None

        self.create_main_frame()
        self.create_input_frame()

    def create_main_frame(self):
        self.main_frame = tk.Frame(self.root, bg='black')
        self.main_frame.pack(pady=20)

        title = tk.Label(self.main_frame, text="AutoExcelPDF", font=("Arial", 24), bg='black', fg='white')
        title.pack(pady=10)

        file_frame = tk.Frame(self.main_frame, bg='black')
        file_frame.pack(pady=10)

        file_label = tk.Label(file_frame, text="File Path:", bg='black', fg='white')
        file_label.grid(row=0, column=0, padx=5)

        self.file_entry = tk.Entry(file_frame, width=50)
        self.file_entry.grid(row=0, column=1, padx=5)

        browse_button = tk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_button.grid(row=0, column=2, padx=5)

        submit_button = tk.Button(self.main_frame, text="Submit", command=self.submit_file)
        submit_button.pack(pady=10)

    def create_input_frame(self):
        self.input_frame = tk.Frame(self.root, bg='black')
        self.input_frame.pack(pady=20)
        self.input_frame.pack_forget()  # Initially hidden

        sheet_label = tk.Label(self.input_frame, text="input sheet name", bg='black', fg='white')
        sheet_label.grid(row=0, column=0, padx=5, pady=5)
        self.sheet_entry = tk.Entry(self.input_frame)
        self.sheet_entry.grid(row=0, column=1, padx=5, pady=5)

        bill_sheet_label = tk.Label(self.input_frame, text="input bill sheet", bg='black', fg='white')
        bill_sheet_label.grid(row=1, column=0, padx=5, pady=5)
        self.bill_sheet_entry = tk.Entry(self.input_frame)
        self.bill_sheet_entry.grid(row=1, column=1, padx=5, pady=5)

        row_label = tk.Label(self.input_frame, text="input row number", bg='black', fg='white')
        row_label.grid(row=2, column=0, padx=5, pady=5)
        self.row_entry = tk.Entry(self.input_frame)
        self.row_entry.grid(row=2, column=1, padx=5, pady=5)

        go_button = tk.Button(self.input_frame, text="GO !", command=self.on_go)
        go_button.grid(row=3, column=0, columnspan=2, pady=10)

        timer_label = tk.Label(self.input_frame, text="you will have 5 seconds to open excel", bg='black', fg='white')
        timer_label.grid(row=4, column=0, columnspan=2, pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*xlsm")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def submit_file(self):
        self.file_path = self.file_entry.get()
        if self.file_path:
            _, file_extension = os.path.splitext(self.file_path)
            if file_extension.lower() in ['.xlsx', '.xls', '.xlsm']:
                self.input_frame.pack()  # Show input frame
            else:
                messagebox.showerror("Error", "Please select an Excel file")
        else:
            messagebox.showerror("Error", "Please enter a file path")

    def on_go(self):
        sheet_name = self.sheet_entry.get()
        bill_sheet = self.bill_sheet_entry.get()
        row_number = self.row_entry.get()
        print(f"Processing file: {self.file_path}")
        print(f"Sheet name: {sheet_name}")
        print(f"Bill sheet: {bill_sheet}")
        print(f"Row number: {row_number}")
        # Here you would add the logic to process the Excel file
        import aggregate
        aggregate.aggregate_data(bill_sheet=bill_sheet, row=row_number)
        
        time.sleep(2)
        
        import test
        test.format_dates(file_path=self.file_path)
        
        time.sleep(2)
        
        import print1
        print1.process_excel_data(sheet_name=sheet_name)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoExcelPDF(root)
    root.mainloop()
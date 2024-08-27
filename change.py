import win32com.client

def modify_excel():
    # Create an instance of the Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Optional: make Excel visible
    excel.Visible = True

    # Open the Excel file
    file_path = r'D:\Projects\test\test.xlsm'
    workbook = excel.Workbooks.Open(file_path)

    # Access a specific sheet by name
    sheet_name = 'YJV'
    sheet = workbook.Sheets(sheet_name)

    # Modify cell values
    # Example: Set the value of cell A1
    with open('bill_no.txt', 'r') as file:
        bill_no = int(file.read())
        sheet.Cells(11, 6).Value = bill_no
        bill_no += 1
    workbook.Save()

    excel.Quit()

# Call the function to modify the Excel file
modify_excel()
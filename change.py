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

    # with open('bill_no.txt', 'w') as file:
    #     file.write(str(bill_no))

    # Example: Set the value of cell B2
    # sheet.Cells(2, 2).Value = 123

    # Example: Set the value of a range of cells
    # sheet.Range('C1:C3').Value = [1, 2, 3]

    # Optional: Save the workbook if changes were made
    workbook.Save()

    # Optional: Close the workbook
    # workbook.Close()

    # Optional: Quit the Excel application
    excel.Quit()

# Call the function to modify the Excel file
modify_excel()
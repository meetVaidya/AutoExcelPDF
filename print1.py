import pyautogui
import time
import file_import
import win32com.client

def process_excel_data(sheet_name):
    print("You have 5 seconds to switch to the Excel application")
    
    # Create an instance of the Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Optional: make Excel visible
    excel.Visible = True

    # Open the Excel file
    file_path = file_import.excel_path
    workbook = excel.Workbooks.Open(file_path)

    # Access and activate a specific sheet by name
    sheet_name = sheet_name
    sheet = workbook.Sheets(sheet_name)
    sheet.Activate()

    time.sleep(5)

    with open('bill_no.txt', 'r') as file:
        bill_no = [line.strip() for line in file]

    with open('company_name.txt', 'r') as file:
        company_name = [line.strip() for line in file]

    with open('month.txt', 'r') as file:
        month = [line.strip() for line in file]

    company_counter = 0
    month_counter = 0

    for bill in bill_no:
        sheet.Cells(11, 6).Value = bill

        time.sleep(1)

        # Simulate pressing Ctrl + P
        pyautogui.hotkey('ctrl', 'p')

        time.sleep(1)

        # Simulate pressing Enter
        pyautogui.press('enter')

        time.sleep(1)

        bill = str(bill)
        company = str(company_name[company_counter])
        m = str(month[month_counter])

        file_name = bill + "_" + company + "_" + m
        pyautogui.typewrite(file_name)
        pyautogui.press('enter')
        time.sleep(0.5)

        company_counter += 1
        month_counter += 1

        # Close the Excel application
        workbook.Save()

    workbook.Close()
    print("All done!")

# # Call the function
# process_excel_data()

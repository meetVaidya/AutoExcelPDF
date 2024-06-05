import pyautogui
import win32com.client
import time

# Create an instance of the Excel application
excel = win32com.client.Dispatch("Excel.Application")

# Optional: make Excel visible
excel.Visible = True

# Open the Excel file
file_path = r'D:\Projects\test\test.xlsm'

workbook = excel.Workbooks.Open(file_path)

# Access and activate a specific sheet by name
sheet_name = 'YJV'
sheet = workbook.Sheets(sheet_name)
sheet.Activate()

# Bring the Excel application to the active screen
# pyautogui.getWindowsWithTitle('Microsoft Excel')[0].activate()
# sheet.Activate()

# bill_no = [4141,4142,4143,4144,4145,4146,4147,4148,4149,4150,4151,4152,4153,4154,4155,4156,4157,4158,4159,4160,4161,4162,4163,4164,4165,4166,4167,4168,4169,4170,4171,4172,4173,4174,4175,4176,4177,4178,4179,4180,4181,4182,4183,4184,4185,4186,4187,4188,4189,4190,4191,4192,4193,4194,4195,4196,4197,4198,4199,4200,4201,4202,4203,4204,4205,4206,4207,4208,4209,4210,4211,4212]

with open('bill_no.txt', 'r') as file:
    bill_no = [line.strip() for line in file]

with open('company_name.txt', 'r') as file:
    company_name = [line.strip() for line in file]

with open('month.txt', 'r') as file:
    month = [line.strip() for line in file]

# print(fruits_list)

for bill in bill_no:
    sheet.Cells(11, 6).Value = bill

    time.sleep(1)   

    # Simulate pressing Ctrl + P
    pyautogui.hotkey('ctrl', 'p')

    time.sleep(1)

    # Simulate pressing Enter
    pyautogui.press('enter')

    time.sleep(1)

    # Input file name using pyautogui
    file_name = str(bill) + str(company_name) + str(month)
    pyautogui.typewrite(file_name)
    pyautogui.press('enter')

    time.sleep(0.5)

    # Close the Excel application
    workbook.Save()

    # str(bill) + str(company_name) + str(month)
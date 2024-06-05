# import win32com.client

# # Create an instance of the Excel application
# excel = win32com.client.Dispatch("Excel.Application")

# # Optional: make Excel visible
# excel.Visible = True

# # Open the Excel file
# file_path = r'D:\test\test.xlsm'
# workbook = excel.Workbooks.Open(file_path)

# # Access a specific sheet by name
# sheet_name = 'YJV'
# sheet = workbook.Sheets(sheet_name)

# # Modify cell values
# # Example: Set the value of cell A1
# sheet.Cells(11, 6).Value = '4211'

# # Example: Set the value of cell B2
# # sheet.Cells(2, 2).Value = 123

# # Example: Set the value of a range of cells
# # sheet.Range('C1:C3').Value = [1, 2, 3]

# # Optional: Save the workbook if changes were made
# workbook.Save()

# # Optional: Close the workbook
# # workbook.Close()

# # Optional: Quit the Excel application
# # excel.Quit()

# file_path = "D:/Projects/test/month.txt"

# # Read the file
# with open(file_path, 'r') as file:
#     lines = file.readlines()

# # Add a comma after each line
# lines_with_comma = [line.strip() + ',' for line in lines]

# # Write the modified lines back to the file
# with open(file_path, 'w') as file:
#     file.writelines(lines_with_comma)

import re

# Read the contents of the file
with open('D:/Projects/test/month.txt', 'r') as file:
    text = file.read()

dates = [text[i:i+12] for i in range(0, len(text), 12)]

# Print each combination on its own line
# for date in dates:
#     print(date)

# Write the formatted content back to the file
with open('D:/Projects/test/month.txt', 'w') as file:
    file.write('\n'.join(dates))
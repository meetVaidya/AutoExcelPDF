import numpy as np
import file_import
import win32com.client

def aggregate_data(bill_sheet, row):
    # Path to the uploaded Excel file
    file_path = file_import.excel_path  # take the path of the file from the user
    bill_text_file = file_import.bill_file  # take the path of the file from the user
    company_text_file = file_import.company_file  # take the path of the file from the user
    month_text_file = file_import.month_file  # take the path of the file from the user

    # Create an instance of the Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Optional: make Excel visible (for debugging purposes)
    excel.Visible = False

    # Open the Excel file
    workbook = excel.Workbooks.Open(file_path)

    # Access the 'Bills List' sheet
    sheet = workbook.Sheets('Bills List')

    # Initialize empty lists to store client names
    bill_no = []
    company_name = []
    month_column = []

    # Define the column indices (assuming 'Name' is in column 1 and 'Client' is in column 2)
    name_column = 1  # take the column number from the user
    bill_column = 2  # take the column number from the user
    company_column = 3  # take the column number from the user
    monthly_column = 11  # take the column number from the user

    # Loop through the rows in the sheet (assuming the data starts from row 2)
    row =  int(row) # take the row number from the user
    while True:
        try:
            # Read the value of the 'Name' column
            name_value = sheet.Cells(row, name_column).Value

            # Break the loop if we have reached the end of the data
            if name_value is None:
                break

            # Check if the 'Name' column contains 'YJV'
            if bill_sheet in name_value:
                # Read the value of the 'Client' column
                client_bill = sheet.Cells(row, bill_column).Value
                client_name = sheet.Cells(row, company_column).Value
                client_month = sheet.Cells(row, monthly_column).Value

                # Append the client data to the lists (accepting strings now)
                bill_no.append(client_bill)
                company_name.append(client_name)
                month_column.append(client_month)

            # Move to the next row
            row += 1

        except Exception as e:
            # Skip the row if an error occurs (e.g., the value is not a number)
            row += 1
            continue

    # Optional: Close the workbook and quit the Excel application
    workbook.Close(False)
    excel.Quit()

    # bill_no list now contains strings, no need for int conversion

    with open(bill_text_file, 'w') as f:
        for item in bill_no:
            # Write each item on a new line
            f.write("%s\n" % item)

    # Do the same for the company_name list
    with open(company_text_file, 'w') as f:
        for item in company_name:
            f.write("%s\n" % item)

    with open(month_text_file, 'w') as f:
        for item in month_column:
            f.write("%s\n" % item)


if __name__ == '__main__':
    aggregate_data()

import win32com.client

def aggregate_data():
    # Path to the uploaded Excel file
    file_path = r'D:\Projects\test\test.xlsm'

    # Create an instance of the Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Optional: make Excel visible (for debugging purposes)
    excel.Visible = False

    # Open the Excel file
    workbook = excel.Workbooks.Open(file_path)

    # Access the 'Bills List' sheet
    sheet = workbook.Sheets('Bills List')

    # Initialize an empty list to store client names
    bill_no = []
    company_name = []
    month_column = []

    # Define the column indices (assuming 'Name' is in column 1 and 'Client' is in column 2)
    name_column = 1
    bill_column = 2
    company_column = 3
    monthly_column = 8

    # Loop through the rows in the sheet (assuming the data starts from row 2)
    row = 1724
    while True:
        try:
            # Read the value of the 'Name' column
            name_value = sheet.Cells(row, name_column).Value

            # Break the loop if we have reached the end of the data
            if name_value is None:
                break

            # Check if the 'Name' column contains 'YJV'
            if 'YJV' in name_value:
                # Read the value of the 'Client' column
                client_bill = sheet.Cells(row, bill_column).Value
                client_name = sheet.Cells(row, company_column).Value
                client_month = sheet.Cells(row, monthly_column).Value

                # Check if the client name is a number
                if isinstance(client_bill, (int, float)):
                    # Append the client name to the list
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

    bill_no = [int(num) for num in bill_no] 
    # Print the list of client names
    # print(bill_no)
    # print(company_name)
    # print(month_column)   

    # with open('bill_no.txt', 'w') as f:
    #     for item in bill_no:
    #         # Write each item on a new line
    #         f.write("%s\n" % item)

    # # Do the same for the company_name list
    # with open('company_name.txt', 'w') as f:
    #     for item in company_name:
    #         f.write("%s\n" % item)

    with open('month.txt', 'w') as f:
        for item in month_column:
            f.write("%s\n" % item)


if __name__ == '__main__':
    aggregate_data()
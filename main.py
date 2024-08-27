# main.py
import time
import file_import

month_file = file_import.month_file

# Import aggregate here, after it's needed
import aggregate
bill_sheet = str(input("Enter the name of the bill sheet: "))
row = int(input("Enter the row number: "))
aggregate.aggregate_data(bill_sheet, row)

time.sleep(2)

import test
test.format_dates(month_file)

time.sleep(2)

# Import print1 here, after it's needed
import print1
sheet_name = str(input("Enter the name of the sheet: "))
print1.process_excel_data(sheet_name)
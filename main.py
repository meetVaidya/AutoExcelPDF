# main.py
import time
import file_import

month_file = file_import.month_file

# Import aggregate here, after it's needed
import aggregate
aggregate.aggregate_data()

time.sleep(2)

import test
test.format_dates(month_file)

time.sleep(2)

# Import print1 here, after it's needed
import print1
print1.process_excel_data()
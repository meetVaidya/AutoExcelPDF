import pywintypes
import datetime
import numpy as np

def format_dates(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()

    # Extract the dates from the file content
    dates = [line.strip() for line in lines]

    # Convert the dates to datetime objects
    datetime_objects = []
    for date in dates:
        if date != 'None':
            datetime_objects.append(datetime.datetime.fromisoformat(date))
        else:
            datetime_objects.append(np.nan)

    # Convert the datetime objects to formatted strings
    formatted_dates = []
    for date in datetime_objects:
        if isinstance(date, float):
            formatted_dates.append(str(date) + ',')
        else:
            formatted_dates.append(date.strftime('%B %Y, '))

    with open(file_path, 'w') as file:
        file.writelines(formatted_dates)

    with open(file_path, 'r') as file:
        text = file.readlines()

    formatted_text = ''.join(text).replace(",", "\n")

    with open(file_path, 'w') as file:
        file.write(formatted_text)
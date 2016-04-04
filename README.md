# rename-folders-python
Rename multiple folders witha specific naming convention using data from XLSX sheet
import os
import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
import string
import subprocess
import re
import datetime

# Get directories
directories = [name for name in os.listdir('.')]

# Load Workbook
wb = openpyxl.load_workbook('database.xlsx')

# Load first work sheet
ws = wb['Sheet1']
rows = 10
dates_used = {}
# start reading the folders and rename them
# split the folder name at the underscore "_" so as to use the timestamp portion for date and numbering of the renamed folders
for i in xrange(2, rows + 2):
    folder_to_change = [x for x in directories if x.startswith(ws['A%s' % i])]
    folder_to_change = folder_to_change[0] if len(folder_to_change) > 0 else None
    if folder_to_change:
        directories.remove(folder_to_change)
        folder_parts = folder_to_change.split("_")
        timestamp = folder_parts[1]
        ts = re.search('(....)(....)(......)', timestamp)
        date = ts.group(2) + ts.group(1)
        time = ts.group(3)
        dates_used[timestamp] = 1 if timestamp not in dates_used else dates_used[timestamp] + 1
        os.rename(folder_to_change,
                  '{0}_{1}_{2}_{3}_{4}_{5}_{6}_{7}_{8}'.format(ws['B%s' % i], ws['C%s' % i], ws['G%s' % i], date, ws['D%s' % i], dates_used[timestamp], ws['F%s' % i], ws['H%s' % i], ws['I%s' % i]))

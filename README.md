import os
import random
import string
import datetime
import re
import openpyxl
from collections import namedtuple

# Get Directories
directories = [name for name in os.listdir(r'.') if name.startswith('COM')]

# Load workbook
wb = openpyxl.load_workbook('database.xlsx')
ws = wb['Sheet1']
rows = 20
for i in xrange(2, rows + 2):
    if ws['A%s' % i].value >= "":
        folder_to_change = [x for x in directories if x.startswith((ws['A%s' % i].value) + '_')]
        folder_to_change = folder_to_change[0] if len(folder_to_change) > 0 else None
        date_used = {}
        if folder_to_change:
            directories.remove(folder_to_change)
            # print folder_to_change
            folder_parts = folder_to_change.split('_')
            comport = folder_parts[0]
            timestamp = folder_parts[1]
            # print timestamp
            ts = re.search('(....)(....)(......)', timestamp)
            date = ts.group(2) + ts.group(1)
            # print date
            time = comport
            # print time
            date_used[comport] = 1 if comport not in date_used else date_used[comport] + 1
            # print date_used
            # print folder_to_change
            os.rename (folder_to_change, '{0}_{1}_{2}_{3}_{4}_{5}_{6}_{7}_{8}'.format(ws['B%s' % i].value, ws['C%s' % i].value, ws['G%s' % i].value, date, ws['D%s' % i].value, date_used[comport], ws['F%s' % i].value, ws['H%s' % i].value, ws['I%s' % i].value))

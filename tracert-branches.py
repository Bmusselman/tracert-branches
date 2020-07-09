#! python 3

import re, os, datetime, time, openpyxl, subprocess, shutil

# excel doc path
excelSheet = r''

# load excel doc
workbook = openpyxl.load_workbook(excelSheet)

# load excel sheet
sheet = workbook['Sheet1']

# time format (YYYY-MM-DD_HOUR-MIN)
timeFormat = datetime.datetime.now().strftime('%Y-%m-%d__%H-%M')

# time of program start
beginTime = datetime.datetime.now()

# tracert command
# -h = max hop count 7 to speed up
# -d to remove host name resolution to speed up
tracert = 'tracert -d -h 7'

# change directory to where I want log saved
# r' to make it a raw string overwise get an error -- unicode escape
os.chdir(r'')

# logFile = (timeFormat + '.txt')
logFile = open(timeFormat + '.txt', 'a+')

# logFileSorted = (timeFormat + '.txt')
logFileSorted = open(timeFormat + '_SORTED.txt', 'a+')

# define list for sorting
tracertByteList = []

# core provider router IP's
# need to make a bunch of seperate strings I don't know another way
coreRouter1 = ''
coreRouter2 = ''
coreRouter3 = ''
coreRouter4 = ''


# create a function for this in the future it iterate over rows/columns given row/column
for row in range(2, sheet.max_row + 1): # for each row, starting row 2 to last
    brIP = sheet.cell(row = row, column = 2).value # sheet B = col 2, every row starting at 2
    if(brIP is None): # if row value is empty continue loop
        continue
    command = subprocess.Popen(tracert + ' ' + brIP, stdout = subprocess.PIPE, universal_newlines = True) 
    tracertByteList.append(command.stdout.read()) # adds each output to list
    command.wait() # wait for command to finish before starting next

# loops through each tracert list item, reformat from byte code to readable string
for tracertByte in tracertByteList:
    logFile.write(tracertByte.replace('\\n', '\n')) # readable

# anything that does not return a core router IP gets added to sorted list 
for tracert in tracertByteList:
    if (tracert.find(coreRouter1) != -1 or
        tracert.find(coreRouter2) != -1 or
        tracert.find(coreRouter3) != -1 or
        tracert.find(coreRouter4) != -1):
        continue
    else:
        logFileSorted.write(tracert.replace('\\n', '\n'))

# print program execution time
execTime = datetime.datetime.now() - beginTime
print('The script execution time is: ' + str(execTime))
logFile.write('\n\nThe script execution time is: ' + str(execTime))
logFileSorted.write('\n\nThe script execution time is: ' + str(execTime))
logFile.close()
logFileSorted.close()


# TODO: find if possible to make it faster
# TODO: on tracert text replace IP's with BR#

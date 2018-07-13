#!/usr/bin/env python3

"""
This program takes an export Excel file and
tabulates TC counts Open and Closed, by screening and priority,
and by date closed.
File type must be .xlsx, older .xls are not supported by openpyxl.
Date of last change:
Jonathan McDonald
Args:
Returns:
Raises:
"""

# Import and include
import openpyxl
import pprint
import datetime
# import os.path
from os.path import isfile as checkFile
# import excelBeanCounter
# from excelBeanCounter import rowCounter

# Let me know I have started
# open downloaded TSM hull export for today
# ask me the file name
##try:
##    myDay = datetime.date.today()
##    print('myDay: ' + str(myDay))
##    today = str(myDay)
##    print('today s1: ' + today)
##    today = today.replace('-', '')
##    print('today s2: ' + today)
##
##    myHullNum = input('What is the Hull Number? ')
##    curTSM_xlsx = today + '_1 LPD ' + myHullNum + ' ALL bean bu.xls'
##    print('The file name is: ' + curTSM_xlsx)
##    # Open handles and objects
##    wb = openpyxl.load_workbook(curTSM_xlsx)  # must be in same directory or bat location
##except:
##    curTSM_xlsx = input('What is the name of the TSM Excel export file? ')
##    print('The file name is: ' + curTSM_xlsx)
##    # Open handles and objects
##    wb = openpyxl.load_workbook(curTSM_xlsx)  # Excel file must be in same directory or bat location


# Row count function
def rowCount(row, bean, dt_BT=None, dt_AT=None):
    # Reset selectors booleans
    myPri_star = False
    myPri_1s = False
    myPri_1 = False
    myPri_2s = False
    myPri_2 = False
    myPri_3s = False
    myPri_other = False

    # Each row in the spreadsheet has data for one Trial Card
    dsp = sheet['A' + str(row)].value  # 'Trial Card #'
    star = sheet['B' + str(row)].value  # 'Star'
    pri = sheet['C' + str(row)].value  # 'Pri'
    safe = sheet['D' + str(row)].value  # 'Saf'
    scrn = sheet['E' + str(row)].value  # 'Scrn'
    ac1 = sheet['F' + str(row)].value  # 'Act 1'
    ac2 = sheet['G' + str(row)].value  # 'Act 2'
    stat = sheet['H' + str(row)].value  # 'Status'
# actkn not currently used
    # actkn = sheet['I' + str(row)].value  # 'Action Taken'
    dt_disc = sheet['J' + str(row)].value  # 'Date Discovered'
    # Reformat dates from Oracle(-) to Microsoft(/)
    if dt_disc is not None:
        # date_disc = datetime.date(dt_disc, '%Y-%m-%d')
        dt_disc = dt_disc.replace('-', '/')
        pass
    dt_close = sheet['K' + str(row)].value  # 'Date Closed'
    # Reformat dates from Oracle(-) to Microsoft(/)
    if dt_close is not None:
        # date_close = datetime.date(dt_close, '%Y-%m-%d')
        dt_close = dt_close.replace('-', '/')
        pass
# trial_ID and event currently not used
    # trial_ID = sheet['L' + str(row)].value  # 'Trial ID'
    # event = sheet['M' + str(row)].value  # 'Event'

    # Combine singleton values
    # May turn this off and not combine screening codes
    if (ac2 != '') and (ac2 is not None):
        scrngs = scrn + '/' + ac1 + '/' + ac2
    else:
        scrngs = scrn + '/' + ac1

    # Check for Starred Cards
    if (star == 'STAR') or (star == 'star') or (star == '*'):
        myPri_star = True
    elif (star == '') or (star is None):
        # No Starred Cards in row
        if int(pri) == 1:
            if safe == 'S':
                myPri_1s = True
            elif safe != 'S':
                myPri_1 = True
            else:
                # Un-captured value in field
                pass
        elif int(pri) == 2:
            if safe == 'S':
                myPri_2s = True
            elif safe != 'S':
                myPri_2 = True
            else:
                # Un-captured value in field
                pass
        elif int(pri) == 3:
            if safe == 'S':
                myPri_3s = True
            elif safe != 'S':
                myPri_other = True
            else:
                # Un-captured value in field
                pass
        else:
            # Un-captured value in field
            myPri_other = True
            print('Row read error Priority in not Valid. Row: ' + str(row) + ' TC Number: ' + dsp)
    else:
        # No Starred Cards in row, catch all
        pass

## Commented out to stop double counting Stared cards
##    if int(pri) == 1:
##        if safe == 'S':
##            myPri_1s = True
##        elif safe != 'S':
##            myPri_1 = True
##        else:
##            # Un-captured value in field
##            pass
##    elif int(pri) == 2:
##        if safe == 'S':
##            myPri_2s = True
##        elif safe != 'S':
##            myPri_2 = True
##        else:
##            # Un-captured value in field
##            pass
##    elif int(pri) == 3:
##        if safe == 'S':
##            myPri_3s = True
##        elif safe != 'S':
##            myPri_other = True
##        else:
##            # Un-captured value in field
##            pass
##    else:
##        # Un-captured value in field
##        myPri_other = True
##        print('Row read error Priority in not Valid. Row: ' + str(row) + ' TC Number: ' + dsp)

    # Make sure the key(s) for these dictionaries exist.
    # The .setdefault() checks if the key exists, if not it creates
    # with default passed value otherwise it will do nothing.
    # By bean type, status and priority count
    tc_Status.setdefault(bean, {})
    tc_Status[bean].setdefault(stat, {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0})
    # By bean type, status, screen and screening count
    tc_Stat_Scrn.setdefault(bean, {})
    tc_Stat_Scrn[bean].setdefault(stat, {})
    tc_Stat_Scrn[bean][stat].setdefault(scrn, {})
    tc_Stat_Scrn[bean][stat][scrn].setdefault(scrngs, {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0})

    # Check for empty date closed
    if (dt_close == '') or (dt_close is None):
        # Don't count, card still open
        # This could be used as a counting validation
        pass
    else:
        # By bean type, date closed and date closed count
        tc_DateClosed.setdefault(bean, {})
        # Check if Date Closed is already in the dictionary
        tc_DateClosed[bean].setdefault(dt_close, {})
        tc_DateClosed[bean][dt_close].setdefault('Closed Count', 0)
        # Increment the date closed count plus 1
        tc_DateClosed[bean][dt_close]['Closed Count'] += 1

        # Build days from trial
        if dt_BT is not None:
            tc_DateClosed[bean][dt_close].setdefault('Days from BT', None)
            dateClosed = datetime.datetime.strptime(dt_close, '%Y/%m/%d')
            date_BT = datetime.datetime.strptime(dt_BT, '%Y/%m/%d')
            date_Delta = (dateClosed - date_BT).days
            tc_DateClosed[bean][dt_close]['Days from BT'] = date_Delta
        else:
            # BT date not given
            pass
        if dt_AT is not None:
            tc_DateClosed[bean][dt_close].setdefault('Days from AT', None)
            dateClosed = datetime.datetime.strptime(dt_close, '%Y/%m/%d')
            date_AT = datetime.datetime.strptime(dt_AT, '%Y/%m/%d')
            date_Delta = (dateClosed - date_AT).days
            tc_DateClosed[bean][dt_close]['Days from AT'] = date_Delta
        else:
            # AT date not given
            pass

    # Update the Status and Screening dictionary's counts
    tc_Status[bean][stat]['Total'] += 1
    tc_Stat_Scrn[bean][stat][scrn][scrngs]['Total'] += 1
    # Update the Priority counts
    if myPri_star is True:
        tc_Status[bean][stat]['Pri STARRED'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri STARRED'] += 1
    elif myPri_1s is True:
        tc_Status[bean][stat]['Pri 1S'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 1S'] += 1
    elif myPri_1 is True:
        tc_Status[bean][stat]['Pri 1'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 1'] += 1
    elif myPri_2s is True:
        tc_Status[bean][stat]['Pri 2S'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 2S'] += 1
    elif myPri_2 is True:
        tc_Status[bean][stat]['Pri 2'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 2'] += 1
    elif myPri_3s is True:
        tc_Status[bean][stat]['Pri 3S'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 3S'] += 1
    elif myPri_other is True:
        tc_Status[bean][stat]['Pri OTHER'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri OTHER'] += 1
    else:
        print('Row read error with Status of Open, Priority in not Valid. '
                  'Row: ' + str(row) + ' TC Number: ' + dsp)
        pass


# Get the current Hull Number
myHullNum = input('What is the Hull Number?\n'
                  'Example: 17: ')

# Build Events List
runEvents = input('\nRun reports by Event?\n'
                  'Y or N: ')
runEvents = runEvents.upper()
if runEvents == 'Y':
    userEvents = input('What Events are being counted?\n'
                  'Example: AT,BT,FCT: ')
    userEvents = userEvents.upper()
    events = userEvents
    curReportEvents = [item for item in userEvents.split(',') if item.strip()]
else:
    # Presumed no, pass
    pass

# Build INSURV List
runTID = input('\nRun reports by Trial ID?\n'
               'Y or N: ')
runTID = runTID.upper()
if runTID == 'Y':
    userTID = input('What INSURV Events are being counted?\n'
                    'Example: C,F or C+: ')
    userTID = userTID.upper()
    trial_ID = userTID
    curReportTrial_ID = [item for item in userTID.split(',') if item.strip()]
else:
    # Presumed no, pass
    pass

# Get BT Date
runBT = input('\nCount closure from BT Trial?\n'
                  'Y or N: ')
runBT = runBT.upper()
if runBT == 'Y':
    userBT = input('What is the date of the BT Trial?\n'
                   'Example: yyyy/mm/dd: ')
    print(userBT)
else:
    # Presumed no
    userBT = None

# Get AT Date
runAT = input('\nCount closure from AT Trial?\n'
                  'Y or N: ')
runAT = runAT.upper()
if runAT == 'Y':
    userAT = input('What is the date of the AT Trial?\n'
                  'Example: yyyy/mm/dd: ')
    print(userAT)
else:
    # Presumed no
    userAT = None

# Get and test the name of Excel file
fileNotFound = True
while fileNotFound is True:
    curTSM_xlsx = input('\nWhat is the name of the TSM Excel export file? ')
    # print('The file name is: ' + curTSM_xlsx)
    if checkFile(curTSM_xlsx):
        fileNotFound = False
    else:
        print('File not found, please re-enter the file name.\n'
              'File must be in the Python root directory.\n')

# Open handles and objects
# Excel file must be in same directory or bat location
wb = openpyxl.load_workbook(curTSM_xlsx)

print('\nOpening workbook...')
try:
    # The default is 'TSM EXPORT'
    sheet = wb.get_sheet_by_name('TSM EXPORT')
except:
    mySheetName = input('\nWhat is the name of the sheet? ')
    sheet = wb.get_sheet_by_name(mySheetName)
    print(sheet)

# Initialize the empty dictionary's
tc_Status = {}  # tc_Status{[stat]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}
tc_Stat_Scrn = {}  # tc_Stat_Scrn{[stat]: {[scrn]: {[scrngs]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}}}
tc_DateClosed = {}  # tc_DateClosed{[dt_close]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}

# Let me know I am reading and looping through file
print('Reading rows...')

for row in range(2, sheet.max_row + 1):
    # Check for Header, Last Row or empty values
    if ((sheet['A' + str(row)].value != 'Trial Card No') or
       (sheet['A' + str(row)].value != 'Trial Card #') or
       (sheet['A' + str(row)].value != 'FOR OFFICIAL USE ONLY') or
       (sheet['A' + str(row)].value != '')):
        # Check for empty TC Status values
        if sheet['H' + str(row)].value != '':
            # Check if Event is in current report range
            if sheet['M' + str(row)].value in curReportEvents:
                # Process row
                rowCount(row, events, userBT, userAT)
            else:
                # Current row is not in the current report events list
                pass
            if sheet['L' + str(row)].value in curReportTrial_ID:
                # Process row
                rowCount(row, trial_ID, userBT, userAT)
            else:
                # Current row is not in the current report events list
                pass

    elif ((sheet['A' + str(row)].value == 'Trial Card No') or
          (sheet['A' + str(row)].value == 'Trial Card #')):
        print('Header row found, pass.')
        pass
    elif sheet['A' + str(row)].value == 'FOR OFFICIAL USE ONLY':
        print('Exported fouo found, pass')
        pass
    elif sheet['A' + str(row)].value == '':
        print('Empty cell found, pass')
        pass
    else:
        # Un-captured value in field
        pass


# Open a  file and write the contents of the dictionaries to it.
print('Writing results...')

resultFile = open('LPD ' + myHullNum + ' Bean Data.py', 'w')
# pprint is python print, and formats as valid python code and structure
resultFile.write('tc_Status = ' + pprint.pformat(tc_Status))
resultFile.write('\ntc_Stat_Scrn = ' + pprint.pformat(tc_Stat_Scrn))
resultFile.write('\ntc_DateClosed = ' + pprint.pformat(tc_DateClosed))
# Dispose of object
resultFile.close()

# Completed
print('Bean dictionary completed.')

# End Of Line
hold = input('Press any key to exit.')

# End Of Line
print('Good Bye')

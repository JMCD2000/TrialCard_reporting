#!/usr/bin/env python3

"""
This program takes an export Excel file and
tabulates TC counts Open and Closed, by screening and priority,
and by date closed.
File type must be .xlsx, older .xls are not supported by openpyxl.
Date of last change 8/10/2017
Jonathan McDonald
Args:
Returns:
Raises:
"""

# import and include
import openpyxl, pprint, datetime
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
def rowCount(row, bean):
    # Reset selectors bools
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
    actkn = sheet['I' + str(row)].value  # 'Action Taken'
# TODO reformat dates from Oracle(-) to Microsoft(/)
    dt_disc = sheet['J' + str(row)].value  # 'Date Discovered'
    dt_close = sheet['K' + str(row)].value  # 'Date Closed'
    trial_ID = sheet['L' + str(row)].value  # 'Trial ID'
    Event = sheet['M' + str(row)].value  # 'Event'

   # Combine singleton values
   # May turn this off and not combine screening codes
#    if (len(ac2) > 0) or (ac2 != '') or (ac2 is not None):
    if (ac2 != '') and (ac2 is not None):
        scrngs = scrn + '/' + ac1 + '/' + ac2
    else:
        scrngs = scrn + '/' + ac1

    # Check for Starred Cards
    if (star == 'STAR') or (star == 'star') or (star == '/*'):
        myPri_star = True
    elif star == '':
        # No Starred Cards in row
        pass
    else:
        # No Starred Cards in row, catch all
        pass

    if int(pri) == 1:
        if safe == 'S':
            myPri_1s = True
        elif safe != 'S':
            myPri_1 = True
        else:
            # Uncaptured value in field
            pass
    elif int(pri) == 2:
        if safe == 'S':
            myPri_2s = True
        elif safe != 'S':
            myPri_2 = True
        else:
            # Uncaptured value in field
            pass
    elif int(pri) == 3:
        if safe == 'S':
            myPri_3s = True
        elif safe != 'S':
            myPri_other = True
        else:
            # Uncaptured value in field
            pass
    else:
        # Uncaptured value in field
        myPri_other = True
        print('Row read error Priority in not Valid. Row: ' + str(row) + ' TC Number: ' + dsp)

    # Makesure the key(s) for these dictionaries exist. The .setdefault() checks if the key exsists, if not it creates with default passed value otherwise it will do nothing.
    # By bean type, status and priority count
    tc_Status.setdefault(bean, {})
    tc_Status[bean].setdefault(stat, {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0})
    # By bean type, status, screen and screening count
    tc_Stat_Scrn.setdefault(bean, {})
    tc_Stat_Scrn[bean].setdefault(stat, {})
    tc_Stat_Scrn[bean][stat].setdefault(scrn, {})
    tc_Stat_Scrn[bean][stat][scrn].setdefault(scrngs, {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0})

    # Check for empty date closed
    if (dt_close == '') or (dt_close is None):
        # Don't count, card still open
        # This could be used as a counting validation
        pass
    else:
        # By bean type, date closed and date closed count
        tc_DateClosed.setdefault(bean, {})
        # Check if Date Closed is allready in the dictionary
        tc_DateClosed[bean].setdefault(dt_close, 0)
        # Increment the date closed count plus 1
        tc_DateClosed[bean][dt_close] += 1

    # Update the Status and Screening dictionarys counts
    tc_Status[bean][stat]['Total'] += 1
    tc_Stat_Scrn[bean][stat][scrn][scrngs]['Total'] +=  1
    # Update the Priority counts
    if myPri_star == True:
        tc_Status[bean][stat]['Pri STARRED'] +=  1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri STARRED'] +=  1            
    elif myPri_1s == True:
        tc_Status[bean][stat]['Pri 1S'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 1S'] += 1
    elif myPri_1 == True:
        tc_Status[bean][stat]['Pri 1'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 1'] += 1
    elif myPri_2s == True:
        tc_Status[bean][stat]['Pri 2S'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 2S'] += 1
    elif myPri_2 == True:
        tc_Status[bean][stat]['Pri 2'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 2'] += 1
    elif myPri_3s == True:
        tc_Status[bean][stat]['Pri 3S'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri 3S'] += 1
    elif myPri_other == True:
        tc_Status[bean][stat]['Pri OTHER'] += 1
        tc_Stat_Scrn[bean][stat][scrn][scrngs]['Pri OTHER'] += 1
    else:
        print('Row read error with Status of Open, Priority in not Valid. Row: ' + str(row) + ' TC Number: ' + dsp)
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
    #print(userEvents)
    userEvents = userEvents.upper()
    events = userEvents
    #print(userEvents)
    curReportEvents = [ item for item in userEvents.split(',') if item.strip()]    
    #print(curReportEvents)
else:
    #presumed no, pass
    pass

# Build INSURV List
runTID = input('\nRun reports by Trial ID?\n'
               'Y or N: ')
runTID = runTID.upper()
if runTID == 'Y':
    userTID = input('What INSURV Events are being counted?\n'
                    'Example: C,F or C+: ')
    #print(userTID)
    userTID = userTID.upper()
    trial_ID = userTID
    #print(userTID)
    curReportTrial_ID = [ item for item in userTID.split(',') if item.strip()]
    #print(curReportTrial_ID)
else:
    #presumed no, pass
    pass

# Get and test the name of Excel file
fileNotFound = True
while fileNotFound is True:
    curTSM_xlsx = input('\nWhat is the name of the TSM Excel export file? ')  # file name
    #print('The file name is: ' + curTSM_xlsx)
    if checkFile(curTSM_xlsx):
        fileNotFound = False
    else:
        print('File not found, please re-enter the file name.\n'
              'File must be in the Python root directory.\n')

# Open handles and objects
# Excel file must be in same directory or bat location
wb = openpyxl.load_workbook(curTSM_xlsx)
# try:
print('\nOpening workbook...')
try:
    sheet = wb.get_sheet_by_name('TSM EXPORT')
except:
    mySheetName = input('\nWhat is the name of the sheet? ')  # The default is 'TSM EXPORT'
    sheet = wb.get_sheet_by_name(mySheetName)
    print(sheet)
    
# Inilize the empty dictionarys
tc_Status = {}  # tc_Status{[stat]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}
tc_Stat_Scrn = {}  # tc_Stat_Scrn{[stat]: {[scrn]: {[scrngs]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}}}
tc_DateClosed = {}  # tc_DateClosed{[dt_close]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}

# Let me know I am reading and looping through file
print('Reading rows...')

for row in range(2,sheet.max_row + 1):
    # Check for Header, Last Row or empty values
    if (sheet['A' + str(row)].value != 'Trial Card No') or (sheet['A' + str(row)].value != 'Trial Card #') or (sheet['A' + str(row)].value != 'FOR OFFICIAL USE ONLY') or (sheet['A' + str(row)].value != ''):
        # Check for empty TC Status values
        if sheet['H' + str(row)].value != '':
            # Check if Event is in current report range
            if sheet['M' + str(row)].value in curReportEvents:
                # Process row
                # Call Row Count Function (row, Events)
                rowCount(row, events)
            else:
                # Current row is not in the current report events list
                pass
# TODO search in a list
            if sheet['L' + str(row)].value in curReportTrial_ID:
                # Process row
                # Call Row Count Function (row, Trial_ID)
                rowCount(row, trial_ID)
            else:
                # Current row is not in the current report events list
                pass
    elif (sheet['A' + str(row)].value == 'Trial Card No') or (sheet['A' + str(row)].value == 'Trial Card #'):
        print('Header row found, pass.')
        pass
    elif sheet['A' + str(row)].value == 'FOR OFFICIAL USE ONLY':
        print('Exported fouo found, pass')
        pass
    elif sheet['A' + str(row)].value == '':
        print('Empty cell found, pass')
        pass
    else:
        # Uncaptured value in field
        pass
# except:
#    print('An uncontrolled error occured.')
    
       
# Open a  file and write the contents of the dictionaries to it.
print('Writing results...')
# Build file header
    # date
    # hull number
    
resultFile = open('LPD ' + myHullNum + ' Bean Data.py', 'w')
# pprint is python print, and formats as valid python code and structure
resultFile.write('tc_Status = ' + pprint.pformat(tc_Status))
resultFile.write('\ntc_Stat_Scrn = ' + pprint.pformat(tc_Stat_Scrn))
resultFile.write('\ntc_DateClosed = ' + pprint.pformat(tc_DateClosed))
# dispose of object
resultFile.close()
# tell me I'm done
print('Bean dictionary completed.')

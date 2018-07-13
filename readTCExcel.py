#! python3
# readTCExcel.py - Tabulates TC counts Open and Closed, By screening and priority
# date of last change 3/13/2017

#import and include
import openpyxl, pprint, datetime

# Let me know I have started
# open downloaded TSM hull export for today
# ask me the file name
try:
    myDay = datetime.date.today()
    #print('myDay: ' + str(myDay))
    today = str(myDay)
    #print('today s1: ' + today)
    today = today.replace('-', '')
    #print('today s2: ' + today)    

    myHullNum = input('What is the Hull Number? ')    
    curTSM_xlsx = today + '_1 LPD ' + myHullNum + ' ALL bean bu.xlsx'
    print('The file name is: ' + curTSM_xlsx)
    # Open handles and objects
    wb = openpyxl.load_workbook(curTSM_xlsx) #must be in same directory or bat location
except:
    curTSM_xlsx = input('What is the name of the TSM Excel export file? ')
    print('The file name is: ' + curTSM_xlsx)
    # Open handles and objects
    wb = openpyxl.load_workbook(curTSM_xlsx) #must be in same directory or bat location

##curTSM_xlsx = input('What is the name of the TSM Excel export file? ')
##print('The file name is: ' + curTSM_xlsx)
### Open handles and objects
##wb = openpyxl.load_workbook(curTSM_xlsx) #must be in same directory or bat location

print('Opening workbook...')

try:
    sheet = wb.get_sheet_by_name('TSM EXPORT')
except:
    mySheetName = input('What is the name of the sheet? ') #default is 'TSM EXPORT'
    sheet = wb.get_sheet_by_name(mySheetName)
    
# inilize the empty dictionarys
tc_Status_INSURV = {} # tc_Status{[stat]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}
tc_Stat_Scrn_INSURV = {} # tc_Stat_Scrn{[stat]: {[scrn]: {[scrngs]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}}}
tc_Status_All = {} # tc_Status{[stat]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}
tc_Stat_Scrn_All = {} # tc_Stat_Scrn{[stat]: {[scrn]: {[scrngs]: {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0}}}}

#TODO: Fill in TC dicts with TC Data.
# Let me know I am reading and looping through file
print('Reading rows...')
for row in range(2,sheet.max_row + 1):

    # Each row in the spreadsheet has data for one Trial Card
    dsp = sheet['A' + str(row)].value
    star = sheet['B' + str(row)].value
    pri = sheet['C' + str(row)].value
    safe = sheet['D' + str(row)].value
    scrn = sheet['E' + str(row)].value
    ac1 = sheet['F' + str(row)].value
    ac2 = sheet['G' + str(row)].value
    stat = sheet['H' + str(row)].value
    actkn = sheet['I' + str(row)].value
    dt_disc = sheet['J' + str(row)].value
    dt_close = sheet['K' + str(row)].value
    trial_ID = sheet['L' + str(row)].value
    Event = sheet['M' + str(row)].value

    if dsp != 'Trial Card No':
        
        if dsp != 'FOR OFFICIAL USE ONLY':
            
            if stat != '':
                
                # Reset selectors bools
                myPri_star = False
                myPri_1s = False
                myPri_1 = False
                myPri_2s = False
                myPri_2 = False
                myPri_3s = False
                myPri_other = False
                
                
               # Combined singleton values
                if len(ac2)>0:
                    scrngs = scrn + '/' + ac1 + '/' + ac2
                else:
                    scrngs = scrn + '/' + ac1

                #if star == '/*' or 'STAR' or 'Star':
                if star == 'STAR':
                    myPri_star = True
                else:
                    if int(pri) == 1:
                        if safe == 'S':
                           myPri_1s = True
                        else:
                            myPri_1 = True
                    elif int(pri) == 2:
                        if safe == 'S':
                            myPri_2s = True
                        else:
                            myPri_2 = True
                    else:
                        if int(pri) == 3:
                            if safe == 'S':
                                myPri_3s = True
                            else:
                                myPri_other = True
                        else:
                            myPri_other = True
                            print('Row read error Priority in not Valid. Row: ' + str(row) + ' TC Number: ' + dsp)

                # Makesure the key for this Dict exists. The .setdefault() checks if it exsists, if not it creates with default passed value otherwise it will do nothing.
                # By status and priority count
                tc_Status_All.setdefault(stat, {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0})
                # By status and screen and screening count
                tc_Stat_Scrn_All.setdefault(stat, {})
                tc_Stat_Scrn_All[stat].setdefault(scrn, {})
                tc_Stat_Scrn_All[stat][scrn].setdefault(scrngs, {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total': 0})

                # Update the Priority and Status dictionary counts
                tc_Status_All[stat]['Total'] += 1
                tc_Stat_Scrn_All[stat][scrn][scrngs]['Total'] +=  1
                if myPri_star == True:
                    tc_Status_All[stat]['Pri STARRED'] +=  1
                    tc_Stat_Scrn_All[stat][scrn][scrngs]['Pri STARRED'] +=  1            
                elif myPri_1s == True:
                    tc_Status_All[stat]['Pri 1S'] += 1
                    tc_Stat_Scrn_All[stat][scrn][scrngs]['Pri 1S'] += 1
                elif myPri_1 == True:
                    tc_Status_All[stat]['Pri 1'] += 1
                    tc_Stat_Scrn_All[stat][scrn][scrngs]['Pri 1'] += 1
                elif myPri_2s == True:
                    tc_Status_All[stat]['Pri 2S'] += 1
                    tc_Stat_Scrn_All[stat][scrn][scrngs]['Pri 2S'] += 1
                elif myPri_2 == True:
                    tc_Status_All[stat]['Pri 2'] += 1
                    tc_Stat_Scrn_All[stat][scrn][scrngs]['Pri 2'] += 1
                elif myPri_3s == True:
                    tc_Status_All[stat]['Pri 3S'] += 1
                    tc_Stat_Scrn_All[stat][scrn][scrngs]['Pri 3S'] += 1
                elif myPri_other == True:
                    tc_Status_All[stat]['Pri OTHER'] += 1
                    tc_Stat_Scrn_All[stat][scrn][scrngs]['Pri OTHER'] += 1
                else:
                    print('Row read error with Status of Open, Priority in not Valid. Row: ' + str(row) + ' TC Number: ' + dsp)

            elif stat == '':
                print('exported empty status found')

        elif dsp == 'FOR OFFICIAL USE ONLY':
            print('exported fouo found')

    elif dsp == 'Trial Card No':
        print('header row found')
   
        
# Open a new text file and write the contents of countyData to it.
print('Writing results...')
# Build file header
    # date
    # hull number
    
resultFile = open('LPD Bean Data.py', 'w')
# pprint is python print, and fromats as valid python code and structure
resultFile.write('tc_Status_All = ' + pprint.pformat(tc_Status_All))
resultFile.write('\ntc_Stat_Scrn_All = ' + pprint.pformat(tc_Stat_Scrn_All))
# dispose of object
resultFile.close()
# tell me I'm done
print('Done.')

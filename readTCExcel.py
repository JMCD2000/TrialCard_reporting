#! python3
# readTCExcel.py - Tabulates TC counts Open and Closed, By screening and priority
# date of last change 3/9/2017

#import and include
import openpyxl, pprint

# Let me know I have started
# open downloaded TSM hull export for today
# ask me the file name
curTSM_xlsx = input('What is the name of the TSM Excel export file? ')
print('Opening workbook...')
# Open handles and objects
wb = openpyxl.load_workbook(curTSM_xlsx) #must be in same directory or bat location
mySheetName = input('What is the name of the sheet? ') #default is 'TSM EXPORT'
sheet = wb.get_sheet_by_name(mySheetName)
# inilize the dictionarys
tc_Status = {'Open': 0, 'Closed': 0, 'Total': 0}
tc_Pri_Open = {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}
tc_Pri_Closed = {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0,   'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}
# inilize the empty dictionarys
sf_Respon = {} # {'SCRNGS': {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}}
gov_Deferd = {} #  {'SCRNGS': {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}}
gov_Invest = {} #  {'SCRNGS': {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}}
con_Invest = {} #  {'SCRNGS': {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}}
con_Respon = {} #  {'SCRNGS': {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}}
gov_Respon = {} #  {'SCRNGS': {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}}
tc_Closed_Scrn = {} #  {'SCRNGS': {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 2': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0}}

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

    # Combined singleton values
    if len(ac2)>0:
        scrngs = scrn + '//' + ac1 + '//' + ac2
    else:
        scrngs = scrn + '//' + ac1

    # TODO: Make this a list myPri(0,0,0,0,0,0,0)
        # myPri[0] = star
    if star == '*' or 'STAR':
        # myPri(1,0,0,0,0,0,0)
        pri_star = 1
        pri_1s = 0
        pri_1 = 0
        pri_2s = 0
        pri_2 = 0
        pri_3s = 0
        pri_other = 0

    else:
        if pri == 1:
            if safe == 'S':
                # myPri(0,1,0,0,0,0,0)
                pri_star = 0
                pri_1s = 1
                pri_1 = 0
                pri_2s = 0
                pri_2 = 0
                pri_3s = 0
                pri_other = 0
            else:
                # myPri(0,0,1,0,0,0,0)
                pri_star = 0
                pri_1s = 0
                pri_1 = 1
                pri_2s = 0
                pri_2 = 0
                pri_3s = 0
                pri_other = 0
        elif pri == 2:
            if safe == 'S':
                # myPri(0,0,0,1,0,0,0)
                pri_star = 0
                pri_1s = 0
                pri_1 = 0
                pri_2s = 1
                pri_2 = 0
                pri_3s = 0
                pri_other = 0
            else:
                # myPri(0,0,0,0,1,0,0)
                pri_star = 0
                pri_1s = 0
                pri_1 = 0
                pri_2s = 0
                pri_2 = 1
                pri_3s = 0
                pri_other = 0
        else:
            if pri == 3:
                if safe == 'S':
                    # myPri(0,0,0,0,0,1,0)
                    pri_star = 0
                    pri_1s = 0
                    pri_1 = 0
                    pri_2s = 0
                    pri_2 = 0
                    pri_3s = 1
                    pri_other = 0
                else:
                    # myPri(0,0,0,0,0,0,1)
                    pri_star = 0
                    pri_1s = 0
                    pri_1 = 0
                    pri_2s = 0
                    pri_2 = 0
                    pri_3s = 0
                    pri_other = 1
            else:
                # myPri(0,0,0,0,0,0,1) # Pri 4 or N
                pri_star = 0
                pri_1s = 0
                pri_1 = 0
                pri_2s = 0
                pri_2 = 0
                pri_3s = 0
                pri_other = 1

# TODO: convert this to a for each in list with a list auto inc.
    if stat == 'O':
        tc_Status['Open'] += 1
        tc_Status['Total'] += 1
        # tc_Pri_Open['Pri STARRED'] +=  myPri[0]
        tc_Pri_Open['Pri STARRED'] +=  pri_star
        tc_Pri_Open['Pri 1S'] += pri_1s
        tc_Pri_Open['Pri 1'] += pri_1
        tc_Pri_Open['Pri 2S'] += pri_2s
        tc_Pri_Open['Pri 2'] += pri_2
        tc_Pri_Open['Pri 3S'] += pri_3s
        tc_Pri_Open['Pri OTHER'] += pri_other
        tc_Pri_Open['Total Open'] += 1

    elif stat == 'X':
        tc_Status['Closed'] += 1
        tc_Status['Total'] += 1
        # tc_Pri_Closed['Pri STARRED'] +=  myPri[0]
        tc_Pri_Closed['Pri STARRED'] +=  pri_star
        tc_Pri_Closed['Pri 1S'] += pri_1s
        tc_Pri_Closed['Pri 1'] += pri_1
        tc_Pri_Closed['Pri 2S'] += pri_2s
        tc_Pri_Closed['Pri 2'] += pri_2
        tc_Pri_Closed['Pri 3S'] += pri_3s
        tc_Pri_Closed['Pri OTHER'] += pri_other
        tc_Pri_Closed['Total Closed'] += 1

    else:
        #skip to my lou my darling

   
    
    # Makesure the key for this Dict exists.
        # .setdefault() checks if it exsists, if not it creates with default passed value
        # otherwise it will do nothing.
    sf_Respon.setdefault(scrngs, {})
    # Make sure the key for this county in this state exists.
    countyData[state].setdefault(scrngs, : {'Pri STARRED': 0, 'Pri 1S': 0, 'Pri 1': 0, 'Pri 2S': 0, 'Pri 3S': 0, 'Pri OTHER': 0, 'Total Open': 0})
    # Each row represents one census tract, so increment by one.
    countyData[state][county]['tracts'] += 1
    # Increase the county pop by the pop in this census tract.
    countyData[state][county]['pop'] =+ int(pop)

# TODO: Open a new text file and write the contents of TC Data to it.
# Open a new text file and write the contents of countyData to it.
print('Writing results...')
resultFile = open('census2010.py', 'w')
# pprint is python print, and fromats as valid python code and structure
resultFile.write('allData = ' + pprint.pformat(countyData))
# dispose of object
resultFile.close()
# tell me I'm done
print('Done.')

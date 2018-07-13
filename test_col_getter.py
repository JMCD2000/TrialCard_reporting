# Column to variable reference value
# Excel column header value : defined variable name
myColumnDict = { 
 'Trial Card #': 'dsp',
 'Star': 'star',
 'Pri': 'pri',
 'Priority': 'pri',
 'Safe': 'safe',
 'S': 'safe',
 's': 'safe',
 'Saf': 'safe',
 'Scrn': 'scrn',
 'Act 1': 'ac1',
 'Act 2': 'ac2',
 'Status': 'stat',
 'Action Taken': 'actkn',
 'Date Discovered': 'dt_disc',
 'Date Closed': 'dt_close',
 'Trial ID': 'trial_ID',
 'Event': 'Event',
 'Aliases': 'dsp',
}
# Hard coded values mapping
def getColumnMap(cellHead):
    if cellHead == 'Star':
        return 'star' #from tblTSM_TC
    elif cellHead == 'Priority':
        return 'priority'
    elif cellHead == 'Safety':
        return 'safety'
    else:
        print('column header not in list')
        return 'None'

# get mapping from dictionary myColumnDict
def getColumnDict(cellHead):
    if myColumnDict.get(cellHead, default=True) == True:
        if myColumnDict.get(cellHead, default=None) == None:
            def_var_name = 'not found'
        else:
            def_var_name = myColumnDict.get(cellHead, default=None)
    else:
        myColumnDict.get(cellHead, default=None)
        def_var_name = 'not found'
            
   
	return def_var_name
	
	
	
myColVal = input('enter a column value: ')

print(getColumnMap(myColVal))
print(getColumnDict(myColVal))


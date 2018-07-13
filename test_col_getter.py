myColumnDict = {'Star': 'star', 'Priority': 'priority', 'Pri': 'priority', 'Safe': 'safety', 'S': 'safety', 's': 'safety'}

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

def getColumnDict(cellHead):
    if myColumnDict.get(cellHead, default=True) == True:
        if myColumnDict.get(cellHead, default=None) == None:
            return 'not found'
        else:
            return myColumnDict.get(cellHead, default=None)
    else:
        myColumnDict.get(cellHead, default=None)
        return 'not found'
            
   

myColVal = input('enter a column value: ')

print(getColumnMap(myColVal))
print(getColumnDict(myColVal))


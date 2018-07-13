#! python3
"""Return  file type and name structure."""
# File Name: getMyFileType.py
# Date of last edit: 7/12/2018
# This program module is checking for naming format type. It will
# return True if matches a pattern and the pattern number. It will
# return False if no match and fnp0

import re


def myFNType(pFile):
    """
    myFNType -> test As Boolean, fnt As String
    This takes a passed in file name to determine slicing pattern
    """
    # Check for left trimmed file name
    testF = pFile[0:2]

    if testF == (r'.\\'):
        newfile = pFile[2:]
        #print('Had to trim file name: ' + pFile + '\nTo new name: ' + newfile)
    else:
        newfile = pFile
        #print('No name modifications: ' + newfile)

    # Build regular expressions
    # This is where the Beans have been ordered with a prefixing parenthesis
        # VBA: myfileNameVar1 = "(??)_LPD" & hullNum & "Bean(DATA)*" 'two digit serial order
    # fnp1
    myFNRegex1 = re.compile(r'\(\d{2}\)_LPD\d{2}Bean(DATA)(.*)')  # (00)_LPD17Bean(DATA)01.15.2008.xlsx
        # VBA: myfileNameVar2 = "(?)_LPD" & hullNum & "Bean(DATA)*" 'one digit serial order
    # fnp2
    myFNRegex2 = re.compile(r'\(\d\)_LPD\d{2}Bean(DATA)(.*)')  # (0)_LPD17Bean(DATA)01.15.2008.xlsx

    # This is for the beans that are un-ordered
        # VBA: myfileNameVar3 = "LPD" & hullNum & "Bean(DATA)(FCT)*"
    # fnp3
    myFNRegex3 = re.compile(r'LPD\d\dBean(DATA)(FCT)(.*)')  # LPD17Bean(DATA)(FCT)01.15.2008.xlsx
        # VBA: myfileNameVar4 = "LPD" & hullNum & "Bean(DATA)(INSURV)*"
    # fnp4
    myFNRegex4 = re.compile(r'LPD\d\dBean(DATA)(INSURV)(.*)')  # LPD17Bean(DATA)(INSURV)01.15.2008.xlsx
        # VBA: myfileNameVar5 = "LPD" & hullNum & "Bean(DATA)*" 'Must be last for this name series
    # fnp5
    myFNRegex5 = re.compile(r'LPD\d\dBean(DATA)(.*)')  # LPD17Bean(DATA)01.15.2008.xlsx

    # This is for the TSM export file
        # VBA: myfileNameVar6 = "????????_? LPD nu *"
    # fnp6
    myFNRegex6 = re.compile(r'\d{8}_\d LPD nu ALL(.*)')  # 20161114_1 LPD nu ALL bean bu.xls
        # VBA: myfileNameVar7 = "????????_? LPD " & hullNum & " *"
    # fnp7
    myFNRegex7 = re.compile(r'\d{8}_\d LPD \d{2} ALL(.*)')  # 20161114_1 LPD 17 ALL bean bu.xls
        # unsupported by VBA
    # fnp8
    myFNRegex8 = re.compile(r'\d{8}_\d LPD \d{2} INSURV(.*)')  # 20161114_1 LPD 17 INSURV bean bu.xls
        # unsupported by VBA
    # fnp9
    myFNRegex9 = re.compile(r'\d{8}_\d LPD nu INSURV(.*)')  # 20161114_1 LPD nu INSURV bean bu.xls
        # unsupported by VBA
    # fnp10
    myFNRegex10 = re.compile(r'\d{8}_LPD_\d\d_Bean_B(.*)')  # 20161114_LPD_17_Bean_Burn.xls

    # Look for a pattern match
    # test = False # as default
    # fnt = None # as default
    if (myFNRegex1.search(newfile) is None) is False:
        test = True
        fnt = 'fnp1'
    elif (myFNRegex2.search(newfile) is None) is False:
        test = True
        fnt = 'fnp2'
    elif (myFNRegex3.search(newfile) is None) is False:
        test = True
        fnt = 'fnp3'
    elif (myFNRegex4.search(newfile) is None) is False:
        test = True
        fnt = 'fnp4'
    elif (myFNRegex5.search(newfile) is None) is False:
        test = True
        fnt = 'fnp5'
    elif (myFNRegex6.search(newfile) is None) is False:
        test = True
        fnt = 'fnp6'
    elif (myFNRegex7.search(newfile) is None) is False:
        test = True
        fnt = 'fnp7'
    elif (myFNRegex8.search(newfile) is None) is False:
        test = True
        fnt = 'fnp8'
    elif (myFNRegex9.search(newfile) is None) is False:
        test = True
        fnt = 'fnp9'
    elif (myFNRegex10.search(newfile) is None) is False:
        test = True
        fnt = 'fnp10'
    else:
        #print('File name: ' + newfile + ' \nDid not match mytype Regex.')
        test = False
        fnt = 'fnp0'

    return (test, fnt)

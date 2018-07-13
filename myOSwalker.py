#! python3
#File Name: myOSwalker.py
#Date of last edit: 3/21/2017
# This program module is looping through file directory

import os, getMyFileNameType

#This will loop through files in the current directory
print('\nhere are the current files: ')
for root, dirs, files in os.walk(".", topdown=False):
    for name in files:
        #print(os.path.join(root, name))
        
        
        # Open the file for processing
        file = open(os.path.join(root, name),'r', 1)
        #print('\ncurrent file handle: ' + file.name)
        cFile = name
        #print('current file name: ' + cFile)

        # get file name type
		# getMyFileNameType returns two items
		# test File is Bool and confirms that a filename format is defined
		# file name type is the matched format to extract the date format
        testF, fnt = getMyFileNameType.myFNType(cFile)
        #print('current file name type: ' + fnt)
        #print('current file test: ' + str(testF))

        # Assign the slicing integers
        if testF == True:
            if fnt != 'fnp0':
                if fnt == 'fnp1':
                    ss1 = 20 #Month 2char
                    es1 = 22
                    ss2 = 23 #Day 2char
                    es2 = 25
                    ss3 = 26 #Year 4char
                    es3 = 30
                    ss4 = 8 #hNumCk 2char
                    es4 = 10
                    fileReport = 'Bean_Data'                    

                elif fnt == 'fnp2':
                    ss1 = 19 #Month 2char
                    es1 = 21
                    ss2 = 22 #Day 2char
                    es2 = 24
                    ss3 = 25 #Year 4char
                    es3 = 29
                    ss4 = 7 #hNumCk 2char
                    es4 = 9
                    fileReport = 'Bean_Data' 

                elif fnt == 'fnp3':
                    ss1 = 20 #Month 2char
                    es1 = 22
                    ss2 = 23 #Day 2char
                    es2 = 25
                    ss3 = 26 #Year 4char
                    es3 = 30
                    ss4 = 3 #hNumCk 2char
                    es4 = 5
                    fileReport = 'Bean_Data' 

                elif fnt == 'fnp4':
                    ss1 = 23 #Month 2char
                    es1 = 25
                    ss2 = 26 #Day 2char
                    es2 = 28
                    ss3 = 29 #Year 4char
                    es3 = 33
                    ss4 = 3 #hNumCk 2char
                    es4 = 5
                    fileReport = 'Bean_Data' 

                elif fnt == 'fnp5':
                    ss1 = 15 #Month 2char
                    es1 = 17
                    ss2 = 18 #Day 2char
                    es2 = 20
                    ss3 = 21 #Year 4char
                    es3 = 25
                    ss4 = 3 #hNumCk 2char
                    es4 = 5
                    fileReport = 'Bean_Data' 

                elif fnt == 'fnp6':
                    ss1 = 4 #Month 2char
                    es1 = 6
                    ss2 = 6 #Day 2char
                    es2 = 8
                    ss3 = 0 #Year 4char
                    es3 = 4
                    ss4 = 0 #hNumCk 2char
                    es4 = 0
                    fileReport = 'TSM_EXPORT' 

                elif fnt == 'fnp7':
                    ss1 = 4 #Month 2char
                    es1 = 6
                    ss2 = 6 #Day 2char
                    es2 = 8
                    ss3 = 0 #Year 4char
                    es3 = 4
                    ss4 = 15 #hNumCk 2char
                    es4 = 17
                    fileReport = 'TSM_EXPORT' 

                elif fnt == 'fnp8':
                    ss1 = 4 #Month 2char
                    es1 = 6
                    ss2 = 6 #Day 2char
                    es2 = 8
                    ss3 = 0 #Year 4char
                    es3 = 4
                    ss4 = 15 #hNumCk 2char
                    es4 = 17
                    fileReport = 'TSM_EXPORT'

                elif fnt == 'fnp9':
                    ss1 = 4 #Month 2char
                    es1 = 6
                    ss2 = 6 #Day 2char
                    es2 = 8
                    ss3 = 0 #Year 4char
                    es3 = 4
                    ss4 = 0 #hNumCk 2char
                    es4 = 0
                    fileReport = 'TSM_EXPORT'

                elif fnt == 'fnp10':
                    ss1 = 4 #Month 2char
                    es1 = 6
                    ss2 = 6 #Day 2char
                    es2 = 8
                    ss3 = 0 #Year 4char
                    es3 = 4
                    ss4 = 13 #hNumCk 2char
                    es4 = 15
                    fileReport = 'TSM_EXPORT' 

                elif fnt == None:
                    print('File name type not assigned a value, is None')
                    fileReport = None
            else:
                #fnt == 'fnp0'
                print('File name type is no match with fnp0')
                fileReport = None
        else:
            #testF == False
            print('Test file was False regex match')
            fileReport = None

        # slice out dates
        if fileReport != None:
            # Build report date slicing
            # extract date from report source name
            myMM = cFile[ss1:es1]
            myDD = cFile[ss2:es2]
            myYYYY = cFile[ss3:es3]
            reportDate = myMM + '/' + myDD + '/' + myYYYY
            #print('reportDate = ' + myMM + '/' + myDD + '/' + myYYYY)
            hNumCk = cFile[ss4:es4]                
            
            if fileReport == 'Bean_Data':
                #TODO: pass in current file name, sheet name
                #OpenXls_BeanReport(cFile, hNumCk, reportDate, fileReport)
                #TODO: check if already read in
                #TODO: read in data
                #TODO: write PUBLIC
                #TODO: write CSV
                #TODO: write to Database
                print('fileReport = Bean_Data')

            elif fileReport == 'TSM_EXPORT':
                #TODO: pass in sheet name
                #TODO: check if already read in
                #TODO: read in data
                #TODO: write PUBLIC
                #TODO: write CSV
                #TODO: write to Database
                print('fileReport = TSM_EXPORT')

            else:
                # fileReport != 'TSM_EXPORT' or 'Bean_Data'
                #TODO: process this error
                print('fileReport != TSM_EXPORT or Bean_Data')

        else:
                # fileReport == None
                #TODO: process this error
                print('fileReport == None')

        # Close the open file 
        file.close

        print('\nnext file')
        
    print('\nnext folder')
    
print('done')


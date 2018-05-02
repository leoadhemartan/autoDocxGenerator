#! /usr/bin/python

#Generate Word Documents from Word Templates and Excel Source File
import os, sys, getopt, docx, openpyxl

def get_script_path(): #get current directory where script is located
    return os.path.dirname(os.path.realpath(sys.argv[0]))

os.chdir(get_script_path())

if len(sys.argv) == 3 :
    sourceFile=str(sys.argv[1])
    templateFile =str(sys.argv[2])
    print ('Source Filename     :'), sys.argv[1]
    print ('Template Filename   :'), sys.argv[2]

else:
    print ('No Source and Template files Specified.')
    sys.exit()

#Read source excel file
wb = openpyxl.load_workbook(sourceFile)
ws = wb['Sheet1']
myList = []
listIndex=0
for row in ws[ws.min_row:ws.max_row]:
    myList.append([])    
    for cell in row:
       myList[listIndex].append(cell.value)
    listIndex += 1
print(myList[0])
print('Source File successfully parsed')
wb.close

textToReplace = myList[0]
for x in myList[1:]:      
    wordDoc = docx.Document(templateFile)
    
    for p in wordDoc.paragraphs: 
        rIndex = 0   #this is the index that will handle all the columns in the current row   
        for tIndex in textToReplace:   
            if tIndex in p.text:
                replacementText =   str(x[rIndex]).strip()
                print('found ' + tIndex)                
                inline = p.runs
                # Loop added to work with runs (strings with same style)
                for i in range(len(inline)):
                    if tIndex in inline[i].text:
                        print(replacementText)    
                        text = inline[i].text.replace(tIndex, replacementText.upper())
                        inline[i].text = text    
            rIndex += 1
        
        filename= str(x[0]).strip() + '.docx'
    wordDoc.save(filename)        
   
    print(filename, ' Successfully Created!')
print('Script Finished')
## auto Docx Generator

This is just a simple script I developed to help my wife. She needed to create 100 contract documents for 100 employees. Contract wording was the same except for the Name, Position and Salary.

I immediately thought the task can be automated.

So from a spreadsheet listing down all the employee details, (**source.xlsx**) , a new Word Document is created from a template file (**source.xlsx**)

The script was tested on python 3.5 and requires the following libraries installed:
openpyxl
python-docx

These can be installed via pip

>pip install openpyxl

>pip install python-docx

**How to use:**
Create a Word document named template.docx and create unique identifiers for texts to be replaced. The script follows whatever formatting the identifier texts have. In the accompanying example files, the template.docx contains the following identifiers: 

xFirstNameMiddleInitialx

xLastNamex

xPositionx

xTitlex

xTestWordx

The first column is reserved for the FileName to use for each generated document.

These identifiers have to be in the first row of the source.xlsx and they serve as headers. Once you've populated the source.xlsx file, just run the script with the following syntax:

>python autoDocxGenerator.py SOURCE TEMPLATE

i.e.

>python autoDocxGenerator.py source.xlsx template.docx
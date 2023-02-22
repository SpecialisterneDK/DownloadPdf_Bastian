This program is made with a person who knows a bit about programmming and know how to navigate visual studio code and install extension

Make sure python is installed on your pc
Get python extension in visual studio code
Get Pip installer extension

To ensure you have the newest version of python go into the commando prompt and enter
You can find the commando prompt by entering cmd in the windows search

cd appdata\local\microsoft\windowsapps\pythonsoftwarefoundation.python (use tab on the keyboard to make it finish the last bit as it can be different from user to user)
when you are in there type pip install --upgrade pip

To run the program you have to install a few packages do this in visual studio code terminal

pip install pandas

pip install requests

pip install openpyxl

pip install xlsxwriter

When everything is installed run the program by pressing f5 after it is done it will create the file called downloaded.xlsx (this process will take about 1-2 hours) 
and say it is done in the terminal.
The file will be placed in the same folder as the program


The GRI_2017_2020_Test.xlsx file is for testing purpose to ensure that the program runs and won't take forever. 
I made it by taking the first 30 rows of the original xlsx file 
and one row further down where the AL link did not work but the AM link did.

if you wanna test the program yourself like this just change the code where it says GRI_2017_2020.xlsx to GRI_2017_2020_Test.xlsx







NOTE 1: some files there is being downloaded will be corrupt but that is not the fault of the program but a corrupt pdf file, the program won't know it is corrupt.

NOTE 2: if you have closed visual studio code and there is the textmining folder and the downloaded file the program can sometimes become confused when you run it again
To fix this just delete the downloaded file and the textmining folder, run the program again and it will work
#The goal of this project is to read data from an excel file, and write desired data to new file

import xlrd; #module for reading from excel files
import xlwt; #module for writing to excel files

workbook = xlrd.open_workbook('SGWS UPS AND DOWNS SEPT 2016.xlsx'); #opens the file and stores as variable to manipulate
worksheet = workbook.sheet_by_index(0);

#create a new, empty file
newWorkbook = xlwt.Workbook();
newSheet = newWorkbook.add_sheet('Sheet 1')
newWorkbook.save('testFile.ods');

#test extraction of data
for i in range(worksheet.nrows):
    if worksheet.cell(i, 10).value > worksheet.cell(i, 8).value:
        print(worksheet.cell(i, 1));

#The goal of this project is to read data from an excel file, and write desired data to new file

import xlrd; #module for reading from excel files
import xlwt; #module for writing to excel files


workbook = xlrd.open_workbook('target.xlsx'); #opens the file and stores as variable to manipulate
worksheet = workbook.sheet_by_index(0);

#create a new, empty file
newWorkbook = xlwt.Workbook();
newSheet = newWorkbook.add_sheet('Sheet 1');

newSheet.col(0).width = 15000;



#test extraction of data
counter = 0;
for i in range(worksheet.nrows):
    if worksheet.cell(i, 10).value > worksheet.cell(i, 8).value:
        #print(worksheet.cell(i, 1)); #test printing of relevant data
        newSheet.write(counter, 0, worksheet.cell(i, 1).value)
        counter += 1;

newWorkbook.save('onPost.ods')
print('Program complete. File saved as onPost.ods');

#It works. Program needs to alter formatting of spreadsheet

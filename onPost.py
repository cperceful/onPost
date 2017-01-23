#The goal of this project is to read data from an excel file, and write desired data to new file

import xlrd; #module for reading from excel files
import xlwt; #module for writing to excel files


def writeOn(oldSheet, newSheet):
    counter = 1; #start at 1, first row will contain column names
    for i in range(oldSheet.nrows):

        if (oldSheet.cell(i, 2).value == '1L'):
            continue; #skip 1L sizes
        elif (oldSheet.cell(i, 12).value > oldSheet.cell(i, 10).value): #used to be >=, doesn't matter for end of month, only regular can buy list
            #print(worksheet.cell(i, 1)); #test printing of relevant data
            newSheet.write(counter, 0, str(oldSheet.cell(i, 0).value) + " " + str(oldSheet.cell(i, 2).value)); #writes product name and size
            newSheet.write(counter, 1, str(oldSheet.cell(i, 10).value)); #writes January price
            newSheet.write(counter, 2, str(oldSheet.cell(i, 12).value)); #writes February price
            counter += 1;
    newSheet.write(counter + 1, 0, 'Vampire Pilot Studios, Enterprise Division');
    newSheet.write(counter + 2, 0, "Hi Rob, didn't you say you'd pay me for this?");

def writeMid(oldSheet, newSheet):
    counter = 1; #start at 1, first row will contain column names
    for i in range(oldSheet.nrows):
        if (oldSheet.cell(i, 2).value == '1L'):
            continue;
        elif (oldSheet.cell(i, 10).value < oldSheet.cell(i, 12).value and (oldSheet.cell(i, 8).value < oldSheet.cell(i, 10).value)):
            newSheet.write(counter, 0, str(oldSheet.cell(i, 0).value) + " " + str(oldSheet.cell(i, 2).value)); #writes product name
            newSheet.write(counter, 1, str(oldSheet.cell(i, 10).value)); #writes January price
            newSheet.write(counter, 2, str(oldSheet.cell(i, 12).value)); #writes February price
            counter += 1;

def writeFlat(oldSheet, newSheet):
    counter = 1;
    for i in range(oldSheet.nrows):
        if (oldSheet.cell(i, 2).value == '1L'):
            continue;
        elif (oldSheet.cell(i, 10).value == oldSheet.cell(i, 12).value):
            newSheet.write(counter, 0, str(oldSheet.cell(i, 0).value) + " " + str(oldSheet.cell(i, 2).value)); #writes product name
            newSheet.write(counter, 1, str(oldSheet.cell(i, 10).value));
            newSheet.write(counter, 2, str(oldSheet.cell(i, 12).value));
            counter += 1;


def setColumns(sheet):
    sheet.col(0).width = 15000;
    sheet.col(1).width = 4000;
    sheet.col(2).width = 4000;
    sheet.write(0, 0, "Product");
    sheet.write(0, 1, "January Price");
    sheet.write(0, 2, "February Price");




workbook = xlrd.open_workbook('target.xls'); #opens the file and stores as variable to manipulate
worksheet = workbook.sheet_by_index(0);

#create a new, empty file
newWorkbook = xlwt.Workbook();
onSheet = newWorkbook.add_sheet('On Post');
midSheet = newWorkbook.add_sheet('Mid Post');
flatSheet = newWorkbook.add_sheet('Flat post');

setColumns(onSheet);
setColumns(midSheet);
setColumns(flatSheet);

writeOn(worksheet, onSheet);
writeMid(worksheet, midSheet);
writeFlat(worksheet, flatSheet);


newWorkbook.save('Posts.ods');
print('Program complete. File saved as Posts.ods');

#It works. Program needs to alter formatting of spreadsheet

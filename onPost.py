#The goal of this project is to read data from an excel file, and write desired data to new file

import xlrd; #module for reading from excel files
import xlwt; #module for writing to excel files

#column 8 is Jan Price, 10 is Feb Price, 12 is Mar price

def writeOn(oldSheet, newSheet):
    counter = 1; #start at 1, first row will contain column names
    for i in range(oldSheet.nrows):

        if (oldSheet.cell(i, 2).value == '1L'):
            continue; #skip 1L sizes
        elif (oldSheet.cell(i, 8).value < oldSheet.cell(i, 10).value and oldSheet.cell(i, 8).value == oldSheet.cell(i, 12).value): #used to be >=, doesn't matter for end of month, only regular can buy list
            #print(worksheet.cell(i, 1)); #test printing of relevant data
            newSheet.write(counter, 0, str(oldSheet.cell(i, 0).value) + " " + str(oldSheet.cell(i, 2).value)); #writes product name and size
            newSheet.write(counter, 1, str(oldSheet.cell(i, 8).value)); #writes January price
            newSheet.write(counter, 2, str(oldSheet.cell(i, 10).value)); #writes February price
            counter += 1;
    newSheet.write(counter + 1, 0, 'Vampire Pilot Studios, Enterprise Division');
    newSheet.write(counter + 2, 0, "Hi Rob, didn't you say you'd pay me for this?");

def writeTwoMonths(oldSheet, newSheet):
    counter = 1;
    for i in range(oldSheet.nrows):
        if (oldSheet.cell(i, 2).value == '1L'):
            continue; #skip 1L sizes
        elif (oldSheet.cell(i, 8).value < oldSheet.cell(i, 10).value and oldSheet.cell(i, 10).value == oldSheet.cell(i, 12).value):
            newSheet.write(counter, 0, str(oldSheet.cell(i, 0).value) + " " + str(oldSheet.cell(i, 2).value)); #product name and size
            newSheet.write(counter, 1, str(oldSheet.cell(i, 8).value)); #write Jan price
            newSheet.write(counter, 2, str(oldSheet.cell(i, 10).value)); #write Feb price
            newSheet.write(counter, 3, str(oldSheet.cell(i, 12).value)); #write Mar price
            counter += 1;


def writeMid(oldSheet, newSheet):
    counter = 1; #start at 1, first row will contain column names
    for i in range(oldSheet.nrows):
        if (oldSheet.cell(i, 2).value == '1L'):
            continue;
        elif (oldSheet.cell(i, 8).value < oldSheet.cell(i, 10).value and (oldSheet.cell(i, 12).value < oldSheet.cell(i, 8).value)):
            newSheet.write(counter, 0, str(oldSheet.cell(i, 0).value) + " " + str(oldSheet.cell(i, 2).value)); #writes product name
            newSheet.write(counter, 1, str(oldSheet.cell(i, 8).value)); #writes January price
            newSheet.write(counter, 2, str(oldSheet.cell(i, 10).value)); #writes February price
            newSheet.write(counter, 3, str(oldSheet.cell(i, 12).value));
            counter += 1;

def writeFlat(oldSheet, newSheet):
    counter = 1;
    for i in range(oldSheet.nrows):
        if (oldSheet.cell(i, 2).value == '1L'):
            continue;
        elif (oldSheet.cell(i, 8).value == oldSheet.cell(i, 10).value and oldSheet.cell(i, 10).value == oldSheet.cell(i, 12).value):
            newSheet.write(counter, 0, str(oldSheet.cell(i, 0).value) + " " + str(oldSheet.cell(i, 2).value)); #writes product name
            newSheet.write(counter, 1, str(oldSheet.cell(i, 8).value)); #write Jan price
            newSheet.write(counter, 2, str(oldSheet.cell(i, 10).value)); #write Feb price
            newSheet.write(counter, 3, str(oldSheet.cell(i, 12).value)); #write March
            counter += 1;


# def setColumns(sheet):
#     sheet.col(0).width = 15000;
#     sheet.col(1).width = 4000;
#     sheet.col(2).width = 4000;
#     sheet.write(0, 0, "Product");
#     sheet.write(0, 1, "January Price");
#     sheet.write(0, 2, "February Price");

def oneMonthColumns(sheet):
    sheet.col(0).width = 15000;
    sheet.col(1).width = 4000;
    sheet.col(2).width = 4000;
    sheet.write(0, 0, 'One Month Posts');
    sheet.write(0, 1, 'January Price');
    sheet.write(0, 2, 'February Price');

def twoMonthColumns(sheet):
    sheet.col(0).width = 15000;
    sheet.col(1).width = 4000;
    sheet.col(2).width = 4000;
    sheet.col(3).width = 4000;
    sheet.write(0, 0, 'Two Month Posts');
    sheet.write(0, 1, 'January Price');
    sheet.write(0, 2, 'February Price');
    sheet.write(0, 3, 'March Price');


workbook = xlrd.open_workbook('target.xls'); #opens the file and stores as variable to manipulate
worksheet = workbook.sheet_by_index(0);

#create a new, empty file
newWorkbook = xlwt.Workbook();
oneMonthPost = newWorkbook.add_sheet('On Post One Month');
twoMonthPost = newWorkbook.add_sheet('On Post Two Months');
midSheet = newWorkbook.add_sheet('Mid Post');
flatSheet = newWorkbook.add_sheet('Flat post');

oneMonthColumns(oneMonthPost);
twoMonthColumns(twoMonthPost);
twoMonthColumns(midSheet);
twoMonthColumns(flatSheet);

writeOn(worksheet, oneMonthPost);
writeTwoMonths(worksheet, twoMonthPost);
writeMid(worksheet, midSheet);
writeFlat(worksheet, flatSheet);


newWorkbook.save('Posts.ods');
print('Program complete. File saved as Posts.ods');

#It works. Program needs to alter formatting of spreadsheet

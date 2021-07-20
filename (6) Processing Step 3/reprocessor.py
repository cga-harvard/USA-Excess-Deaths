# Author: Vansh Tibrewal (vanshtibrewal2004@gmail.com)

#processes monthly data back to the original data format

import openpyxl
import pandas
import datetime

#loading data and setting up writer
book = openpyxl.load_workbook('/Users/Vansh/PycharmProjects/CODprocessing/processedSplit3.xlsx') #input
writer = pandas.ExcelWriter('/Users/Vansh/PycharmProjects/CODprocessing/reprocessedSplit3.xlsx', engine='openpyxl') #output
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
refbook = openpyxl.load_workbook('/Users/Vansh/PycharmProjects/CODprocessing/refSplit3.xlsx') #reference sheet(a master sheet that connects the split code to the state and cause of death names and codes. In this case the reference sheet contains the original data as well, however the actual number of cases in the original data is not required, only the ability to link codes to names

#creating reference dataframe
for sheet in refbook.worksheets:
    if not sheet.title == "Master":
        continue
    data = sheet.values
    cols = next(data)[0:]
    data = list(data)
    refdf = pandas.DataFrame(data, columns=cols)
    del refdf["Year"]
    del refdf["Year Code"]
    del refdf["Month Code"]
    del refdf["Month"]
    del refdf["Deaths"]

#function to look up the state, state code, cause of death and join code from the reference dataframe
def findMyReference(splitval):
    for index,row in refdf.iterrows():
        if row["Split Code"] == splitval:
            return row["State"], row["State Code"], row["Cause of Death"], row["Join Code.New"]

#looping and processing the sheets
for sheet in book.worksheets:
    sheetname = sheet.title
    data = sheet.values
    cols = next(data)[0:]
    data = list(data)
    df = pandas.DataFrame(data, columns=cols)
    State = []
    State_Code = []
    Year = []
    Year_Code = []
    Month = []
    Month_Code = []
    Deaths = []
    Cause_Of_Death = []
    JoinCode_New = []
    SplitCode = []

    curstate, curstatecode, curcauseofd, curjoincodenew = findMyReference(sheetname)

    for index,row in df.iterrows():
        yr = int(row["Year"])
        for i in range(12):
            State.append(curstate)
            State_Code.append(curstatecode)
            Cause_Of_Death.append(curcauseofd)
            JoinCode_New.append(curjoincodenew)
            Year.append(yr)
            Year_Code.append(yr)
            SplitCode.append(sheetname)
        #jan
        Deaths.append(row["January"])
        Month.append(datetime.date(yr,1,1))
        Month_Code.append(datetime.date(yr, 1, 1))
        #feb
        Deaths.append(row["February"])
        Month.append(datetime.date(yr,2,1))
        Month_Code.append(datetime.date(yr, 2, 1))
        #march
        Deaths.append(row["March"])
        Month.append(datetime.date(yr,3,1))
        Month_Code.append(datetime.date(yr, 3, 1))
        #april
        Deaths.append(row["April"])
        Month.append(datetime.date(yr,4,1))
        Month_Code.append(datetime.date(yr, 4, 1))
        #may
        Deaths.append(row["May"])
        Month.append(datetime.date(yr,5,1))
        Month_Code.append(datetime.date(yr, 5, 1))
        #june
        Deaths.append(row["June"])
        Month.append(datetime.date(yr,6,1))
        Month_Code.append(datetime.date(yr, 6, 1))
        #july
        Deaths.append(row["July"])
        Month.append(datetime.date(yr,7,1))
        Month_Code.append(datetime.date(yr, 7, 1))
        #aug
        Deaths.append(row["August"])
        Month.append(datetime.date(yr,8,1))
        Month_Code.append(datetime.date(yr, 8, 1))
        #sept
        Deaths.append(row["September"])
        Month.append(datetime.date(yr,9,1))
        Month_Code.append(datetime.date(yr, 9, 1))
        #oct
        Deaths.append(row["October"])
        Month.append(datetime.date(yr,10,1))
        Month_Code.append(datetime.date(yr, 10, 1))
        #nov
        Deaths.append(row["November"])
        Month.append(datetime.date(yr,11,1))
        Month_Code.append(datetime.date(yr, 11, 1))
        #dec
        Deaths.append(row["December"])
        Month.append(datetime.date(yr,12,1))
        Month_Code.append(datetime.date(yr, 12, 1))

    data = {
        'State':State,
        'State Code':State_Code,
        'Year':Year,
        'Year Code': Year_Code,
        'Month': Month,
        'Month Code': Month_Code,
        'Deaths': Deaths,
        'Cause of Death': Cause_Of_Death,
        'Join Code.New': JoinCode_New,
        'Split Code': SplitCode,
    }
    #writing the processed data onto sheets
    writer.book.remove(writer.book[sheetname])
    sheetname = sheetname + " " #workaround to save time by only having to do writer.save() once
    dfprocessed = pandas.DataFrame(data, columns=['State','State Code','Year','Year Code','Month','Month Code','Deaths','Cause of Death','Join Code.New','Split Code'])
    dfprocessed.to_excel(writer, sheetname, index=False)

for sheet in writer.book:
    sheet.title = sheet.title.strip() #workaround(as mentioned above)(without changing the names of sheets)

writer.save()
print("done")
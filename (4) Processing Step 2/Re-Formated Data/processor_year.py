# Author: Vansh Tibrewal (vanshtibrewal2004@gmail.com)

#processes yearly data sheets

import openpyxl
from itertools import islice
import pandas

#loading data and setting up writer
book = openpyxl.load_workbook('/Users/Vansh/PycharmProjects/CODprocessing/Split3Year.xlsx') #input
writer = pandas.ExcelWriter('/Users/Vansh/PycharmProjects/CODprocessing/processedSplit3Year.xlsx', engine='openpyxl') #output
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

#looping and processing the sheets
for sheet in book.worksheets:
    if sheet.title == "Master":
        continue
    if sheet.title == "SplitCode":
        continue
    sheetname = sheet.title
    data = sheet.values
    cols = next(data)[1:]
    data = list(data)
    idx = [r[0] for r in data]
    data = (islice(r, 1, None) for r in data)
    df = pandas.DataFrame(data, index=idx, columns=cols)
    del df["State Code"]
    del df["Year Code"]
    del df["Join Code.New"]
    del df["SplitCode"]
    year = []
    deaths = []

    for index,row in df.iterrows():
        year.append(row["Year"])
        if(row["Deaths"]=="Suppressed"):
            deaths.append("Suppressed")
        else:
            deaths.append(int(row["Deaths"]))

    data = {
        'Year':year,
        'Deaths':deaths
    }
    #writing the processed data onto sheets
    writer.book.remove(writer.book[sheetname])
    sheetname = sheetname + " " #workaround to save time by only having to do writer.save() once
    dfprocessed = pandas.DataFrame(data, columns=['Year','Deaths'])
    dfprocessed.to_excel(writer, sheetname, index=False)

for sheet in writer.book:
    sheet.title = sheet.title.strip() #workaround(as mentioned above)(without changing the names of sheets)

writer.save()
print("done")

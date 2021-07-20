# Author: Vansh Tibrewal (vanshtibrewal2004@gmail.com)

#processes monthly data sheets

import openpyxl
from itertools import islice
import pandas

#loading data and setting up writer
book = openpyxl.load_workbook('/Users/Vansh/PycharmProjects/CODprocessing/Split3Month.xlsx') #input
writer = pandas.ExcelWriter('/Users/Vansh/PycharmProjects/CODprocessing/processedSplit3Month.xlsx', engine='openpyxl') #output
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
    del df["Month Code"]
    del df["Join Code.New"]
    del df["Split Code"]
    year = []
    m1 = []
    m2 = []
    m3 = []
    m4 = []
    m5 = []
    m6 = []
    m7 = []
    m8 = []
    m9 = []
    m10 = []
    m11 = []
    m12 = []
    for index,row in df.iterrows():
        tempmonth = row["Month"].month
        if (int(tempmonth) % 12 == 1):
            year.append(row["Year"])
            if(row["Deaths"]=="Suppressed"):
                m1.append("Suppressed")
            else:
                m1.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 2):
            if(row["Deaths"]=="Suppressed"):
                m2.append("Suppressed")
            else:
                m2.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 3):
            if(row["Deaths"]=="Suppressed"):
                m3.append("Suppressed")
            else:
                m3.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 4):
            if(row["Deaths"]=="Suppressed"):
                m4.append("Suppressed")
            else:
                m4.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 5):
            if(row["Deaths"]=="Suppressed"):
                m5.append("Suppressed")
            else:
                m5.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 6):
            if(row["Deaths"]=="Suppressed"):
                m6.append("Suppressed")
            else:
                m6.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 7):
            if(row["Deaths"]=="Suppressed"):
                m7.append("Suppressed")
            else:
                m7.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 8):
            if (row["Deaths"] == "Suppressed"):
                m8.append("Suppressed")
            else:
                m8.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 9):
            if (row["Deaths"] == "Suppressed"):
                m9.append("Suppressed")
            else:
                m9.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 10):
            if (row["Deaths"] == "Suppressed"):
                m10.append("Suppressed")
            else:
                m10.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 11):
            if (row["Deaths"] == "Suppressed"):
                m11.append("Suppressed")
            else:
                m11.append(int(row["Deaths"]))
        if (int(tempmonth) % 12 == 0):
            if (row["Deaths"] == "Suppressed"):
                m12.append("Suppressed")
            else:
                m12.append(int(row["Deaths"]))
    data = {
        'Year':year,
        'January':m1,
        'February':m2,
        'March': m3,
        'April': m4,
        'May': m5,
        'June': m6,
        'July': m7,
        'August': m8,
        'September': m9,
        'October': m10,
        'November': m11,
        'December': m12
    }
    #writing the processed data onto sheets
    writer.book.remove(writer.book[sheetname])
    sheetname = sheetname + " " #workaround to save time by only having to do writer.save() once
    dfprocessed = pandas.DataFrame(data, columns=['Year','January','February','March','April','May','June','July','August','September','October','November','December'])
    dfprocessed.to_excel(writer, sheetname, index=False)

for sheet in writer.book:
    sheet.title = sheet.title.strip() #workaround(as mentioned above)(without changing the names of sheets)

writer.save()
print("done")

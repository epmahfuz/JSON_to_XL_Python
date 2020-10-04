import os
import json
import xlsxwriter 
import xlrd

#Localization file path
path = "D:\\Translation - Addexpert - Python\\Auto-Translation\\main\\i18n"
files = os.listdir(path)

#Excel File name
workbook = xlsxwriter.Workbook('TranslationWithoutDuplicate5.xlsx') 
worksheet = workbook.add_worksheet("My sheet")

data = ()
loadedData = []
tempData = []

def leafValue(val):
    for a, b in val.items():
        if isinstance(b, str):
            temp = b
            temp = temp.replace(" ", "")
            temp = temp.replace(".", "")
            temp = temp.replace("_", "")
            temp = temp.upper()
            if(temp not in tempData):
                tempData.append(temp)
                loadedData.append(b)
        elif len(b)!=0:
            leafValue(b)

def openFile(filePath):
	with open(filePath, 'r', encoding="utf8") as file:
		otherfile = json.load(file)
	return otherfile

for file in files:
    f = path + "\\"+ file +"\\en.json"
    
    print(os.stat(f).st_size)
    if(os.stat(f).st_size != 0):
        val = openFile(f)
        leafValue(val)

for i in loadedData:
    data = data +([i, "", ""],)


# Start from the first cell. Rows and columns are zero indexed. 
row = 0
col = 0
  
# Iterate over the data and write it out row by row. 
for en, de, fr in (data):
    worksheet.write(row, col, en)
    worksheet.write(row, col + 1, de)
    worksheet.write(row, col + 2, fr)
    row += 1
  
workbook.close()
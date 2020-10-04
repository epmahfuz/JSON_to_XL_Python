import os
import json
import xlsxwriter 
import xlrd

#Localization file path
path = "D:\\Translation - Addexpert - Python\\Auto-Translation\\i18n"

#Excel File name
workbook = xlsxwriter.Workbook('AllTranslationData.xlsx') 
worksheet = workbook.add_worksheet("My sheet") 
#wb = xlrd.open_workbook(loc) 
#sheet = wb.sheet_by_index(0)

#Load File Directory
files = os.listdir(path)

data = ()
loadedLocalizedKeyData = {}

def leafValue(val):
    for a, b in val.items():
        if isinstance(b, str):
            print(a)
            loadedLocalizedKeyData[a] = b
        elif len(b)!=0:
            leafValue(b)

def openFile(filePath):
	#print(filePath)
	with open(filePath, 'r', encoding="utf8") as file:
		otherfile = json.load(file)
	return otherfile

for file in files:
	#print(file)
    f = "D:\\Translation - Addexpert - Python\\Auto-Translation\\i18n\\"+ file +"\\en.json"
    
    print(os.stat(f).st_size)
    if(os.stat(f).st_size != 0):
        val = openFile(f)
        leafValue(val)

for i,j in loadedLocalizedKeyData.items():
    data = data +([i,loadedLocalizedKeyData[i],"",""],)


# Start from the first cell. Rows and 
# columns are zero indexed. 
row = 0
col = 0
  
# Iterate over the data and write it out row by row. 
for key, en,de,fr in (data): 
    worksheet.write(row, col, key) 
    worksheet.write(row, col + 1, en)
    worksheet.write(row, col + 2, de)
    worksheet.write(row, col + 3, fr) 
    row += 1
  
workbook.close() 

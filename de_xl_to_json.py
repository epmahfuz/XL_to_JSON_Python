#give: en.json and sample.xlsx with key, en and de translation
#get: de.json
import os
import json
import xlsxwriter 
import xlrd
import excel2json

#All JSON Folders Loaded Here
json_path = "D:\\Translation - Addexpert - Python\\Auto-Translation\\test\\i18n"
jsonFolders = os.listdir(json_path)

#xl File Loaded Here
xlPath = "D:\\Translation - Addexpert - Python\\AllTranslationDataTest.xlsx"
wb = xlrd.open_workbook(xlPath) 
xlSheet = wb.sheet_by_index(0)

#Functions Section Start
def search_in_xl_sheet(key, eng, lan):
    for i in range(xlSheet.nrows):
        if (xlSheet.cell_value(i, 0) == key):
            if(lan == "de"):
                return xlSheet.cell_value(i, 2)
            else:
                return xlSheet.cell_value(i, 3)

def update_dict(dict_to_update, path, value):
    obj = dict_to_update
    key_list = path.split(".")
    for k in key_list:
        if(k==key_list[-1]):
            obj[k] = value
        else:
            obj = obj[k]

def leafValue(value, loc):
    for key, eng in value.items():
        if(loc == ""):
            tempLoc = key
        else:
            tempLoc = loc + '.' + key
        
        if isinstance(eng, str):
            de_value = search_in_xl_sheet(key, eng, "de")
            update_dict(temp_de_json, tempLoc, de_value)
        elif len(eng) != 0:
            leafValue(eng, tempLoc)
    
def openFile(filePath):
	with open(filePath, 'r', encoding="utf8") as file:
		otherfile = json.load(file)
	return otherfile
#Functions Section End

for folder in jsonFolders:
	en_path = json_path + "\\" + folder + "\\en.json"
	de_path = json_path + "\\" + folder + "\\de.json"

	if (os.stat(en_path).st_size != 0 ):
		global temp_de_json
		en_json = openFile(en_path)

		temp_de_json = en_json

		leafValue(en_json, "")
        
		with open(de_path, "w") as de_json:
			json.dump(temp_de_json, de_json)

#give: en.json and sample.xlsx with key, en and de translation
#get: de.json
import os
import json
import xlsxwriter 
import xlrd
import excel2json

#All JSON Folders Loaded Here
json_path = "D:\\Translation - Addexpert - Python\\Auto-Translation\\main\\i18n"
jsonFolders = os.listdir(json_path)

#xl File Loaded Here
xlPath = "D:\\Translation - Addexpert - Python\\clientProvided_2.xlsx"
wb = xlrd.open_workbook(xlPath) 
xlSheet = wb.sheet_by_index(0)

#Functions Section Start
#here making string solid
def make_solid_string(string):
    string = string.replace(" ", "")
    string = string.replace(".", "")
    string = string.replace("_", "")
    string = string.upper()
    return string

def search_in_xl_sheet(eng, lan):
    temp_eng = make_solid_string(eng)
    for i in range(xlSheet.nrows):
        temp_val = make_solid_string(xlSheet.cell_value(i, 0))
        if (temp_eng == temp_val):
            if(lan == "de"):
                print(eng)
                return xlSheet.cell_value(i, 2)
            else:
                return xlSheet.cell_value(i, 3)
    return ""

def update_dict(dict_to_update, path, value):
    #print(path)
    obj = dict_to_update
    key_list = path.split(".")
    for k in key_list:
        if(k==key_list[-1]):
            obj[k] = value
        else:
            obj = obj[k]

#value: an object where i'm in. loc: it keeps the nesting address of the object 
def leafValue(value, loc):
    #print(value)
    for key, eng in value.items():
        if(loc == ""):
            tempLoc = key
        else:
            tempLoc = loc + '.' + key
        
        if isinstance(eng, str): # if it is string, its the leaf key, it will be searched
            de_value = search_in_xl_sheet(eng, "de")
            if(de_value !=""):
                update_dict(temp_de_json, tempLoc, de_value)
        elif len(eng) != 0: #its a object & lenght getter than 0
            leafValue(eng, tempLoc)
    
def openFile(filePath):
	with open(filePath, 'r', encoding="utf8") as file:
		otherfile = json.load(file)
	return otherfile
#Functions Section End

for folder in jsonFolders:
	en_path = json_path + "\\" + folder + "\\en.json"
	de_path = json_path + "\\" + folder + "\\de.json"
	#print("Forlder: "+folder)
	if (os.stat(en_path).st_size != 0 ):
		global temp_de_json
		en_json = openFile(en_path)
		de_json = openFile(de_path)

		temp_de_json = de_json

		leafValue(en_json, "")
        
		with open(de_path, "w") as de_json_single:
			json.dump(temp_de_json, de_json_single)

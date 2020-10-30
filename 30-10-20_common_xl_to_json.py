#give: en.json and sample.xlsx with key, en and de translation
#get: de.json
import os
import json
import xlsxwriter 
import xlrd
import excel2json

#All JSON Folders Loaded Here
json_path = "D:\\Translation - Addexpert - Python\\Auto-Translation\\main\\i18n" #Edit Here
jsonFolders = os.listdir(json_path)

#xl File Loaded Here
xlPath = "D:\\Translation - Addexpert - Python\\clientProvided.xlsx" #Edit Here
wb = xlrd.open_workbook(xlPath) 
xlSheet = wb.sheet_by_index(0)

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
                return xlSheet.cell_value(i, 2)
            else:
                return xlSheet.cell_value(i, 3)
    print("Not Found: "+eng)
    return eng


def update_dict(dict_to_update, path, value):
    obj = dict_to_update
    key_list = path.split(".")
    for k in key_list:
        if(k==key_list[-1]):
            obj[k] = value
        else:
            obj = obj[k]

#value: an object where i'm in. loc: it keeps the nesting address of the object 
def leafValue(value, loc, lan):
    for key, eng in value.items():
        if(loc == ""):
            tempLoc = key
        else:
            tempLoc = loc + '.' + key
        
        if isinstance(eng, str): # if it is string, its the leaf key, it will be searched
            trans_value = search_in_xl_sheet(eng, lan)
            update_dict(trans_json, tempLoc, trans_value)
        elif len(eng) != 0: #its a object & length getter than 0
            leafValue(eng, tempLoc, lan)
    
def openFile(filePath):
	with open(filePath, 'r', encoding="utf8") as file:
		otherfile = json.load(file)
	return otherfile

#Start From Here
for folder in jsonFolders:
	en_path = json_path + "\\" + folder + "\\en.json"
	trans_path = json_path + "\\" + folder + "\\fr.json" #Edit Here
	if (os.stat(en_path).st_size != 0 ):
		global trans_json
		language = "fr" #Edit Here
		en_json = openFile(en_path)
		trans_json = en_json
		leafValue(en_json, "", language)
        
		with open(trans_path, "w") as single_json:
			json.dump(trans_json, single_json)

# I have a xl file that contains three rows eng, fr, de. I have also a translation folder including many folders. In every folder, it contains a en.json. From this en.json i will pick a line that has a KEY and Value. I will compare with every lines in eng row. If i found same eng line, i will pick the de/fr value from the sheet and will put in my targated json file. Firstly, i make a copy of en.json and gradually updating the json.
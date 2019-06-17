#import needed libraries
from stix2 import TAXIICollectionSource, Filter
from taxii2client import Server, Collection
import json
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment

#enterprise attack source
collection = Collection("https://cti-taxii.mitre.org/stix/collections/95ecc380-afe9-11e4-9b6c-751b66dd541e/")
# supply the TAXII2 collection to TAXIICollection
tc_source = TAXIICollectionSource(collection)

#filter to only techniques
filt = Filter('type', '=', 'attack-pattern')
techniques = tc_source.query([filt])


#generate list of all kill chain phases
kc_list = []

for technique in techniques:
	for k,v in technique.items():
		for i in technique["kill_chain_phases"]:
			for k,v in i.items():
				kc_list.append(v.encode("utf-8"))

#remove duplicates and sort
deduped_kc = sorted(list(dict.fromkeys(kc_list)))
deduped_kc.remove('mitre-attack')
print deduped_kc

#seperate platforms
windows = []
linux = []
mac = []
all_techniques = []

for technique in techniques:
    for k,v in technique.items():
    	if k == 'x_mitre_platforms':
    		all_techniques.append(technique)
    		if 'Windows' in v:
    			windows.append(technique)			
    		if "Linux" in v:
    			linux.append(technique)
    		if "macOS" in v:
    			mac.append(technique)


#create worksheet tabs
count = 1

wb = Workbook()
sheet1 = wb.get_sheet_by_name('Sheet')
sheet1.title = 'All_Techniques'
sheet2 = wb.create_sheet('Windows')
sheet3 = wb.create_sheet('Linux')
sheet4 = wb.create_sheet('macOS')

list_of_lists = []

def parse_json(platform):
	global count
	#create empty list for each kill chain phase:

	list_of_lists = []

	for i in deduped_kc:
		list_name = []
		list_of_lists.append(list_name)


	#for each kill chain phases

	for lst in platform:
			for technique in lst:
				for index in enumerate(deduped_kc):
					for i in technique["kill_chain_phases"]:
						for k,v in i.items():
							if index[1] == v:
								list_of_lists[index[0]].append(technique["name"].encode("utf-8"))

	# write to excel file all the data, making first row bold
	bold_font = Font(bold = True)		#set bold font variable
	wrap_text = Alignment(wrap_text=True)

	if count == 1:
		current_sheet = sheet1
	elif count == 2:
		current_sheet = sheet2
	elif count == 3:
		current_sheet = sheet3
	elif count == 4:
		current_sheet = sheet4

	current_sheet.append(deduped_kc)
	for cell in current_sheet["1:1"]:
		cell.font = bold_font
		cell.alignment = wrap_text
	
	#This part is how the Excel columns are built, incrementing character and count
	column = '@'

	for lst in list_of_lists:
		column = chr(ord(column) + 1)
		row = 2
		for technique in lst:
			current_sheet[column + str(row)] = technique
			current_sheet[column + str(row)].alignment = wrap_text
			row += 1 

	
	count += 1


	wb.save('file_name.xlsx')							


parse_json([all_techniques])
parse_json([windows]) 
parse_json([linux])
parse_json([mac])



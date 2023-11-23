import openpyxl
import xml.etree.ElementTree

# Открытие файла 
file = openpyxl.load_workbook('listOfDispa.xlsx')

sheet = file['Лист2']

length = sheet.max_column-1
listOfPatient = [[] for i in range(length)]

for i in range(1, length+1):
	for j in sheet:
		listOfPatient[i-1] += [j[i].value]

# print(listOfPatient[0][0])

# for i in range(0, len(listOfPatient)):
# 	for j in range(0, len(listOfPatient[i])):
# 		print(listOfPatient[i][j], end=' ')
# 		break

# print(listOfPatient[0][333]) # [1-7][0-333]


for i in range(len(listOfPatient[0])): # range(len(listOfPatient))

	print("<FAM>" + listOfPatient[0][i] + "</FAM>")
	print("<IM>" + listOfPatient[1][i] + "</IM>")
	print("<OT>" + listOfPatient[2][i] + "/<OT>")
	print("<W>" + listOfPatient[3][i] + "</W>")
	print("<DR>" + listOfPatient[4][i] + "</DR>")
	print("<PHONE>" + listOfPatient[5][i] + "</PHONE>")
	print("<ENP>" listOfPatient[6][i] + "</ENP>")
	print("<SNILS>" + listOfPatient[7][i] + "</SNILS>")

	# for j in range(len(listOfPatient[0])):
	# 	print(listOfPatient[i][j])

# f = open('some.xml', 'a')
# f.close()
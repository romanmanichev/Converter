from openpyxl import load_workbook
from lxml import etree


# Загрузка excel файла и листа
wb = load_workbook('listOfDispa.xlsx')
sheet = wb['Лист2']

# Создание <ZL_LIST> тега
ZL_LIST = etree.Element("ZL_LIST")


# Теги для формарования xml файла
tags = ['N_ZAP', 'FAM', 'IM', 'OT', 'W', 'DR', 'PHONE', 'ENP', 'SNILS'] # 9 

# Перебор ecxel файла
for i in range(1, sheet.max_row+1):
	# Создание <ZAP> тега
	column = etree.SubElement(ZL_LIST, "ZAP")
	# Получение итерируемой строки
	for j in range(1, sheet.max_column+1):
		
		# Формирование тега из переменной tags
		column1 = etree.SubElement(column, tags[j-1])
		# Добавление соответствующего текста в тег
		column1.text = str(sheet.cell(row=i, column=j).value)

# Создание дерева и xml файла 
tree = etree.ElementTree(ZL_LIST)
tree.write("example.xml", pretty_print=True, encoding='UTF-8')

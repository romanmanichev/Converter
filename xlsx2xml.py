from openpyxl import load_workbook
from lxml import etree
import datetime as dt


def xlsx2xml(fileEXCEL, fileXML, tags, NameOfList='Лист1'):
	# Загрузка excel файла и листа
	wb = load_workbook(f"{fileEXCEL}.xlsx")
	sheet = wb[f"{NameOfList}"]

	# Создание <ZL_LIST> тега
	ZL_LIST = etree.Element("ZL_LIST")

	# Формирование тега <ZGLV>
	ZGLV = etree.SubElement(ZL_LIST, "ZGLV")
	filename = etree.SubElement(ZGLV, "FILENAME")
	data = etree.SubElement(ZGLV, "DATA")
	code_mo = etree.SubElement(ZGLV, "CODE_MO")
	year = etree.SubElement(ZGLV, "YEAR")
	r = etree.SubElement(ZGLV, "R")

	# заполнение тега <ZGLV>
	filename.text = fileXML
	data.text = '2023-01-01'
	code_mo.text = '352506'
	year.text = str(dt.date.today().year)
	r.text = '1'

	# Счетчик для идентифиактора N_ZAP
	counter = 1

	# Функция для проверки даты по фильтру
	def check_date(olddate, filter):
		try:
			newdate = dt.datetime.strptime(olddate, filter)
			return newdate.date()
		except:
			return False

	# Перебор excel файла
	for i in range(1, sheet.max_row+1):
		# Создание <ZAP> тега
		column = etree.SubElement(ZL_LIST, "ZAP")
		# Получение итерируемой строки
		for j in range(1, sheet.max_column+1):
			# Проврека для N_ZAP и DR
			if tags[j-1] == "N_ZAP":

				# Формирование тега из переменной tags
				column1 = etree.SubElement(column, tags[j-1])

				# Добавление соответствующего текста в тег
				column1.text = str(counter)
				counter += 1
			
			elif tags[j-1] == "DR":

				# Преобразование даты в следующую маску ГГГГ-ММ-ДД
				olddate = str(sheet.cell(row=i, column=j).value)

				# Фильтр даты (костыль)
				column1 = etree.SubElement(column, tags[j-1])

				if check_date(olddate, '%Y-%m-%d %H:%M:%S') != False:
					column1.text = str(check_date(olddate, '%Y-%m-%d %H:%M:%S'))
				elif check_date(olddate.strip(' '), '%d.%m.%Y') != False:
					column1.text = str(check_date(olddate.strip(' '), '%d.%m.%Y'))
				else:
					# Даты, которые не попали под фильтрацию
					column1.text = olddate
			elif tags[j-1] == "PHONE":
				
				# Формирование тега из переменной tags
				column1 = etree.SubElement(column, tags[j-1])

				# Добавление соответствующего текста в тег
				column1.text = str('79111111111')

			else:

				# Формирование тега из переменной tags
				column1 = etree.SubElement(column, tags[j-1])

				# Добавление соответствующего текста в тег
				if sheet.cell(row=i, column=j).value == None:
					column1.text = ""
				else:
					column1.text = str(sheet.cell(row=i, column=j).value)

		# Ввод в xml файл доп.тегов
		column2 = etree.SubElement(column, "MDR")
		column2.text = str(2)
		column3 = etree.SubElement(column, "ADDRESDP")
		column3.text = str("г.Вологда, ул.Окружное ш., д.3в")
		column4 = etree.SubElement(column, "DISP_TYP")
		column4.text = str(4)
		column5 = etree.SubElement(column, "DATADP")
		column5.text = str("2023-11-20")
		column6 = etree.SubElement(column, "TIMEDP")
		column6.text = str("09:00:00")


	# Создание дерева и xml файла
	tree = etree.ElementTree(ZL_LIST)
	tree.write(f"{fileXML}.xml", pretty_print=True, encoding='windows-1251')

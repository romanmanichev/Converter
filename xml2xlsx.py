import xml.etree.ElementTree as ET
from openpyxl import Workbook


def xml2xlsx(fileXML, fileEXCEL):
    # Открытие xml файла
    tree = ET.parse(f"{fileXML}.xml")
    root = tree.getroot()
    

    # Данные xml для внесения в excel файл
    datas = [() for i in range(len(root))]

    # Парсинг каждого элемента xml файла
    for elem in range(1, len(root)):
        for subelem in root[elem]:
            datas[elem-1] += (subelem.text, )


    # Создание листа excel
    wb = Workbook()
    lst = wb.active

    # Создание строки с заголовками
    # lst.append(("name", "surname"))

    for data in datas:
        lst.append(data)


    # Создание excel файла
    wb.save(f"{fileEXCEL}.xlsx")

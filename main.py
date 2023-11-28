from xml2xlsx import xml2xlsx
from xlsx2xml import xlsx2xml


print("конвертировать из xml в xlsx файл[1]\nКонвертировать из xlsx в xml[2]")
if int(input('Введите цифру: ')) == 1:

    # Конвертация файла из xml в xlsx
    # Названия файлов для конвертации (без рассширении)
    # Входной файл
    fileXML = 'dp-0101'

    # Выходной файл
    fileEXCEL = 'xl'

    xml2xlsx(fileXML, fileEXCEL)
else:

    # Конвертация файла из xlsx в xml
    # Названия файлов для конвертации (без рассширении)
    # Входной файл
    fileEXCELforImport = 'listOfDispa'
    NameOfList = 'Лист2'

    # Выходной файл
    fileXMLforExport = 'dp-0101'

    # Теги для формарования xml файла
    tags = ['N_ZAP', 'FAM', 'IM', 'OT', 'W', 'DR', 'PHONE', 'ENP', 'SNILS'] # 9

    xlsx2xml(fileEXCELforImport, fileXMLforExport, tags, NameOfList)

from tkinter import filedialog
import customtkinter
from xlsx2xml import xlsx2xml, get_list_from_excel_file
from xml2xlsx import xml2xlsx
import os.path


customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("dark-blue")

app = customtkinter.CTk()
app.geometry("350x520")
app.title("Конвертер xml и excel файлов")
app.resizable(width=False, height=False)

tabView = customtkinter.CTkTabview(app, width=600, height=520)
tabView.pack(padx=0, pady=0)
tabView.add("из excel в xml")
tabView.add("из xml в excel")


# Функция выбора excel файла
def chooseExcelFileForImport():
    # Переменная для получения абсолютного пути excel файла
    directoryChosenFile = filedialog.askopenfilename()
    # Проверка если файл был выбран
    if directoryChosenFile != '':
        global directoryExcelFileForConvert

        # Получение названии листов excel файла
        try:
            listOfExcel.configure(values=get_list_from_excel_file(directoryChosenFile))
        except ValueError:
            label6.configure(text='выбран формат не excel файла', text_color='#FF0000')
            return None
        except:
            label6.configure(text='ошибка при выборе файла', text_color='#FF0000')
            return None

        # Получение абсолютного пути файла для конвертации
        directoryExcelFileForConvert = directoryChosenFile[:]

        # Получение названия excel файла
        directoryChosenFile = str(os.path.basename(directoryChosenFile))
        if len(directoryChosenFile) > 20:
            directoryExcelFileForDisplay.set(directoryChosenFile[:20] + '...')
        else:
            directoryExcelFileForDisplay.set(directoryChosenFile)


# Листы ексел файла
directoryExcelFile = customtkinter.StringVar()
directoryExcelFile.set("Файл не выбран")

# Переменная excel файла для отображения
directoryExcelFileForDisplay = customtkinter.StringVar()
directoryExcelFileForDisplay.set('файл не выбран')

# Переменная excel файла для конвертации
directoryExcelFileForConvert = ''

label1 = customtkinter.CTkLabel(tabView.tab("из excel в xml"),
                                text="Выбор excel файла:",
                                width=20,
                                height=20)
label1.place(relx=0.02, rely=0)

# Лейбл для отоброжения, выбранного файла
fileNameForImport = customtkinter.CTkLabel(tabView.tab("из excel в xml"),
                                           textvariable=directoryExcelFileForDisplay,
                                           width=20,
                                           height=20)
fileNameForImport.place(relx=0.02, rely=0.060)

# Кнопка для выбора excel файла
chooseFileForImport = customtkinter.CTkButton(tabView.tab("из excel в xml"),
                                              text="выбрать файл",
                                              command=chooseExcelFileForImport)
chooseFileForImport.place(relx=0.56, rely=0.05)

label2 = customtkinter.CTkLabel(tabView.tab("из excel в xml"),
                                text="Выбор excel листа:",
                                width=20,
                                height=20)
label2.place(relx=0.02, rely=0.13)

# Выпадающий список для выбора необходимого excel листа
listOfExcel = customtkinter.CTkComboBox(tabView.tab("из excel в xml"), values=[''])
listOfExcel.place(relx=0.02, rely=0.18)

label3 = customtkinter.CTkLabel(tabView.tab("из excel в xml"),
                                text="Выбор тегов:",
                                width=20,
                                height=20)
label3.place(relx=0.02, rely=0.26)

# Создание переменной для тегов по умолчанию
DefaultTags = customtkinter.StringVar()
DefaultTags.set("N_ZAP, FAM, IM, OT, W, DR, ENP")

# Строка ввода тегов для формирования xml файла
NameOfTags = customtkinter.CTkEntry(tabView.tab("из excel в xml"), width=323, textvariable=DefaultTags)
NameOfTags.place(relx=0.02, rely=0.32)

label4 = customtkinter.CTkLabel(tabView.tab("из excel в xml"),
                                text="Название xml файла для экспорта:",
                                width=20,
                                height=20)
label4.place(relx=0.02, rely=0.41)

# Строка ввода названия xml файла
NameOfXmlFile = customtkinter.CTkEntry(tabView.tab("из excel в xml"), placeholder_text="название файла без расширения", width=323)
NameOfXmlFile.place(relx=0.02, rely=0.47)

# Функция конвертации из excel файла в xml
def convertFromExcelToXml():
    if directoryExcelFileForConvert == '':
        label6.configure(text='excel файл не выбран', text_color='#FF0000')
        return None
    elif listOfExcel.get() == '':
        label6.configure(text='лист excel файла не выбран', text_color='#FF0000')
        return None
    elif NameOfTags.get() == '':
        label6.configure(text='теги не были написаны', text_color='#FF0000')
        return None
    elif NameOfXmlFile.get() == '':
        label6.configure(text='не указано название xml файла', text_color='#FF0000')
        return None

    if filterStatus.get():
        if errorLogFile.get() == '':
            label6.configure(text='не указано название лог файла', text_color='#FF0000')
            return None

    xlsx2xml(directoryExcelFileForConvert,
             NameOfXmlFile.get(),
             NameOfTags.get().strip(' ').split(', '),
             errorLogFile.get(),
             listOfExcel.get(),
             filterStatus.get())

    label6.configure(text='файл был создан', text_color='#008000')


# Флажок для фильтрации
filterStatus = customtkinter.CTkCheckBox(tabView.tab("из excel в xml"), text='Использовать фильтр', onvalue=True, offvalue=False)
filterStatus.place(relx=0.02, rely=0.56)


# Лейбл для названия лог файла для ошибок
label5 = customtkinter.CTkLabel(tabView.tab("из excel в xml"), text="Название лог файла:").place(relx=0.02, rely=0.63)

# Название лог файла для ошибок
errorLogFile = customtkinter.CTkEntry(tabView.tab("из excel в xml"), placeholder_text='название файла без расширения', width=323)
errorLogFile.place(relx=0.02, rely=0.7)

# Кнопка для конвертации
convertButton = customtkinter.CTkButton(tabView.tab("из excel в xml"),
                                        text="конвертировать",
                                        width=323,
                                        command=convertFromExcelToXml)
convertButton.place(relx=0.02, rely=0.8)

# Лейбл отображения статуса
label6 = customtkinter.CTkLabel(tabView.tab("из excel в xml"), width=300, text='', font=('bold', 18))
label6.place(relx=0.02, rely=0.9)


# Вторая вкладка

# Функция выбора xml файла
def chooseXmlFile():
    # Переменная для получения пути xml файла
    directoryXmlFile = filedialog.askopenfilename()

    if directoryXmlFile != '':
        global nameOfXmlFile

        # Проверка xml формата
        if os.path.splitext(directoryXmlFile)[1] != '.xml':
            label9.configure(text='был выбран файл не xml формата', text_color='#FF0000')
            return None

        # Получение абсолютного пути xml файла
        nameOfXmlFile = directoryXmlFile[:]

        # Получение названия файла для отображения
        directoryXmlFile = str(os.path.basename(nameOfXmlFile))
        if len(directoryXmlFile) > 20:
            NameOfXmlFileForDisplay.set(directoryXmlFile[:20] + '...')
        else:
            NameOfXmlFileForDisplay.set(directoryXmlFile)


# Переменная названия файла для отображения
NameOfXmlFileForDisplay = customtkinter.StringVar()
NameOfXmlFileForDisplay.set('файл не выбран')

# Переменная хранения абсолютного пути
nameOfXmlFile = ''


label7 = customtkinter.CTkLabel(tabView.tab("из xml в excel"), text='Выбор xml файла:')
label7.place(relx=0.02, rely=0)

label8 = customtkinter.CTkLabel(tabView.tab("из xml в excel"), textvariable=NameOfXmlFileForDisplay)
label8.place(relx=0.02, rely=0.06)

# Кнопка для выбора xml файла
chooseXmlFileButton = customtkinter.CTkButton(tabView.tab("из xml в excel"), text='выбрать файл', command=chooseXmlFile)
chooseXmlFileButton.place(relx=0.56, rely=0.05)

label9 = customtkinter.CTkLabel(tabView.tab("из xml в excel"), text='Название excel файла:')
label9.place(relx=0.02, rely=0.12)

# Название xml файла
nameOfExcelFile = customtkinter.CTkEntry(tabView.tab("из xml в excel"), placeholder_text='название файла без расширения', width=323)
nameOfExcelFile.place(relx=0.02, rely=0.19)


# Функция конвертации из xml в excel формат
def convertFromXmlToExcel():

    if nameOfXmlFile == '':
        label9.configure(text='xml не был выбран', text_color='#FF0000')
        return None
    elif nameOfExcelFile.get() == '':
        label9.configure(text='название excel файла не было указано', text_color='#FF0000')
        return None

    # Проверка xml формата
    if os.path.splitext(nameOfXmlFile)[1] != '.xml':
        label9.configure(text='был выбран файл не xml формата', text_color='#FF0000')
        return None

    xml2xlsx(nameOfXmlFile, nameOfExcelFile.get())

    label9.configure(text='файл был создан', text_color='#008000')

# Кнопка конвертации xml файла в excel формат
convertFromXmlToExcelButton = customtkinter.CTkButton(tabView.tab("из xml в excel"), text='конвертировать', command=convertFromXmlToExcel, width=323)
convertFromXmlToExcelButton.place(relx=0.02, rely=0.28)

# Лейбл отображения статуса
label9 = customtkinter.CTkLabel(tabView.tab("из xml в excel"), text='', width=300, font=('bold', 18))
label9.place(relx=0.02, rely=0.38)

app.mainloop()

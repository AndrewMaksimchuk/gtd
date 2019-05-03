import re
from openpyxl import *
wb = Workbook()
ws = wb.create_sheet("Format_gtd", 0)
Format_gtd = wb["Format_gtd"]

Format_gtd.cell(row=1, column=1, value="Опис товара")
Format_gtd.cell(row=1, column=2, value="Мито")
Format_gtd.cell(row=1, column=4, value="Мито, %")
Format_gtd.cell(row=1, column=5, value="Митна вартість, грн")
Format_gtd.cell(row=1, column=6, value="Сума мита")
Format_gtd.cell(row=1, column=7, value="Сума НДС")
Format_gtd.column_dimensions['A'].width = 30
Format_gtd.column_dimensions['B'].width = 15
Format_gtd.column_dimensions['D'].width = 15
Format_gtd.column_dimensions['E'].width = 20
Format_gtd.column_dimensions['F'].width = 15
Format_gtd.column_dimensions['G'].width = 15

# Відкриваємо файл для зчитування інформації
newGtdFile = input("Enter name of GTD file: ")
newGtdFile = newGtdFile + ".xlsx"
loadWorkbook = load_workbook(newGtdFile)
sheet = loadWorkbook.active

# Рядок з якого починаємо записувати дані у новий файл
startRow = 2
# Максимальна кількість рядків у відкритому файлі
maxRowRange = sheet.max_row - 1

# Словник з відсотками мита
toll = {}
# Значення мита у відсотках
tollValue = ''

for i in range(6, maxRowRange):
    # Отримуємо назви товарів
    nameProdact = sheet.cell(row=i, column=5).value
    # Список назв товарів з однаковою відсотковою ставкою
    # listNameProdact = re.findall("[- [A-Z0-9-./]+ - [0-9]+ [шт;|м;]+]*", nameProdact) #Початковий варіант регулярного виразу, є недопрацювання, помилка
    # listNameProdact = re.findall("[- [A-Z0-9-./]+ - [0-9]шт;", nameProdact)
    listNameProdact = re.findall("- [A-Z0-9-./ ]+шт;", nameProdact)
    
    # Перебираємо список товарів і кожному вказуємо відповідний відсоток
    for x in listNameProdact:
        Format_gtd.cell(row=startRow, column=1, value=x)

        # Отримуємо відсотки мита
        muto = str(sheet.cell(row=i, column=16).value * 100) + "%"
        if len(muto) > 10:
            pos = muto.find("%")
            finalString = muto[0:pos + 1]
            tollValue = finalString
            Format_gtd.cell(row=startRow, column=2, value=finalString)
        else:
            tollValue = muto
            Format_gtd.cell(row=startRow, column=2, value=muto)
        startRow = startRow + 1

    # Отримуємо значення таможньої вартості, 0 - позиція у списку
    customsValue = sheet.cell(row=i, column=15).value
    # Отримуємо значення суми мита, 1 - позиція у списку
    amountOfDuty = sheet.cell(row=i, column=17).value
    # Отримуємо значення суми НДС, 2 - позиція у списку
    amountOds = sheet.cell(row=i, column=19).value

    # Умова яка перевіряє чи існує ключ з таким значенням мита у словнику
    if tollValue in toll:
        toll[tollValue][0] = toll[tollValue][0] + customsValue
        toll[tollValue][1] = toll[tollValue][1] + amountOfDuty
        toll[tollValue][2] = toll[tollValue][2] + amountOds
    else:
        toll[tollValue] = []
        toll[tollValue].append(customsValue)
        toll[tollValue].append(amountOfDuty)
        toll[tollValue].append(amountOds)

startRow2 = 2
for b in toll:
    Format_gtd.cell(row=startRow2, column=4, value=b)
    Format_gtd.cell(row=startRow2, column=5, value=toll[b][0])
    Format_gtd.cell(row=startRow2, column=6, value=toll[b][1])
    Format_gtd.cell(row=startRow2, column=7, value=toll[b][2])
    startRow2 = startRow2 + 1

wb.save("gtd.xlsx")
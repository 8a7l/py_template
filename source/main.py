#===================================================
# Підключенні бібліотеки.
#===================================================
import openpyxl
import os
#===================================================


#===================================================
# Головні змінні.
#===================================================
file_way=os.path.dirname(__file__)
file_name=os.path.basename(__file__)
file_name=file_name.split('.')
file_name=file_name[0]
database_excel=file_name+'.xlsx'
#===================================================


#===================================================
# Функції
#===================================================
def database_excel_open():
    workbook = openpyxl.load_workbook(database_excel)
    workbook_sheet=workbook['sheet1']
    a1=workbook_sheet['A1'].value
    a2=workbook_sheet['A2'].value
    a3=workbook_sheet['A3'].value
    print(a1)


def database_excel_create():
    workbook = openpyxl.Workbook()
    workbook_sheet=workbook.active
    workbook_sheet.title = 'sheet1'
    workbook_sheet['A1'] = 1
    workbook_sheet['A2'] = 2
    workbook_sheet['A3'] = 3
    workbook.save(database_excel)


def fail_message():
    print('Виявлена помилка.')
    print('Всі налаштування скинуто до початкових.')
    input()

#===================================================


#===================================================
# Точка входу.
#===================================================
def main():
    try:
        database_excel_open()
    except:
        database_excel_create()
        fail_message()


if __name__ == '__main__':
    main()
#===================================================

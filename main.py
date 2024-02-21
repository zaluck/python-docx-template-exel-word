import openpyxl
from docxtpl import DocxTemplate

wb = openpyxl.load_workbook(filename='staff.xlsx')
sheet = wb['staff']

doc = DocxTemplate('vacation_template.docx')

for num in range(2, len(list(sheet.rows)) + 1):
    name = sheet['B' + str(num)].value
    last_name = sheet['A' + str(num)].value
    company = sheet['D' + str(num)].value
    start_data = sheet['F' + str(num)].value.date()
    end_data = sheet['G' + str(num)].value.date()

    context = {
        'company': company,
        'name': name,
        'last_name': last_name,
        'start_data': start_data,
        'end_data': end_data
    }

    doc.render(context)
    doc.save(last_name + ' заявление на отпуск.docx')

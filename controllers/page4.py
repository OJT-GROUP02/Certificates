import openpyxl
import psycopg2
import itertools

from operator import itemgetter
from openpyxl.drawing.image import Image
from datetime import datetime as dt
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, fills, numbers, PatternFill, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.descriptors.excel import UniversalMeasure, Relation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


wb = Workbook()

wb.create_sheet("Page 4", 0)

ws = wb["Page 4"]
ColumnDimension(ws, auto_size=True)
ws.page_setup.paperHeight = '13in'
ws.page_setup.paperWidth = '8.5in'
ws.page_margins.left = 0.50
ws.page_margins.rigt = 0.50
document_font = Font(name="Times New Roman", size=12)

# column sizes
ws.column_dimensions["A"].width = 5
ws.column_dimensions["B"].width = 32
ws.column_dimensions["C"].width = 9
ws.column_dimensions["D"].width = 9
ws.column_dimensions["E"].width = 12
ws.column_dimensions["F"].width = 25

# Image
img = Image(r"static/images/bicol-university-logo.png")
img.anchor = 'B1'
img.width = 105
img.height = 105
ws.add_image(img)

# Header
ws.append(['Republic of the Philippines'])
ws.merge_cells('A1:H1')
ws.append(['Bicol University'])
ws.merge_cells('A2:H2') 
ws.append(['OFFICE OF THE UNIVERSITY REGISTRAR'])
ws.merge_cells('A3:H3')
ws['A3'].font = Font(name="Times New Roman", size=12, bold=True)
ws.append(['Legazpi City'])
ws.merge_cells('A4:H4')

ws.append(['Tel. No (052) 820-6809'])
ws.merge_cells('A5:H5')
ws.append(['E-mail Add: bu_uro@yahoo.com'])
ws.merge_cells('A6:H6')

for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(
            wrap_text=True, horizontal='center', vertical='center')
        cell.font = document_font

ws['B7'].value = "ISO 9001:2008"
ws['B7'].font = Font(color = "57C4E5")
ws['B7'].alignment = Alignment(horizontal='left')

ws['B8'].value = "Certificate No."
ws['B8'].font = Font(color = "57C4E5")
ws['B8'].alignment = Alignment(horizontal='left')

ws['B9'].value = "TUV100 05 1782"
ws['B9'].font = Font(color = "57C4E5")
ws['B9'].alignment = Alignment(horizontal='left')

ws['A12'].value = "C E R T I F I C A T I O N"
ws.merge_cells('A12:H13')
ws['A12'].alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
ws['A12'].font = Font(size='23', bold=True)

ws['B15'].value = 'TO WHOM IT MAY CONCERN:'
ws['B15'].font = Font(name="Times New Roman", size=12, bold=True)
ws['A15'].alignment = Alignment(horizontal='left')

# End of Header

# Database Connection

cur = psycopg2.connect(database='certification_db', user='postgres',
                        password='1612', host='localhost',
                        port="5432").cursor()

student_list = []
cur.execute('SELECT * from undergrad_stud')
for student in cur:
    stud_row = list(student)
    student_list.append(stud_row)

course_list = []
cur.execute('SELECT * from degree_courses')
for course in cur:
    course_row = list(course)
    course_list.append(course_row)

major_list = []
cur.execute('SELECT * from majors')
for major in cur:
    major_row = list(major)
    major_list.append(major_row)

award_list = []
cur.execute('SELECT * from awards')
for award in cur:
    award_row = list(award)
    award_list.append(award_row)

reg_list = []
cur.execute('SELECT * from registrar')
for reg in cur:
    reg_row = list(reg)
    reg_list.append(reg_row)

# Body

# 1st Paragraph
ws['B18'].value = '         This is to certify that Ms. ' + student_list[2][1].upper() + ' Q. ' + student_list[2][3].upper() + ' has graduated with the degree'
ws.merge_cells('B18:H18')
ws['B19'].value = 'of ' + course_list[2][1] + ' (' + course_list[2][2] + ') ' +'major in ' + major_list[2][2] + ','
ws.merge_cells('B19:H19')
ws['B20'].value = award_list[2][1] + ' on ' + 'April 03, 1997 per Resolution No. 1, s. 1997 of the Board of Regents,'
ws.merge_cells('B20:H20')
ws['B21'].value = 'Bicol University.'
ws.merge_cells('B21:H21')

# 2nd Paragraph
ws['B23'].value = '         It is further certified that she is of good moral character and has never been'
ws.merge_cells('B23:H23')
ws['B24'].value = 'subjected to any disciplinary action during her entire stay in this University.'
ws.merge_cells('B24:H24')

# 3rd Paragraph
ws['B26'].value = '         Issued this 21st day of June, 2010 upon the request of Ms. ' + student_list[2][3] + ' for reference'
ws.merge_cells('B26:H26')
ws['B27'].value = 'purposes.'
ws.merge_cells('B27:H27')

# Registrar
ws['F32'].value = reg_list[0][1]
ws['F33'].value = reg_list[0][2]

# Footer
ws['B38'].value = 'BU-F-UREG-54'
ws['G38'].value = 'Revision: 1'
ws['B39'].value = 'Effectivity Date: Mar. 9,2011'

wb.save('static/Certificate_Page4.xlsx')
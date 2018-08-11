import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, colors, Alignment


wb = Workbook()
sheet = wb.active
sheet.title = "new sheet"
sheet['C3'] = 'hello'


bold_itatic24_font = Font(name='等线', size=24, italic=True, color = colors.RED, bold=True)
sheet['A1'].font = bold_itatic24_font
# 设置B1中的数据垂直居中和水平居中
sheet['B1'].alignment = Alignment(horizontal = 'center', vertical = 'center')
# 第2行行高
sheet.row_dimensions[2].height = 40
# C列列宽
sheet.column_dimensions['C'].width = 40

for i in range(10):
    sheet["A%d" % (i+1)].value = i+1
sheet["E1"].value = "=SUM(A:A)"

sheet.merge_cells('B1:G1') # 合并一行中的几个单元格
sheet.merge_cells('A1:C3') # 合并一个矩形区域中的单元格

wb.save("saveNew.xlsx")






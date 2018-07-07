import xlwt
# workbook = xlrd.open_workbook()
workbook = xlwt.Workbook('input.xlsx')
worksheet = workbook.add_sheet('My Worksheet')
font = xlwt.Font() # Create the Font
font.name = 'Times New Roman'
font.underline = all
style = xlwt.XFStyle() # Create the Style
style.font = font # Apply the Font to the Style
worksheet.write(0, 0, label = 'one')
worksheet.write(1, 0, label = 'two', style=style) # Apply the Style to the Cell
# worksheet.write(1, 0, label = 'Formatted value', style = style) # Apply the Style
style = xlwt.easyxf('font: bold 1,height 280;')
workbook.save('fontxl.xls')
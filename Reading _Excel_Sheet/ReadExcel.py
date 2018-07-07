import xlrd
import pandas as pd
import xlwt
# file_location = open("C:\Users\S530459\Desktop\GDP-1\Reading _Excel_Sheet\input.xlsx")
# file_location = pd.read_excel("C:/Users/S530459/Desktop/GDP-1/Reading _Excel_Sheet/chandoor_input.xlsx")
file_location = pd.read_excel("chandoor_input.xlsx")
# sheet = file_location.write(1, 1, 'Vipul')
# sheet = xlrd.open_workbook(0)
# sheet.font_list.count
# xlrd.formatting.Font.bold(0)
# xlrd.formatting.EqNeAttrs.
font = xlwt.Font()
font.name = 'Times New Roman'
style = xlwt.XFStyle() # Create the Style
style.font = font # Apply the Font to the Style
style = xlwt.easyxf('font: bold 1,height 280;')
# file_location.write(0, 0, label = 'Unformatted value')
# file_location.write(1, 0, label = 'Formatted value') # Apply the Style to the Cell
# file_location.save('fontxl.xls')

# print(file_location)
# cell_format = 

writer = pd.ExcelWriter("output2.xlsx")
file_location.to_excel(writer)
# writer.formatting.fontxl()
writer.save()
# arr    = file_location.values[:,:-2]    # just the numbers
# offset = file_location.values[:,-1]     # just the offsets
# column_pad = 2
# arr2 = np.zeros( (arr.shape[0],arr.shape[1]+column_pad) )

# writer.file_location


# >>> writer = pd.ExcelWriter('output.xlsx')
# >>> df1.to_excel(writer,'Sheet1')
# >>> df2.to_excel(writer,'Sheet2')
# >>> writer.save()
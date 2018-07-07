import pandas as pd
import datetime
import color
import xlsxwriter 
file_location = pd.read_excel("Chandoor_Input.xlsx")
file_location = file_location.sort_values(by=['Genre','Critic Score'],ascending=[True,False])
file_location['Release Date'] = file_location['Release Date'].dt.strftime('%d-%b-%y')
i = 1
workbook = xlsxwriter.Workbook('Chandoor_Output.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:A', 5)
worksheet.set_column('B:B', 14)
worksheet.set_column('C:C', 10)
worksheet.set_column('D:D', 35)
worksheet.set_column('E:E', 20)
worksheet.set_column('F:F', 14)
for elementindex in file_location.index:
    file_location.loc[elementindex,'SNO'] = i
    i = i + 1
row = 3
col = 0
result1 = ['SNO', 'Genre', 'Credit Score', 'Album  Name', 'Artist', 'Release Date']
result = [0,2,5,1,3,4]
data_format1 = workbook.add_format({
    'border': 1,
    'align': 'center',
    'bg_color': '#C00000',
    'bold' : 1,
    'color': 'white'
    })
data_format2 = workbook.add_format({
    'border': 1,
    'bg_color': '#FF2CC'
    })
data_format3 = workbook.add_format({
    'border': 1,
    'bg_color': '#C6E0B4'
    })
data_format2_col5 = workbook.add_format({
    'border': 1,
    'bg_color': '#FF2CC',
    'align': 'right'
    })
data_format3_col5 = workbook.add_format({
    'border': 1,
    'bg_color': '#C6E0B4',
    'align': 'right'
    })
data_format4 = workbook.add_format({
    'border': 1,
    'bg_color': '#FF2CC',
    'align':'right'
    })
merge_format = workbook.add_format({
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'italic': 1,
    'underline': 1,
    'fg_color': '#C6E0B4'})
merge_format1 = workbook.add_format({
    'border': 1,
    'align': 'left',
    'bold' : 1,
    'valign': 'vcenter',
    'italic': 0,
    'underline': 0,
    'fg_color': '#C6E0B4'})
worksheet.write(1,1, 'Name', merge_format1)
worksheet.merge_range('C2:D2', 'Vipul,Chandoor', merge_format)
oldGenre = ''
for element in file_location.values:
    if row == 3:
        oldGenre = element[2]
        for ele in result1:
           worksheet.write(row,col, ele, data_format1)
           col = col + 1
        row = row + 1
        col = 0
    if oldGenre != element[2]:
        data_format2,data_format3 = data_format3,data_format2
        data_format2_col5,data_format3_col5 = data_format3_col5,data_format2_col5
        oldGenre = element[2]
    while col < element.size:
        if col == 5:
            worksheet.write(row,col, element[result[col]], data_format2_col5)
        else:
            worksheet.write(row,col, element[result[col]], data_format2)
        col = col + 1
    col = 0
    row = row + 1
workbook.close()
print("Output file generated successfully")


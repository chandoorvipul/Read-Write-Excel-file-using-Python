import pandas as pd
import datetime
import color
import xlsxwriter 
file_location = pd.read_excel("chandoor_input.xlsx")
# mod_df = ws_dict['existing_worksheet']
file_location = file_location.sort_values(by=['Genre','Critic Score'],ascending=[True,False])
# file_location = file_location.sort_values(['Critic Score', ascending=False])
# sort(['a', 'b'], ascending=[True, False])
file_location['Release Date'] = file_location['Release Date'].dt.strftime('%d-%b-%y')
i = 1
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:A', 5)
worksheet.set_column('B:B', 14)
worksheet.set_column('C:C', 14)
worksheet.set_column('D:D', 25)
worksheet.set_column('E:E', 20)
worksheet.set_column('F:F', 16)


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

data_format4 = workbook.add_format({
    'border': 1,
    'bg_color': '#FF2CC',
    'align':'right'
    })
# worksheet.set_row(3, cell_format=data_format1)

# for row in range(0, 29, 2):
#     worksheet.set_row(row, cell_format=data_format2)
#     worksheet.set_row(row + 1, cell_format=data_format3)

# Create a format to use in the merged range.
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
# Merge 3 cells.
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
        oldGenre = element[2]
    while col < element.size:
        if col == 5:
            worksheet.write(row,col, element[result[col]], data_format4)
        else:
            worksheet.write(row,col, element[result[col]], data_format2)
        col = col + 1
    col = 0
    row = row + 1

# A:A = 20 
print(workbook)
workbook.close()
#writer = pd.ExcelWriter("output.xlsx")
# work = writer.sheets['sheet1']
# writer.book = book
# book.write(2, 2, label = 'Name : Vipul chandoor')
# file_location.to_excel(writer, index=False, startrow = 3)
# pd.DataFrame([['Name','Vipul,Chandoor']]).to_excel(writer, startrow = 1, startcol = 1,header= False,index = False,merge_cells=False)
# print(pd.DataFrame([['Name','Vipul,Chandoor']]))
# writer.write(1, 1, 'hello')
# writer.write(1, 1, label = 'one')
# file_locationclose()


#writer.save()

# # Add a format. Light red fill with dark red text.
# format1 = file_location.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
# # Add a format. Green fill with dark green text.    
# format2 = file_location.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
# worksheet.conditional_format(color_range, {'type': 'top', 'value': '5', 'format': format2})
# # Highlight the bottom 5 values in Red
# worksheet.conditional_format(color_range, {'type': 'bottom', 'value': '5', 'format': format1})






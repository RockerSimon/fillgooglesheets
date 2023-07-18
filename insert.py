import gspread
from oauth2client.service_account import ServiceAccountCredentials
import cryptography.exceptions
import sys
import json
import datetime

if len(sys.argv) >= 3:
    arg1 = sys.argv[1]
    arg2 = sys.argv[2]
    arg3 = sys.argv[3] if len(sys.argv) >= 4 else None

# 设置凭证和作用域
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('/opt/simon/positive-notch-390914-f7928ccdf9a9.json', scope)

# 从 JSON 文件中加载数据
with open('taw2.json', 'r') as file:
    json_data = json.load(file)

# 传参为第一个值为 key2 时，对应 JSON 取 item3 的值为 "表格名称"
#param = list(json_data.keys())
#print (param)
value3 = json_data[arg1][0]
value2 = json_data[arg1][1]
value1 = json_data[arg1][2]
# 授权访问并打开目标表格
gc = gspread.authorize(credentials)
sheet = gc.open(value2).sheet1  # 替换为你要操作的表格名称

# 获取最后一行的索引
#last_row = len(sheet.get_all_values()) + 1
last_row = len(sheet.get_all_values()) + 1
print (last_row)
# 获取 ADFG 列的列索引
column_indices = ['A', 'B', 'C', 'D', 'E', 'G', 'H', 'I']

# 获取当前日期并截取年份、月份和日期
today = datetime.datetime.now()
year_month_day = today.strftime("%Y-%m-%d")
print (year_month_day)
# 在最后一行追加数据
#sheet.append_row(data, value_input_option='RAW')

# 将值写入表HDFC Taw 02最后一行的第ABCI列
if value1 == 'HDFC Taw 02':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.5/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-E{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7))
    result6 = sheet.cell(last_row-1,7).value
    result5 = int(float(result6))
    result9 = sheet.cell(last_row,4).value
    cell8 = "H" + str(last_row)
    if result5 < 0:
        sheet.update_acell(cell8, (result9))
    else:
        sheet.update_acell(cell8, (result8))
    print("值已写入表HDFC Taw 02的最后一行的第ABCDEGHI列")

# 将值写入HDFC Monish 02最后一行的第ABCI列
if value1 == 'HDFC Monish 02':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.8/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7))
    cell8 = "H" + str(last_row)
    sheet.update(cell8, (result8))
    print("值已写入表HDFC Taw 02的最后一行的第ABCDEGHI列")

# 将值写入表HDFC Taw最后一行的第ABCI列
if value1 == 'HDFC Taw':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.5/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-E{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7))
    cell8 = "H" + str(last_row)
    sheet.update(cell8, (result8))
    print("值已写入表HDFC Taw的最后一行的第ABCDEGHI列")

# 将值写入表HDFC Fame最后一行的第ABCI列
if value1 == 'HDFC Fame':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.5/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-E{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7))
    result6 = sheet.cell(last_row-1,7).value
    result5 = int(float(result6))
    result9 = sheet.cell(last_row,4).value
    cell8 = "H" + str(last_row)
    if result5 < 0:
        sheet.update_acell(cell8, (result9))
    else:
        sheet.update_acell(cell8, (result8))
    print("值已写入表HDFC Fame的最后一行的第ABCDEGHI列")

# 将值写入表HDFC Jay new最后一行的第ABCI列
if value1 == 'HDFC Jay new':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.2/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-E{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7))
    cell8 = "H" + str(last_row)
    sheet.update(cell8, (result8))
    print("值已写入表HDFC Jay new的最后一行的第ABCDEGHI列")

# 将值写入表HDFC Prath最后一行的第ABCI列
if value1 == 'HDFC Prath':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.2/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7))
    cell8 = "H" + str(last_row)
    sheet.update(cell8, (result8))
    print("值已写入表HDFC Prath的最后一行的第ABCDEGHI列")

# 将值写入表HDFC Pratap最后一行的第ABCI列
if value1 == 'HDFC Pratap':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.5/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-E{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7)/1000)*1000
    cell6 = "F" + str(last_row)
    sheet.update(cell6, (result8))
    result9 = int(float(result7)/1000)
    cell8 = "J" + str(last_row)
    sheet.update(cell8, "will send {}k  pending".format(result9))
    print("值已写入表HDFC Pratap的最后一行的第ABCDEFGIJ列")

# 将值写入表HDFC Arjun最后一行的第ABCI列
if value1 == 'HDFC Arjun':
    cell1 = "A" + str(last_row)
    sheet.update(cell1, (year_month_day))
    cell2 = "B" + str(last_row)
    sheet.update(cell2, (value3))
    cell3 = "C" + str(last_row)
    sheet.update_acell(cell3, (arg2))
#有chargeback填写退款金额，没有传参不填值为none自动跳过
    if arg3 :
      cell9 = "I" + str(last_row)
      sheet.update_acell(cell9, (arg3))
    cell4 = "D" + str(last_row)
    formula4 = '=C{}-I{}'.format((last_row),(last_row))
    sheet.update_acell(cell4, (formula4))
    cell5 = "E" + str(last_row)
    formula5 = '=D{}*1.2/100'.format((last_row))
    sheet.update_acell(cell5, (formula5))
    cell7 = "G" + str(last_row)
    formula7 = '=D{}-E{}-F{}-H{}+G{}'.format((last_row),(last_row),(last_row),(last_row),((last_row)-1))
    sheet.update_acell(cell7, (formula7))
    result7 = sheet.cell(last_row,7).value
    result8 = int(float(result7))
    cell8 = "H" + str(last_row)
    sheet.update(cell8, (result8))
    print("值已写入表HDFC Arjun的最后一行的第ABCDEGHI列")

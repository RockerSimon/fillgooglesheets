from flask import Flask, request, jsonify
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import datetime

app = Flask(__name__)

# Load JSON data
with open('/opt/simon/taw2.json', 'r') as file:
    json_data = json.load(file)

# Set credentials and scope
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('/opt/simon/positive-notch-390914-f7928ccdf9a9.json', scope)

@app.route('/update_sheet', methods=['POST'])
def update_sheet():
    arg1 = request.json.get('arg1')
    arg2 = request.json.get('arg2')
    arg3 = request.json.get('arg3')

    if not arg1 or not arg2:
        return jsonify({'error': 'Missing arguments'}), 400

    # Get values from JSON data
    if arg1 in json_data:
        value3 = json_data[arg1][0]
        value2 = json_data[arg1][1]
        value1 = json_data[arg1][2]
    else:
        return jsonify({'error': 'Invalid argument'}), 400

    # Authorize access and open target spreadsheet
    gc = gspread.authorize(credentials)
    sheet = gc.open(value2).sheet1

    # Get last row index
    last_row = len(sheet.get_all_values()) + 1

    # Get current date
    today = datetime.datetime.now()
    year_month_day = today.strftime("%Y-%m-%d")

    # Update sheet based on value1
    if value1 == 'HDFC Taw 02':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
        if arg3:
            cell9 = "I" + str(last_row)
            sheet.update_acell(cell9, (arg3))
        cell4 = "D" + str(last_row)
        formula4 = '=C{}-I{}'.format(last_row, last_row)
        sheet.update_acell(cell4, (formula4))
        cell5 = "E" + str(last_row)
        formula5 = '=D{}*1.5/100'.format(last_row)
        sheet.update_acell(cell5, (formula5))
        cell7 = "G" + str(last_row)
        formula7 = '=D{}-E{}-F{}-H{}+G{}'.format(last_row, last_row, last_row, last_row, last_row - 1)
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
        return jsonify({'message': 'Values updated in HDFC Taw 02 sheet'}), 200

    # ????HDFC Monish 02??????ABCI?
    if value1 == 'HDFC Monish 02':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
    #?chargeback??????,????????none????
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
        return jsonify({'message': 'Values updated in HDFC Monish 02 sheet'}), 200

    # ?????HDFC Taw??????ABCI?
    if value1 == 'HDFC Taw':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
    #?chargeback??????,????????none????
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
        return jsonify({'message': 'Values updated in HDFC Taw sheet'}), 200
    
    # ?????HDFC Fame??????ABCI?
    if value1 == 'HDFC Fame':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
    #?chargeback??????,????????none????
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
        return jsonify({'message': 'Values updated in HDFC Fame sheet'}), 200
    
    # ?????HDFC Jay new??????ABCI?
    if value1 == 'HDFC Jay new':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
    #?chargeback??????,????????none????
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
        return jsonify({'message': 'Values updated in HDFC Jay new sheet'}), 200
    
    # ?????HDFC Prath??????ABCI?
    if value1 == 'HDFC Prath':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
    #?chargeback??????,????????none????
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
        return jsonify({'message': 'Values updated in HDFC Prath sheet'}), 200
    
    # ?????HDFC Pratap??????ABCI?
    if value1 == 'HDFC Pratap':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
    #?chargeback??????,????????none????
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
        return jsonify({'message': 'Values updated in HDFC Pratap sheet'}), 200
    
    # ?????HDFC Arjun??????ABCI?
    if value1 == 'HDFC Arjun':
        cell1 = "A" + str(last_row)
        sheet.update(cell1, (year_month_day))
        cell2 = "B" + str(last_row)
        sheet.update(cell2, (value3))
        cell3 = "C" + str(last_row)
        sheet.update_acell(cell3, (arg2))
    #?chargeback??????,????????none????
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
        return jsonify({'message': 'Values updated in HDFC Arjun sheet'}), 200



    # Update other sheets based on value1 (HDFC Monish 02, HDFC Taw, HDFC Fame, etc.)
    # ...

    return jsonify({'error': 'Value1 not recognized'}), 400

if __name__ == '__main__':
    app.run()

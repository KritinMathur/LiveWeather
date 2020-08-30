import openpyxl as pyxl
import json

wb = pyxl.Workbook()
ws1 = wb.active
ws2 = wb.create_sheet()

ws1.title = "Main Sheet"
ws2.title = "List of all cities"

ws1['A1'] = 'City ID'
ws1['B1'] = 'City name'
ws1['C1'] = 'Temprature'
ws1['D1'] = 'Humidity'
ws1['E1'] = 'UNIT'
ws1['F1'] = 'Update(0/1)'

ws2['A1'] = 'City ID'
ws2['B1'] = 'City name'

ws1['A2'] = 1261481
ws1['A3'] = 1275339
ws1['A4'] = 1259229
ws1['A5'] = 2643741

ws1['B2'] = 'Delhi'
ws1['B3'] = 'Mumbai'
ws1['B4'] = 'Pune'
ws1['B5'] = 'City of London'

ws1['E2'] =  'C'
ws1['E3'] =  'F'
ws1['E4'] =  'C'
ws1['E5'] =  'F'

ws1['F2'] =  1
ws1['F3'] =  1
ws1['F4'] =  1
ws1['F5'] =  1

with open('bin/city.list.json',encoding = 'utf8') as f :
    data = json.load(f)

    i = 2
    for record in data['data']:
        #print(record['id'],record['name'])
        ws2['A'+str(i)] = record['id']
        ws2['B'+str(i)] = record['name']
        i += 1

wb.save("Project.xlsx")

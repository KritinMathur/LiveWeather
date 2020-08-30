import xlwings as xw
import openpyxl as pyxl
import requests, json
import time

update_time = float(input("Updation Time Interval(in sec):- "))

def openWeatherApiRequest(city_id):

    api_key = "2c6acf18ea619a0cfc79a6f4fec25c20"
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    complete_url = base_url + "appid=" + api_key + "&id=" + city_id
    response = requests.get(complete_url) 
    x = response.json()
    
    if x["cod"] != "404":

        y = x["main"]
        current_temperature = y["temp"] 
        current_humidity = y["humidity"]
        return (current_temperature,current_humidity)
    else: 
        raise Exception("no city found") 



wb = xw.Book('Project.xlsx')
sht1 = wb.sheets['Main Sheet']


row_count = pyxl.load_workbook('Project.xlsx').worksheets[0].max_row

while True:
    for i in range(2,row_count+1):

        if(sht1.range('F'+str(i)).value==1):
        
            tempInk,hum = openWeatherApiRequest(str(int(sht1.range('A'+str(i)).value)))

            if(sht1.range('E'+str(i)).value == 'C'):
                fintemp = tempInk - 273.15
            elif(sht1.range('E'+str(i)).value == 'F'):
                fintemp = (tempInk - 273.15)*(9/5) +  32
            else:
                raise Exception('Unit not specified')

            fintemp = round(fintemp,2)

            sht1.range('C'+str(i)).value = fintemp
            sht1.range('D'+str(i)).value = hum  

    time.sleep(update_time)

import requests as rq
import xlwings as xw
import time 

def cel_to_fer(cel):
    fer = ((cel * 9)/5) + 32
    return fer

def fer_to_cel(fer):
    fer = ((fer - 32)*5)/9
    return fer

CITIES = ["Pune", "Jabalpur", "Delhi", "Jaipur", "Tokyo", "Rio", "Berlin", "Moscow", "Denver", "Nairobi", "Helsinki", "Oslo", "Lisbon", "Stockholm"]
length = len(CITIES)
Unit, Updation, convtTemp = ['C']*length, [1]*length, [0]*length
temp, changeTemp=[],[0]*length

for city in CITIES:
    url = "http://api.openweathermap.org/data/2.5/weather?q={}&appid=48fd523d81cd14d6b41f190d770cba6d".format(city)
    data = rq.get(url).json()
    temp.append(data['main']['temp'] - 273.15)


# excelSheets = xw.Book()
# excelSheets.save('Weather Report.xlsv')

excelSheets = xw.Book('Weather Report.xlsv')
sheet1 = excelSheets.sheets[0]
sheet2 = excelSheets.sheets[1]

sheet1.range('A1').value = ['City','Temp','Unit','Updation']

sheet1.range('A2').options(transpose=True).value = CITIES
sheet1.range('B2').options(transpose=True).value = temp
sheet1.range('C2').options(transpose=True).value = Unit
sheet1.range('D2').options(transpose=True).value = Updation

j=0
while(j<100):

	
	for i in range(len(CITIES)):
		
		if(('F' == sheet1.range('C'+str(i+2)).value or 'f' == sheet1.range('C'+str(i+2)).value) and (convtTemp[i] == 0)):
			convtTemp[i] = 1
			temp[i] = cel_to_fer(temp[i])

		elif(('C' == sheet1.range('C'+str(i+2)).value or 'c' == sheet1.range('C'+str(i+2)).value) and (convtTemp[i] == 1)):
			convtTemp[i] = 0
			temp[i] = fer_to_cel(temp[i])

		if(sheet1.range('D'+str(i+2)).value == 1 ):
			url	= "http://api.openweathermap.org/data/2.5/weather?q={}&appid=48fd523d81cd14d6b41f190d770cba6d".format(CITIES[i])
			data = rq.get(url).json()
			temp[i] = data['main']['temp'] - 273.15


			if(((sheet1.range('C'+str(i+2)).value == 'F') or (sheet1.range('C'+str(i+2)).value) == 'f') and (changeTemp[i] == 0)):
				changeTemp[i] = 0
				temp[i] = cel_to_fer(temp[i])
			elif(('C' == sheet1.range('C'+str(i+2)).value or 'c' == sheet1.range('C'+str(i+2)).value) and (changeTemp[i] == 1)):
				changeTemp[i] = 0
				temp[i] = fer_to_cel(temp[i])

		sheet1.range('B2').options(transpose=True).value = temp



	time.sleep(2.5)
	j=j+1
import openpyxl
import datetime
from datetime import date, timedelta
from dateutil import parser
otchet = openpyxl.load_workbook('Zott - 14.05.xlsx')
moscow = openpyxl.load_workbook('Zott.xlsx')
saint_p = openpyxl.load_workbook('07.05-13.05.18_ZOTT_ОТЧЕТ _СПб.xlsx')

victoria = otchet['Виктория']
victoria_m = moscow['Виктория']
result = otchet['Отчёт']

today = date.today()
result.cell(row = 2, column = 3, value = today)
#print(yesterday.strftime('%d.%m.%y'))
print(victoria.cell(row=26,column=13).value)

for i in range (2,28):
    victoria.cell(row=i, column = 13, value = victoria_m.cell(row=i, column=5).value.date())
#lenta.cell(row=65, column=30, value='=COUNTIF(O3:O62'+',"='+ yesterday.strftime('%d.%m.%y')+'"')

for k in range (2,28):
    for m in range(15,26):
        victoria.cell(row=k, column=m, value=victoria_m.cell(row=k,column=m-8).value)


lenta = otchet['Лента']
lenta_m = moscow['Лента']

for j in range (2,11):
    lenta.cell(row=j, column=13, value = lenta_m.cell(row=j, column=5).value.date())

for l in range (2,11):
    for a in range (15,22):
        lenta.cell(row=l, column=a, value = lenta_m.cell(row=l, column=a-8).value)

globus = otchet['ГиперГлобус']
globus_m = moscow['ГиперГлобус']

for b in range (2,8):
    globus.cell(row=b, column=13, value=globus_m.cell(row=b, column=5).value.date())

for c in range (2,8):
    for i in range (15,22):
        globus.cell(row=c, column=i, value=globus_m.cell(row=c, column=i-8).value)

karusel = otchet['Карусель']
karusel_m = moscow['Карусель']

for b in range (2,23):
    karusel.cell(row=b, column=12, value=karusel_m.cell(row=b, column=5).value.date())

for c in range (2,23):
    for i in range (14,26):
        karusel.cell(row=c, column=i, value=karusel_m.cell(row=c, column=i-7).value)

metro = otchet['Метро']
metro_m = moscow['Метро']

for b in range (2,20):
    metro.cell(row=b, column=13, value=metro_m.cell(row=b, column=5).value.date())

for c in range (2,20):
    for i in range (15,31):
        metro.cell(row=c, column=i, value=metro_m.cell(row=c, column=i-8).value)

perek = otchet['Перекрёсток']
perek_m = moscow['Перекрёсток']

for b in range (2,80):
    perek.cell(row=b, column=13, value=perek_m.cell(row=b, column=6).value.date())

for c in range (2,80):
    for i in range (15,21):
        perek.cell(row=c, column=i, value=perek_m.cell(row=c, column=i-8).value)

okay = otchet['Окей']
okay_m = moscow['Окей']

for b in range (2,11):
    okay.cell(row=b, column=13, value=okay_m.cell(row=b, column=5).value.date())

for c in range (2,11):
    for i in range (15,20):
        okay.cell(row=c, column=i, value=okay_m.cell(row=c, column=i-8).value)

'''for s in range (2,29):
	for t in range (5,18):
		lenta.cell(row=s, column=t, value= lenta_m.cell(row=s+34, column=t).value)
'''


#print(lenta.cell(row=38, column=15).value.date())

#test_parser = parser.parse(lenta.cell(row=38, column=15).value, dayfirst=True)
#print(test_parser)
#lenta.cell(row=7, column =15, value = str(lenta.cell(row=7, column=15).value[1:]))

#for l in range (3,30):
#    print(lenta_m.cell(row=l, column=7).value.strftime("%d.%m.%y"))
otchet.save('test.xlsx')

#print(type(lenta_m.cell(row=3, column=7).value))




#dates_lenta_m=[]
#for l in range (2,30):
    #dates_lenta_m+= datetime.strptime(lenta_m.cell(row=l, column=7).value.strftime("%d.%m.%Y"),'%d.%m.%Y')
    #print(dates_lenta_m)

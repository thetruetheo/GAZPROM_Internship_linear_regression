#при разработке использовалась сторонняя библиотека openpyxl
from openpyxl import load_workbook

#Вставьте абсолютный путь к .xlsx-файлу в ковычки
source=load_workbook('example.xlsx')

active_worksheet=source.active
counter_x=2
counter_y=2
x_value=0
y_value=0
x_avg=0
x_summ=0
y_avg=0
y_summ=0
numerator=0
denominator=0
result=0

while(active_worksheet["A"+str(counter_x)].value!=None):
    x_avg+=active_worksheet["A" + str(counter_x)].value
    counter_x+=1
x_avg=(x_avg/(counter_x-2))

counter_x=2

while(active_worksheet["B"+str(counter_y)].value!=None):
    y_avg+=active_worksheet["B" + str(counter_y)].value
    counter_y+=1
y_avg=(y_avg/(counter_y-2))

counter_y=2


while(active_worksheet["A"+str(counter_x)].value!=None and active_worksheet["B"+str(counter_y)].value!=None):
    numerator+=((active_worksheet["A"+str(counter_x)].value-x_avg)*(active_worksheet["B"+str(counter_y)].value-y_avg))
    counter_x+=1
    counter_y+=1

counter_x=2
counter_y=2
while(active_worksheet["A"+str(counter_x)].value!=None):
    denominator+=((active_worksheet["A"+str(counter_x)].value-x_avg)**2)
    counter_x+=1
counter_x=2

result=numerator/denominator
print("Результат вычислений: "+str(result))
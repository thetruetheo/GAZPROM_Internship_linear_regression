#при разработке использовалась сторонняя библиотека openpyxl
from openpyxl import load_workbook

#Вставьте абсолютный путь к .xlsx-файлу в ковычки
source=load_workbook('example.xlsx')

active_worksheet=source.active
counter=2
x_value=0
y_value=0
a_value=0
b_value=0
while(active_worksheet["A"+str(counter)].value!=None and active_worksheet["B"+str(counter)].value!=None):
    x_value = active_worksheet["A" + str(counter)].value
    y_value = active_worksheet["B" + str(counter)].value
    a_value = active_worksheet["C" + str(counter)].value
    if(x_value==0 and a_value!=None):
        print("\nВходные данные: X = " + str(x_value) + " Y = " + str(y_value)+ " A = " + str(a_value))
        print("Коэффициент X = 0, Коэффициент A = "+str(a_value))
        counter+=1
        continue

    if(x_value==0 and a_value==None):
        print("\nНекорректный состав данных в документе, пожалуйста, проверьте содержимое файла на корректность.")
        break
    elif(x_value==None and a_value!=None):
        print("\nНекорректный состав данных в документе, пожалуйста, проверьте содержимое файла на корректность.")
        break
    elif(x_value==None and y_value!=None):
        print("\nНекорректный состав данных в документе, пожалуйста, проверьте содержимое файла на корректность.")
        break
    elif (x_value != None and y_value == None):
        print("\nНекорректный состав данных в документе, пожалуйста, проверьте содержимое файла на корректность.")
        break
    b_value=y_value/x_value
    print("\nВходные данные: X = "+str(x_value)+" Y = "+str(y_value))
    print("Результаты вычислений: Коэффициент B = "+str(b_value))
    counter+=1
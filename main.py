import pandas as pd
from openpyxl import load_workbook
import os

#Direccion = input("Pegue la dirección a su archivo, ejemplo: \n /Users/rodrigovillanueva/Desktop/    ")
os.chdir ('/Users/rodrigovillanueva/Desktop/')
#leer1 = input('Ponga nombre de archivo ')
libro = load_workbook('archivo.xlsx')
# obtenemos la pestaña/hoja activa (nada mas abrir es la primera)
hoja = libro.active

#nombrehoja = input("Ponga el nombre de la hoja: ")
hoja = libro['Hoja1']
a= int(input("escribe la fila inicial "))
b= int(input("escribe la fila final "))

for i in range(a,b):
    d= "A"+str(i+1)
    av = "B" + str(i + 1)
    j = "F" + str(i + 1)
    celda0 =(hoja[d].value)
    celda1=bool(celda0)
    if celda1 == False:
        celda = 0
    else:
        celda = celda0

    celda00 = (hoja[av].value)
    celda4 = bool(celda00)
    if celda4 == False:
        celda3=0
    else:
        celda3= celda00

    totales= (celda +celda3)
    hoja[j] = totales


libro =libro.save('/Users/rodrigovillanueva/Desktop/archivo.xlsx')










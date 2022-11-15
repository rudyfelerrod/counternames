from openpyxl import load_workbook
import os

Direccion = input("Pegue la dirección a su archivo, ejemplo: \n /Users/rodrigovillanueva/Desktop/    ")
os.chdir (Direccion)
leer1 = input('Ponga nombre de archivo ')
libro = load_workbook(leer1)
# obtenemos la pestaña/hoja activa (nada mas abrir es la primera)
hoja = libro.active

nombrehoja = input("Ponga el nombre de la hoja: ")
hoja = libro[nombrehoja]
a=3
#a= int(input("escribe la fila inicial "))
b= int(input("escribe la fila final "))

for i in range(a,b):
    d= "U"+str(i+1)
    j = "AJ" + str(i + 1)
    d0 = (hoja[d].value)
    d2 = bool(d0)
    if d2 == True:
        d1 = d0
    else:
        continue
    d6 = d1.split(' ')
    x = d6
    fh = '*'
    counter = 0
    diccionario_frecuencias = dict()
    for line in fh:
        words = line.rstrip().split("\n")
        for word in words:
            diccionario_frecuencias[word] = counter
            words=""
    palabras = x
    for palabra in palabras:
        if palabra in diccionario_frecuencias:
            diccionario_frecuencias[palabra]+=1



    for palabra in diccionario_frecuencias:
        frecuencia = diccionario_frecuencias[palabra]




    hoja[j] = sum(diccionario_frecuencias.values())


    for i in range(a, b):
        d = "V" + str(i + 1)
        j1 = "AK" + str(i + 1)
        d0 = (hoja[d].value)
        d2 = bool(d0)
        if d2 == True:
            d1 = d0
        else:
            continue
        d6 = d1.split(' ')
        x = d6

        fh = '*'
        counter = 0
        diccionario_frecuencias = dict()
        for line in fh:
            words = line.rstrip().split("\n")
            for word in words:
                diccionario_frecuencias[word] = counter
                words = ""
        palabras = x
        for palabra in palabras:
            if palabra in diccionario_frecuencias:
                diccionario_frecuencias[palabra] += 1

        for palabra in diccionario_frecuencias:
            frecuencia = diccionario_frecuencias[palabra]


        hoja[j1] = sum(diccionario_frecuencias.values())


        for i in range(a, b):
            d = "W" + str(i + 1)
            j2 = "AL" + str(i + 1)
            d0 = (hoja[d].value)
            d2 = bool(d0)
            if d2 == True:
                d1 = d0
            else:
                continue
            d6 = d1.split(' ')
            x = d6
            import os

            fh = '*'
            counter = 0
            diccionario_frecuencias = dict()
            for line in fh:
                words = line.rstrip().split("\n")
                for word in words:
                    diccionario_frecuencias[word] = counter
                    words = ""
            palabras = x
            for palabra in palabras:
                if palabra in diccionario_frecuencias:
                    diccionario_frecuencias[palabra] += 1

            for palabra in diccionario_frecuencias:
                frecuencia = diccionario_frecuencias[palabra]
            # print(f"la pabra '{palabra}' aparece {frecuencia} veces")

            hoja[j2] = sum(diccionario_frecuencias.values())

libro =libro.save(Direccion+leer1)

print("Se ha ejecutado con éxito, gracias!")

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
    j = "F" + str(i + 1)
    d0 = (hoja[d].value)
    d2 = bool(d0)
    if d2 == True:
        d1 = d0
    else:
        continue
    d3 = d1.upper()
    d4 = d3.replace(',', " ")
    d5 = d4.replace('\n', " ")
    d6 = d5.split(' ')
    x = d6
    import os
    os.chdir("/Users/rodrigovillanueva/PycharmProjects/pythonProject/Carpeta1")
    fname = "nombres2.txt"
    fh = open(fname)
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
    #print(f"la pabra '{palabra}' aparece {frecuencia} veces")



    total1 = sum(diccionario_frecuencias.values())
    #print('hay ' + str(total1) + ' mujeres')
    hoja[j] = total1

    for i in range(a, b):
        d = "B" + str(i + 1)
        j1 = "G" + str(i + 1)
        d0 = (hoja[d].value)
        d2 = bool(d0)
        if d2 == True:
            d1 = d0
        else:
            continue
        d3 = d1.upper()
        d4 = d3.replace(',', " ")
        d5 = d4.replace('\n', " ")
        d6 = d5.split(' ')
        x = d6
        import os

        os.chdir("/Users/rodrigovillanueva/PycharmProjects/pythonProject/Carpeta1")
        fname = "nombres2.txt"
        fh = open(fname)
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

        total2 = sum(diccionario_frecuencias.values())
        #print('hay ' + str(total2) + ' mujeres')
        hoja[j1] = total2

        for i in range(a, b):
            d = "C" + str(i + 1)
            j2 = "H" + str(i + 1)
            d0 = (hoja[d].value)
            d2 = bool(d0)
            if d2 == True:
                d1 = d0
            else:
                continue
            d3 = d1.upper()
            d4 = d3.replace(',', " ")
            d5 = d4.replace('\n', " ")
            d6 = d5.split(' ')
            x = d6
            import os

            os.chdir("/Users/rodrigovillanueva/PycharmProjects/pythonProject/Carpeta1")
            fname = "nombres2.txt"
            fh = open(fname)
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

            total3 = sum(diccionario_frecuencias.values())
           # print('hay ' + str(total3) + ' mujeres')
            hoja[j2] = total3

            #totalfinal = total1 + total2 + total3



libro =libro.save('/Users/rodrigovillanueva/Desktop/archivo.xlsx')
import csv
import openpyxl # TRABAJA CON HOJAS DE EXCEL

def buscar_palabra(ruta,dicc):
    lista =[] # lista donde se almacenaran las busquedas
    cont = 0 # recorre la lista
    with open(ruta, 'r') as archivo:
        lector_csv = csv.reader(archivo)
    with open(ruta, 'r') as archivo:
        lector_csv = csv.reader(archivo)
        # ESTABLECE LA PRIMER PALABRA A BUSCAR QUE SERA LA SECCION DE LAS COSAS A BUSCAR
        palabra_buscada = dicc[0]
        # RECORRE EL ARCHIVO
        for fila in lector_csv: 
            # BUSCA LA PALABRA CLAVE EN EL ARCHIVO
            if len(fila)>0:
                # print(palabra_buscada)
                if fila[0].find(palabra_buscada)>-1:
                    print(palabra_buscada)
                    if palabra_buscada == dicc[0]:
                        lista.append("Encabezado")
                        cont = 1
                        palabra_buscada = dicc[cont]
                    else:
                        lista.append(fila[1])
                        cont = cont + 1
                        if cont<len(dicc):
                            palabra_buscada = dicc[cont]
                        else:
                            break
    return lista


# DICCIONARIO VACIO CON  VALORES QUE SE BUSCARAN EN LA HOJA CSV
# EL PRIMER VALOR DE CADA LISTA HACE REFERENCIA A LA SECCION DEL ARCHIVO
cpu =["Caja del sistema","Fabricante:","Número de serie:"] # LISTA CON ITEMS NECESESARIOS PARA CPU
procesador = ["Procesador(es) central","Nombre del procesador:", "Frecuencia del procesador original:"]
ram1 =["Ram","Tamaño del módulo:","Fabricante del módulo:","Número de pieza del módulo:"]
ram2 =["Number Of Banks:","Tamaño del módulo:","Fabricante del módulo:","Número de pieza del módulo:"]
disco = ["Unidades de disco","Modelo de unidad:","Número de serie de la unidad:","Capacidad de la unidad:"]
cd = ["DVD","Modelo de unidad:","Número de serie:"]
monitor = ["Monitor","Nombre del monitor:","Nombre del monitor (del fabricante):","Número de serie:"]
red = ["Ethernet","Dirección MAC:"]
# keys = ["Fabricante:","Número de serie:", #CPU
#         "Nombre del procesador:", "Frecuencia del procesador original:", # PROCESADOR
#         "Fabricante del módulo:","Tamaño del módulo:","Número de pieza del módulo:", # RAM
#         "Modelo de unidad:","Capacidad de la unidad:","Número de serie de la unidad:", #Disco
#         "Unidad de cd:", "Serial CD/DVD:", # Unidad de CD
#         "Nombre del monitor:", "Nombre del monitor (del fabricante):","Serial Monitor:",# Monitor
#         "Marca teclado:","Modelo teclado:"# Tecaldo
#         ]

# my_dict = dict.fromkeys(cpu)

ruta = "COMUNICACIONES2.CSV"
dicc_final = buscar_palabra(ruta,cpu)
print(dicc_final)

# CODIGO PARA PONER EN HOJA DE EXCEL

workbook = openpyxl.Workbook()
hoja = workbook.active

mi_lista = ['elemento 1', 'elemento 2', 'elemento 3']

for i, elemento in enumerate(dicc_final):
    hoja.cell(row=i+1, column=1).value = elemento

workbook.save('mi_libro.xlsx')




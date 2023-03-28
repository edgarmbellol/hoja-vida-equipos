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
                if fila[0].find(palabra_buscada)>-1:
                    if palabra_buscada == dicc[0] and dicc[0] != dicc[1]:
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

def escribir_excel(datos,camposExcel):
    # CODIGO PARA PONER EN HOJA DE EXCEL

    workbook = openpyxl.load_workbook('formato.xlsx')
    hoja = workbook['HOJA DE VIDA DE EQUIPOS']

    for i in range(0,len(datos),1):
        #PONER EN LA HOJA DE EXCEL LOS DATOS CORRRESPONDIENTES
        hoja[camposExcel[i]] = datos[i]

    workbook.save('formato.xlsx')
    return

# DICCIONARIO VACIO CON  VALORES QUE SE BUSCARAN EN LA HOJA CSV
# EL PRIMER VALOR DE CADA LISTA HACE REFERENCIA A LA SECCION DEL ARCHIVO
equipo = ["Nombre del ordenador:", "Nombre del ordenador:"] # nombre del ordenador
cpu =["Caja del sistema","Fabricante:","Número de serie:"] # LISTA CON ITEMS NECESESARIOS PARA CPU
procesador = ["Procesador(es) central","Nombre del procesador:", "Frecuencia del procesador original:","CPU ID:"]
ram1 =["Ram","Tamaño del módulo:","Fabricante del módulo:","Número de pieza del módulo:"]
ram2 =["Number Of Banks:","Tamaño del módulo:","Fabricante del módulo:","Número de pieza del módulo:"]
disco = ["Unidades de disco","Modelo de unidad:","Número de serie de la unidad:","Capacidad de la unidad:"]
cd = ["DVD","Modelo de unidad:","Número de serie:"]
monitor = ["Monitor","Nombre del monitor:","Nombre del monitor (del fabricante):","Número de serie:"]
red = ["Ethernet","Dirección MAC:"]
sistema = ["Nombre del ordenador:","Sistema operativo:"]

# CAMPOS DE EXCEL NECESARIOS PARA COLOCAR DATOS
# contiene el nombre de la celda donde debe ser colocados los datos, la lista debe tener el mismo tamaño
# que la lista que arroja el programa
camposequipo =["Y11"]
camposcpu =["M15","Z15"]
camposprocesador = ["G17","T17","Z17"]
camposram1 = ["G19","T19","Z19"]
camposram2 = ["G21","T21","Z21"]
camposdisco = ["G23","Z23","T23"]
camposcd = ["G24","Z24"]
camposmonitor = ["G26","Z26","Z27"]
camposred = ["Y36"]
campossistema = ["K42"]

ruta = "COMUNICACIONES2.CSV"
lista_final = buscar_palabra(ruta,equipo)
escribir_excel(lista_final,camposequipo) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,cpu)
escribir_excel(lista_final,camposcpu) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,procesador)
escribir_excel(lista_final,camposprocesador) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,ram1)
escribir_excel(lista_final,camposram1) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,ram2)
escribir_excel(lista_final,camposram2) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,disco)
escribir_excel(lista_final,camposdisco) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,cd)
escribir_excel(lista_final,camposcd) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,monitor)
escribir_excel(lista_final,camposmonitor) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,red)
escribir_excel(lista_final,camposred) # Impresion en archivo de excel
lista_final = buscar_palabra(ruta,sistema) 
escribir_excel(lista_final,campossistema) # Impresion en archivo de excel
print("PROCESO FINALIZADO")






# CODIGO PARA PONER EN HOJA DE EXCEL

# workbook = openpyxl.load_workbook('formato.xlsx')
# hoja = workbook['HOJA DE VIDA DE EQUIPOS']

# for i, elemento in enumerate(dicc_final):
#     #PONER EN LA HOJA DE EXCEL LOS DATOS CORRRESPONDIENTES
#     # hoja.cell(row=,column=).value=
#     print(elemento)
#     hoja.cell(row=73+i+1, column=1).value = elemento

# workbook.save('formato.xlsx')




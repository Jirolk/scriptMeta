#! python3
# import os
# from time import sleep
import xlrd
#carga del documento
archivo = 'data.xls'
#abrimos el documento  
wb = xlrd.open_workbook(archivo) 
hoja = wb.sheet_by_index(0) 
hoja = wb.sheet_by_name('Sheet1') 
nombres=[]

for i in range(0,hoja.nrows):
     #se carga un el vector todos los nombres de la columna Nombre
    nombres.append(hoja.cell_value(i,8))
    

#se ordena los registros no repetidos
nombres = list(set(nombres))
nombres.remove("NOME")
entrada=[]
salida=[]

for i in range(0,hoja.nrows) :
    for n in range(len(nombres)):
        if(hoja.cell_value(i,8) == nombres[n]):
            #nombre,hora,dia
            if(hoja.cell_value(i,4)=="Entrada"):
                entrada.append(list([nombres[n],"Entrada",hoja.cell_value(i,11),"Día",hoja.cell_value(i,14)]))
            elif(hoja.cell_value(i,4)=="Saída"):
                salida.append(list([nombres[n],"Salida",hoja.cell_value(i,11),"Día",hoja.cell_value(i,14)]))

mayor=0

for i in range(len(entrada)):
    for n in range(len(nombres[n])):
        # if(entrada[i][i]==nombres[n]):
        #    # mayor=int(max(entrada[i][2]))     
            # print("Entrada: ", entrada[i][i])           
        a=0

print("Salida: ", salida[1])


print("mayor: ", mayor)
print("registros de Entradas son: ",len(entrada))
print("Registros de Salida son: ",len(salida))
#for i in range(len(nombres)):

input()

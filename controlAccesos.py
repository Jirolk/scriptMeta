#! python3     
#abrimos el documento  
import xlrd
import xlwt
archivo="data.xls"
wb = xlrd.open_workbook(archivo) 
hoja = wb.sheet_by_index(0) 
hoja = wb.sheet_by_name('Sheet1') 
nombres=[]
horas=[]
entrada=[]
salida=[]
dia=[]

funcionarios=[]


for i in range(0,hoja.nrows):
     #se carga un el vector todos los nombres de la columna Nombre
    nombres.append(hoja.cell_value(i,8))
    dia.append(hoja.cell_value(i,14))

#se ordena los registros no repetidos
nombres = list(set(nombres))
nombres.remove("NOME")
dia= list(set(dia))
#eliminamos los datos que no nos interesan y los datos en blanco
dia.remove("DT_LEITORA - Dia")
dia.remove("")
#print("valor de dia: ", dia ,"cantidad de registro: ", len(dia))
#print("Cantidad de nombres: ", len(nombres))
#d es de dias
# h es de hoja de calculo
# n es de nombre

for d in range(len(dia)):
    for h in range(0,hoja.nrows):
        if(dia[d] == hoja.cell_value(h,14)):
            for n in range(len(nombres)):
                if(hoja.cell_value(h,8) == nombres[n]):
                #nombre,hora,dia
                    if(hoja.cell_value(h,4)=="Entrada"):
                       
                        entrada.append(list([nombres[n],"Entrada",hoja.cell_value(h,11),"Día",hoja.cell_value(h,14)]))
                        
                    elif(hoja.cell_value(h,4)=="Saída"):
                        salida.append(list([nombres[n],"Salida",hoja.cell_value(h,11),"Día",hoja.cell_value(h,14)]))

#entrada=list(set(max(entrada)))
#trabajamos creando un archivo excel para mandar los datos obtenidos
libro= xlwt.Workbook()
hoja_libro = libro.add_sheet("Entrada",cell_overwrite_ok=True)
hoja_salida = libro.add_sheet("Salida",cell_overwrite_ok=True)


hoja_libro.write(0,0,"Nombre")
hoja_libro.write(0,1,"Accion")
hoja_libro.write(0,2,"hora")
hoja_libro.write(0,3,"dia")


hoja_salida.write(0,0,"Nombre")
hoja_salida.write(0,1,"Accion")
hoja_salida.write(0,2,"hora")
hoja_salida.write(0,3,"dia")


for f in range(len(entrada)):
    #insertamos los que son de entrada
    hoja_libro.write(f+1,0,entrada[f][0])
    hoja_libro.write(f+1,1,entrada[f][1])
    hoja_libro.write(f+1,2,entrada[f][2])
    hoja_libro.write(f+1,3,entrada[f][4])
    
for f in range(len(salida)):
       #insertarmos lo que son de salida
    hoja_salida.write(f+1,0,salida[f][0])
    hoja_salida.write(f+1,1,salida[f][1])
    hoja_salida.write(f+1,2,salida[f][2])
    hoja_salida.write(f+1,3,salida[f][4])


libro.save("./HorariodeAccesos.xls")



print("registros de Entradas son: ",len(entrada))
print("Registros de Salida son: ",len(salida))

input()

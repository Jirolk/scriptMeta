#! python3     
#abrimos el documento  
import xlrd
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
print("Cantidad de nombres: ", len(nombres))
#d es de dias
# h es de hoja de calculo
# n es de nombre

for d in range(len(dia)):
    for h in range(0,hoja.nrows):
        if(dia[d] == hoja.cell_value(h,14)):
            for n in range(len(nombres)):
                if(hoja.cell_value(i,8) == nombres[n]):
                #nombre,hora,dia
                    if(hoja.cell_value(h,4)=="Entrada"):
                       
                        entrada.append(list([nombres[n],"Entrada",hoja.cell_value(h,11),"Día",hoja.cell_value(h,14)]))
                        
                    elif(hoja.cell_value(h,4)=="Saída"):
                        salida.append(list([nombres[n],"Salida",hoja.cell_value(h,11),"Día",hoja.cell_value(h,14)]))

               
         
        





mayor=0

for i in range(len(entrada)):
    for n in range(len(nombres[n])):
        # if(entrada[i][i]==nombres[n]):
        #    # mayor=int(max(entrada[i][2]))     
            # print("Entrada: ", entrada[i][i])           
        a=0

print(entrada)


print("mayor: ", mayor)
print("registros de Entradas son: ",len(entrada))
print("Registros de Salida son: ",len(salida))

input()

#! python3     
#abrimos el documento  
import xlrd
import xlwt

estiloHora = xlwt.XFStyle()
estiloHora.num_format_str = 'HH:MM:SS'



archivo="data.xls"
wb = xlrd.open_workbook(archivo) 
hoja = wb.sheet_by_index(0) 
hoja = wb.sheet_by_name('Sheet1') 
nombres=[]
horas=[]
entrada=[]
entradaAux=[]

salida=[]
salidaAux=[]

dia=[]

funcionarios={}

tupla=[]
for i in range(0,hoja.nrows):
     #se carga un el vector todos los nombres de la columna Nombre
    nombres.append(hoja.cell_value(i,8))
    dia.append(hoja.cell_value(i,14))
    print(hoja.cell_value(i,8), hoja.cell_value(i,5), hoja.cell_value(i,14))
    
    funcionarios= {
        hoja.cell_value(i,8),
        hoja.cell_value(i,5),
        hoja.cell_value(i,14)
    }
    
    # funcionarios.update(funcionarios)    

print(funcionarios)
print(len(funcionarios))
input("oprima una tecla para continuar")
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
                        entrada.append(list([nombres[n],hoja.cell_value(h,5),"Entrada",hoja.cell_value(h,11),hoja.cell_value(h,14)]))
                    elif(hoja.cell_value(h,4)=="Saída"):
                        salida.append(list([nombres[n],hoja.cell_value(h,5),"Salida",hoja.cell_value(h,11),hoja.cell_value(h,14)]))

maxiEn=[0]
maxiEntra=[0]

for e in range(len(entrada)):
    for d in range(len(dia)):
        for n in range(len(nombres)):
            if(entrada[e][4]==23):
                if(entrada[e][3] > max(maxiEn) ):
                    maxiEn.append(entrada[e][3])
                    maxiEntra.append(entrada[e])
           
           
           
           
maxiEn.remove(0)
maxiEntra.remove(0)                    
                 

print("Prueba: ", maxiEn)                 
print("el mayor: ", max(entrada))
print("Entrada vista: ", len(maxiEntra))
print("vista entrada: ", maxiEntra)
#trabajamos creando un archivo excel para mandar los datos obtenidos
libro= xlwt.Workbook()
hoja_libro = libro.add_sheet("Entrada",cell_overwrite_ok=True)
hoja_salida = libro.add_sheet("Salida",cell_overwrite_ok=True)

#se escribe en la hoja de entrada
hoja_libro.write(0,0,"Nombre")
hoja_libro.write(0,1,"Empresa")
hoja_libro.write(0,2,"Accion")
hoja_libro.write(0,3,"hora")
hoja_libro.write(0,4,"dia")

#se escribe en la hoja de salida
hoja_salida.write(0,0,"Nombre")
hoja_salida.write(0,1,"Empresa")
hoja_salida.write(0,2,"Accion")
hoja_salida.write(0,3,"hora")
hoja_salida.write(0,4,"dia")


for f in range(len(entrada)):
    #insertamos los que son de entrada
    hoja_libro.write(f+1,0,entrada[f][0])
    hoja_libro.write(f+1,1,entrada[f][1])
    hoja_libro.write(f+1,2,entrada[f][2])
    hoja_libro.write(f+1,3,entrada[f][3], estiloHora)
    hoja_libro.write(f+1,4,entrada[f][4])
    
for f in range(len(salida)):
       #insertarmos lo que son de salida
    hoja_salida.write(f+1,0,salida[f][0])
    hoja_salida.write(f+1,1,salida[f][1])
    hoja_salida.write(f+1,2,salida[f][2])
    hoja_salida.write(f+1,3,salida[f][3], estiloHora)
    hoja_salida.write(f+1,4,salida[f][4])


libro.save("./HorariodeAccesos.xls")


print("\n registros de Entradas son: ",len(entrada))
print("Registros de Salida son: ",len(salida))

input()

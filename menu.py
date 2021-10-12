
import tkinter as tk
from tkinter import Message, messagebox
import os
from tkinter.constants import BOTH 
import xlrd

empreNombre=[]

ventana= tk.Tk()

try:
        
    #carga del documento
    archivo = 'download.xls'
    #abrimos el documento  
    wb = xlrd.open_workbook(archivo) 
    hoja = wb.sheet_by_index(0) 
    hoja = wb.sheet_by_name('Folha1') 
    for i in range(0,hoja.nrows):
        #se carga un el vector todos los nombres de la columna Empresa
        empreNombre.append(hoja.cell_value(i,5))
    
    empreNombre = list(set(empreNombre))
    empreNombre.remove("Empresa")
    empreNombre.sort()
    
except OSError as error:
    if messagebox.showerror(message=error,title="Algo salío Mal"):
        ventana.destroy()
        
def salu():
    messagebox.showwarning(message="En desarrollo!!",title="prueba" )
    
def salir():
    if messagebox.askokcancel("Salir","Estas Seguro que quieres Cerrar?"):
        ventana.destroy()    
def verEmpresas():
  
    
    os.system("cls")
    #Se imprime en pantalla la cantidad de Empresas
    print("\n\tSe posee actualmente: ", len(empreNombre)," empresas\n")
                
    for i in range(len(empreNombre)):
        print("\t♦",i+1,") ",empreNombre[i],".")        

def sexoEmpresa(empresa):
    '''Función para Calcular la cantidad de
        mujeres y hombres'''
    empresaM=0;
    empresaF=0;
    for i in range(0,hoja.nrows):

        if (hoja.cell_value(i,5)== empresa
            and hoja.cell_value(i,2)== 'Feminino'):
            empresaF+=1;
        elif (hoja.cell_value(i,5)== empresa
            and hoja.cell_value(i,2)== 'Masculino'):
            empresaM+=1;

    print("\n============================================")
    print(" La Empresa ", empresa," \n\t\ttiene Mujeres: ", empresaF)
    print("\ttiene Hombres: ", empresaM)
    print("\tTotal de Colaboradores: ", empresaM+empresaF)
    print("============================================\n")
    # input("Oprima enter para continuar: ")
    #os.system("pause")

    
def listarEmpresas():
    os.system("cls")
    for i in range(len(empreNombre)):
        
        if(empreNombre[i] != "Empresa"):
             sexoEmpresa(empreNombre[i])


def manodeObra():
    sexoF=0;
    sexoM=0;
    para=0;
    bra=0;
    sui=0;
    lati=0;
    concep=0;
    sanpedro=0;
    cordillera=0;
    guaira=0;
    caazapa=0;
    caaguazu=0;
    itapua=0;
    misiones=0;
    paraguari=0;
    altoParana=0;
    central=0;
    neembucu=0;
    amambay=0;
    canindeyu=0;
    pteHayes=0;
    boqueron=0;
    distritoCapi=0;
    altoPy=0;
    otroDepart=0;
    for i in range(0,hoja.nrows):
        #Se calcula la cantidad por sexo y nacionalidad
        if (hoja.cell_value(i,2)== 'Feminino'):
            sexoF+=1;        
        elif(hoja.cell_value(i,2) == 'Masculino'):
            sexoM+=1
        if (hoja.cell_value(i,3) == 'PARAGUAYA'
            or hoja.cell_value(i,3)== 'Paraguaya'
            or hoja.cell_value(i,3)== 'Paraguaio'
            or hoja.cell_value(i,3)== 'PARAGUAYO'
            or hoja.cell_value(i,3)== 'PARAGAYA'
            or hoja.cell_value(i,3)== 'PARAGUAI'):
            para+=1
        elif (hoja.cell_value(i,3) == 'BRASILEIRA'
            or hoja.cell_value(i,3)== 'BRASILEIRO'
            or hoja.cell_value(i,3)== 'BRASILERA'
            or hoja.cell_value(i,3)== 'BRASILERO'):
            bra+=1
        elif (hoja.cell_value(i,3) == 'SUIZO'):
            sui+=1
        elif(hoja.cell_value(i,3)!= 'Nacionalidade'
            and hoja.cell_value(i,3)!= 'SUIZO'):
            lati+=1

    for i in range(0,hoja.nrows) :
        #cuenta por ciudades.
        if(hoja.cell_value(i,4) == 'CONCEPCIÓN'):
            concep+=1;
        elif(hoja.cell_value(i,4) == 'SAN PEDRO'):
            sanpedro+=1;
        elif(hoja.cell_value(i,4) == 'CORDILLERA'):
            cordillera+=1;  
        elif(hoja.cell_value(i,4)== 'GUAIRÁ'):
            guaira+=1;
        elif(hoja.cell_value(i,4)== 'CAAZAPÁ'):
            caazapa+=1;
        elif(hoja.cell_value(i,4)== 'CAAGUAZÚ'):
            caaguazu+=1;        
        elif(hoja.cell_value(i,4)== 'ITAPÚA'):
            itapua+=1;
        elif(hoja.cell_value(i,4)== 'MISIONES'):
            misiones+=1;
        elif(hoja.cell_value(i,4)== 'PARAGUARÍ'):
            paraguari+=1;
        elif(hoja.cell_value(i,4)== 'ALTO PARANÁ'):
            altoParana+=1;
        elif(hoja.cell_value(i,4)== 'CENTRAL'):
            central+=1;
        elif(hoja.cell_value(i,4)== 'ÑEEMBUCÚ'):
            neembucu+=1;
        elif(hoja.cell_value(i,4)== 'AMAMBAY'):
            amambay+=1;
        elif(hoja.cell_value(i,4)== 'CANINDEYÚ'):
            canindeyu+=1;
        elif(hoja.cell_value(i,4)== 'PRESIDENTE HAYES'):
            pteHayes+=1;
        elif(hoja.cell_value(i,4)== 'BOQUERÓN'):
            boqueron+=1;
        elif(hoja.cell_value(i,4)== 'DISTRITO CAPITAL'):
            distritoCapi+=1;
        elif(hoja.cell_value(i,4)== 'ALTO PARAGUAY'):
            altoPy+=1;
        elif (hoja.cell_value(i,4) != 'NaturalidadeUF'):
            otroDepart+=1
    os.system("cls")
    print("\n\n \tCantidad de Mujeres: ", sexoF, " \n\tCantidad de Hombres: ", sexoM)
    print("\n\tCantidad de Paraguayos: ", para)
    print("\tCantidad de Brasileros: ", bra)
    print("\tCantidad de Suizos: ", sui)
    print("\tCantidad de Latinos: ", lati)
    print("\n\tCantidad de Colaboradores: ",hoja.nrows-1)
    print("\n\t De Concepción son: ", concep)
    print("\t De San Pedro son: ", sanpedro)
    print("\t De Cordillera son: ", cordillera )
    print("\t De Guaira son: ", guaira)
    print("\t De Caazapa son: ", caazapa)
    print("\t De Caaguazú son: ", caaguazu)
    print("\t De Itapua son: ",itapua)
    print("\t De Misiones son: ", misiones)
    print("\t De Paraguari son: ", paraguari)
    print("\t De Alto Parana son: ", altoParana)
    print("\t De Central son: ", central)
    print("\t De Ñe'embucu son: ", neembucu)
    print("\t De Amambay son: ", amambay)
    print("\t De Canindeyu son: ", canindeyu)
    print("\t De Pte. Hayes son: ", pteHayes)
    print("\t De Boquerón son: ", boqueron)
    print("\t De Distrito Capital son: ", distritoCapi)
    print("\t De Alto Paraguay son: ", altoPy)
    print("\t De otros Departamentos son: ", otroDepart)

def porDepartamento():
    con=0; horqueta=0; belen=0; loreto=0; sanLazaro=0; jose=0;
    concepM=0; horquetaM=0; belenM=0; loretoM=0; sanLazaroM=0; joseM=0;
    concepF=0
    for i in range(0,hoja.nrows) :
            #cuenta por ciudades.
        if(hoja.cell_value(i,7) == 'CONCEPCIÓN'):
            con+=1
            
        elif(hoja.cell_value(i,7) == 'HORQUETA'):
            horqueta+=1
        elif(hoja.cell_value(i,7) == 'BELÉN'):
            belen+=1;  
        elif(hoja.cell_value(i,7)== 'LORETO'):
            loreto+=1
        elif(hoja.cell_value(i,7)== 'SAN LAZARO'):
            sanLazaro+=1
        elif(hoja.cell_value(i,7)== 'SARGENTO JOSÉ FÉLIX LÓPEZ'):
            jose+=1        
    
    
            
def limpiarPantalla():
    os.system("cls")
    print ("\n Datos a tenes encuenta para extraer la planilla\n    NOME\n    Sexo\n    DataNascimento\n    Nacionalidade\n    NaturalidadeUF\n    Empresa\n    Naturalidade \n    Bairro ")
    
ventana.title("Reportes Meta-Xs")
ventana.geometry("300x200")

btn1 = tk.Button(ventana, text="Ver Empresas", command=verEmpresas).pack(expand=True, fill=tk.BOTH)

btn2 = tk.Button(ventana, text="Listar Empresas", command=listarEmpresas).pack(expand=True, fill=tk.BOTH)

btn3 = tk.Button(ventana, text="Ver Mano de Obra", command=manodeObra).pack(expand=True, fill=tk.BOTH)
btn4 = tk.Button(ventana, text="Limpiar Pantalla", command=limpiarPantalla).pack(expand=True, fill=tk.BOTH)
btn4 = tk.Button(ventana, text="Salir", command=salir).pack(expand=True, fill=tk.BOTH)
os.system("mode con: cols=80")
ventana.mainloop()

#! python3
# import os
# from time import sleep
import xlrd

#os.system("cls")
#os.system("echo                        Resumen de los datos Extraidos del sistema")

#carga del documento
archivo = 'download.xls'
#abrimos el documento  
wb = xlrd.open_workbook(archivo) 
hoja = wb.sheet_by_index(0) 
hoja = wb.sheet_by_name('Folha1') 

#varaibles locales 
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

empreNombre=[]

#funciones 

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
    print("\tLa Empresa ", empresa," \n\t\ttiene Mujeres: ", empresaF)
    print("\t\ttiene Hombres: ", empresaM)
    print("\tTotal de Colaboradores: ", empresaM+empresaF)
    print("============================================\n")
    input()
    #os.system("pause")



def alerta():
    ''''Función de alerta para ajustar las columnas del documentos que se extrae
    del sistema'''
    #os.system("cls")
    #os.system("echo                        Resumen de los datos Extraidos del sistema")
    #print("\n\t\t\t Los datos a extraerse tienen que ser")
    #print("\n\t 'Nome'  \t 'DataNascimento' \t 'Sexo' \t 'Nacionalidade' \t 'NaturalidadeUF' \t 'Empresa'\n\n\n\n ")





for i in range(0,hoja.nrows):
    if (i==1):
        #para tener controlado la iteración
        break
    elif(hoja.cell_value(i,0) != 'Nome' or hoja.cell_value(i,1) != 'DataNascimento' 
        or hoja.cell_value(i,2) != 'Sexo'
        or hoja.cell_value(i,3) != 'Nacionalidade' or hoja.cell_value(i,4) != 'NaturalidadeUF' 
        or hoja.cell_value(i,5) != 'Empresa'):
        alerta()
        print ("\t\t\t\t\tFijarse que campos faltan")
        print ("\t '"+hoja.cell_value(i,0)+"'"+" \t '"+hoja.cell_value(i,1)+"'"+" \t '"+hoja.cell_value(i,2)+"'"" \t '"+hoja.cell_value(i,3)+"'"
               " \t '"+hoja.cell_value(i,4)+"'"+" \t '"+hoja.cell_value(i,5)+"'\n\n\n")
     #   os.system("pause")
        break
    
    else:      
        for i in range(0,hoja.nrows):

                    
            #se carga un el vector todos los nombres de la columna Empresa
            empreNombre.append(hoja.cell_value(i,5))

            #Se calcula la cantidad por sexo y nacionalidad
            if (hoja.cell_value(i,2)== 'Feminino'):
                sexoF+=1;        
            elif(hoja.cell_value(i,2) == 'Masculino'):
                sexoM+=1;
            if (hoja.cell_value(i,3) == 'PARAGUAYA'
                or hoja.cell_value(i,3)== 'Paraguaya'
                or hoja.cell_value(i,3)== 'Paraguaio'
                or hoja.cell_value(i,3)== 'PARAGUAYO'
                or hoja.cell_value(i,3)== 'PARAGAYA'
                or hoja.cell_value(i,3)== 'PARAGUAI'):
                para+=1;
            elif (hoja.cell_value(i,3) == 'BRASILEIRA'
                or hoja.cell_value(i,3)== 'BRASILEIRO'
                or hoja.cell_value(i,3)== 'BRASILERA'
                or hoja.cell_value(i,3)== 'BRASILERO'):
                bra+=1;
            elif (hoja.cell_value(i,3) == 'SUIZO'):
                sui+=1;
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

        #Se extrae solo los datos que no se repiten
        empreNombre = list(set(empreNombre))
        #Se recorre el vector hasta la cantidad de empresas
        for i in range(len(empreNombre)):
            if(empreNombre[i] != "Empresa"):
                sexoEmpresa(empreNombre[i])
        
        print("\n=====================================================================================")
        print("\t\tConclusión ")
        print("\tCantidad de Mujeres: ", sexoF, " \n\tCantidad de Hombres: ", sexoM)
        print("\tCantidad de Paraguayos: ", para)
        print("\tCantidad de Brasileros: ", bra)
        print("\tCantidad de Suizos: ", sui)
        print("\tCantidad de Latinos: ", lati)
        print("\tCantidad de Colaboradores: ",hoja.nrows-1)
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

        print("=======================================================================================\n")
input("ndeeeeeeeeeee")
        #os.system("pause")
        

#Python
import tkinter as tk
from tkinter import messagebox
import time
import xlrd
wb=""; hoja = object; empreNombre=[]

def comprobarExcel(empreNombre,hoja):
    
    try:
        wb = xlrd.open_workbook('download.xls')
        hoja = wb.sheet_by_index(0) 
        hoja = wb.sheet_by_name('Folha1') 
        for i in range(0,hoja.nrows):
            #se carga un el vector todos los nombres de la columna Empresa
            empreNombre.append(hoja.cell_value(i,5))
       
        
    except OSError as error:
        
        if messagebox.showerror("Hubo un Problema ,\n Verifique el nombre del archivo",error):
            if messagebox.showinfo("Hasta luego","Verifique el nombre del Archivo, nos vemos!"):
                ventana.destroy()
       
    
    
def parpadear():
    ventana.iconify()
    time.sleep(3)
    ventana.deiconify()

def imprimir():
    print("Acabas de presionar el boton de imprimir")
def salir():
    if messagebox.askokcancel("Salir","Estas Seguro en cerrar la aplicación?"):
        ventana.destroy()

def verEmpresas(empreNombre, hoja):
    empreNombre = list(set(empreNombre))
    empreNombre.remove("Empresa")
    empreNombre.sort()
    
    #Se imprime en pantalla la cantidad de Empresas
    print("\n\tSe posee actualmente: ", len(empreNombre)," empresas \n son: \n")
                
    for i in range(len(empreNombre)):
        print("\t\t ♦",i+1,") ",empreNombre[i],".")
    
    print("\tCantidad de Colaboradores: ",hoja.nrows-1,"\n")



ventana= tk.Tk()
ventana.title("Reportes del Sistema Meta-X")
ventana.geometry("400x200")
comprobarExcel(empreNombre,hoja)
barMenu= tk.Menu(ventana)
barMenu.add_command(label="Ver Empresas", command=verEmpresas(empreNombre,hoja))
barMenu.add_command(label="Listar Empresas")
barMenu.add_command(label="Salir", command=salir)    

    

    

ventana.config(menu=barMenu)
ventana.mainloop()

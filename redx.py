import pandas as pd
import sys
import os
import re
import pickle
import tkinter as tk
from tkinter import filedialog


def main():

    if (len(sys.argv) > 1):
        while True:
            inp = input('1 - rutas\n2 - config cuentas\n3 - config cols excel\n0 - SALIR\n')
            try:
                if int(inp) == 1:
                    create_gui()
                elif (int(inp) == 2):
                    config_cuentas()
                elif (int(inp) == 3):
                    create_gui_cols()
                    order_cols()
                elif (int(inp) == 0):
                    break
                    
            except ValueError:
                print('error del usuario')
                pass 
                        

    try:
        mes = int(input("Mes: "))
        año = int(input("Año (ultimas dos cifras unicamente): "))

    except ValueError:
        print("Preciso el cada uno en dos cifras (##).")
    catch_bad_human(año,mes)    

    with open('route.pickle', 'rb') as file:
        # Deserialize the object
        dir_exl = pickle.load(file)
        
    #dir_final = input("Direccon de la carpeta de la empresa: C:")
    dir_final = r"C:\Users\eliva23\Desktop\ProjectsVS"

    final=create_list(dir_exl,mes,año)
    f = open (f"{dir_final}/IM{año}{mes:02d}.txt", "w")
    f.writelines(final)
    f.close()
    
    print(f"Finalizado, el archivo IM{año}{mes:02d} fue creado")


def browse_file():
    # Add your code here to process the file route or perform any desired action
    filename = filedialog.askopenfile(initialdir="/", title="Select File")
    ruta = filename.name
    print(ruta)
    with open('route.pickle', 'wb') as file:
        # Serialize the object and write it to the file
        pickle.dump(ruta, file)
    
def create_gui():
    # Create the main window
    window = tk.Tk()
    window.title("File Route Search")

    # Create a frame
    frame = tk.Frame(window, padx=20, pady=20)
    frame.pack()

    # Create a label
    label = tk.Label(frame, text="Enter the route of a file:")
    label.pack()

    
    #browsew nutton
    browse_button = tk.Button(frame, text="Browse", command=browse_file)
    browse_button.pack()

    # Start the Tkinter event loop
    window.mainloop()

def config_cuentas():
    ctas_dict = {}
    
    caja = input('Cuenta caja m/n: ')
    ctas_dict['cta_caja'] = caja
    
    deudores = input('Cuenta deudores varios: ')
    ctas_dict['cta_deudores'] = deudores
    
    ivabas = input('Cuenta ventas tasa basica: ')
    ctas_dict['cta_ivabas'] = ivabas
    
    ivamin = input('Cuenta ventas tasa minima: ')
    ctas_dict['cta_ivamin'] = ivamin

    ivaexe = input('Cuenta ventas tasa excenta: ')
    ctas_dict['cta_exce'] = ivaexe 
    
    codivab = input('Codigo iva basico: ')
    ctas_dict['cod_ivabas'] = codivab
    
    codivam = input('Codigo iva minimo: ')
    ctas_dict['cod_ivamin'] = codivam
    
    with open('.\ctas.pickle', 'wb') as file:
        # Serialize the object and write it to the file
        pickle.dump(ctas_dict, file)
        
options = []
lst = []

def get_selected_options():
    selected_options = [option.get() for option in options]
    ##print("Selected options:")
    
    fin_selec = []
    for o in range(len(selected_options)):
        if (selected_options[o]):
            fin_selec.append(lst[o])
            ##print(lst[o])
    with open('cols.pickle', 'wb') as file:
        # Serialize the object and write it to the file
        pickle.dump(fin_selec, file)
                    

    
def create_gui_cols():
    root = tk.Tk()
    root.title("Multiple Choice")

    frame = tk.Frame(root)
    frame.pack(padx=20, pady=20)

    # Create variables to store the selected options
    with open('route.pickle', 'rb') as file:
        # Deserialize the object
        arch_ventas = pickle.load(file)
    data = pd.read_excel(f'{arch_ventas}')

    for x in range(len(data.columns)):
        lst.append(f'{data.columns[x]}')
    
    # Create Checkbuttons
    for x in range(len(lst)):     
        option1 = tk.BooleanVar()
        option1_checkbutton = tk.Checkbutton(frame, text=f"{lst[x]}", variable=option1, onvalue=True, offvalue=False)
        options.append(option1)
        if (x < 15):
            option1_checkbutton.grid(row=x, column=1, sticky="w")
        elif (x < 30):
            option1_checkbutton.grid(row=(x-15), column=2, sticky="w")
        elif (x < 45):
            option1_checkbutton.grid(row=(x-30), column=3, sticky="w")
        else:
            break               


    # Create a button to get the selected options
    button = tk.Button(frame, text="Get Selected Options", command=get_selected_options)
    button.grid(row=0, column=0, pady=10)

    root.mainloop()

def order_cols():
    order_cols_list = [0,1,2,3,4,5,6,7,8,9]
    with open('cols.pickle', 'rb') as file:
        # Deserialize the object
        cols_list = pickle.load(file)
        for x in range(len(cols_list)):
            print(f'{x}: {cols_list[x]}', end = ' | ')
    
    print()
    print('Ordena las columnas con el orden ya predeterminado, presiona n si la columnaa no existe.')
    tipo = input('Columna que contiene el tipo de comprobante: ')
    num = input('Numero de comprobante: ')
    fecha_e = input('Fecha de emision: ')
    forma = input('Forma de pago: ')
    mon = input('Moneda: ')
    iva_min = input('IVA minimo: ')    
    iva_bas = input('IVA basico: ')
    neto_min = input('Neto minimo: ')
    neto_bas = input('Neto basico: ')
    tot = input('Total: ')
    
    if (tipo != 'n'):
        order_cols_list [0] = cols_list[int(tipo)]
    else:
        order_cols_list [0] = -1
    if (num != 'n'):    
        order_cols_list [1] = cols_list[int(num)]
    else:
        order_cols_list [1] = -1  
    if (fecha_e != 'n'):    
        order_cols_list [2] = cols_list[int(fecha_e)]
    else:
        order_cols_list [2] = -1   
    if (forma != 'n'):    
        order_cols_list [3] = cols_list[int(forma)]
    else:
        order_cols_list [3] = -1    
    if (mon != 'n'):    
        order_cols_list [4] = cols_list[int(mon)]
    else:
        order_cols_list [4] = -1   
    if (iva_min != 'n'):    
        order_cols_list [5] = cols_list[int(iva_min)]
    else:
        order_cols_list [5] = -1   
    if (iva_bas != 'n'):    
        order_cols_list [6] = cols_list[int(iva_bas)]
    else:
        order_cols_list [6] = -1  
    if (neto_min != 'n'):    
        order_cols_list [7] = cols_list[int(neto_min)]
    else:
        order_cols_list [7] = -1  
    if (neto_bas != 'n'):    
        order_cols_list [8] = cols_list[int(neto_bas)]
    else:
        order_cols_list [8] = -1  
    if (tot != 'n'):    
        order_cols_list [9] = cols_list[int(tot)]
    else:
        order_cols_list [9] = -1    

    with open('ocols.pickle', 'wb') as file:
        # Serialize the object and write it to the file
        pickle.dump(order_cols_list, file)
        

def catch_bad_human(año,mes):

    if len(str(año)) !=2:
        sys.exit("Error: el año debe tener solo 2 cifras")  
    if 1<=mes<=12:
        pass
    else:
        sys.exit("Error: el mes no existe") 


def create_list(dir_exl,mes,año):

    archivo_excel = pd.read_excel(f"{dir_exl}")


    #KEY X COLUMNA
    with open('ocols.pickle', 'rb') as file:
        # Deserialize the object
        ocols_list = pickle.load(file)
        print(f'col list: {ocols_list}')
        
  
    if (ocols_list[0] != -1):    
        tipo = archivo_excel[f'{ocols_list[0]}'].values
    else:
        tipo = []
        for x in range(500):
            tipo.append(0)  
    if (ocols_list[1] != -1):             
        num = archivo_excel[f'{ocols_list[1]}'].values
    else:
        num = []
        for x in range(500):
            num.append(0)        
    if (ocols_list[2] != -1):  
        fecha_e = archivo_excel[f'{ocols_list[2]}'].values
    else:
        fecha_e = []
        for x in range(500):
            fecha_e.append(0)         
    if (ocols_list[3] != -1):       
        forma = archivo_excel[f'{ocols_list[3]}'].values
    else:
        forma = []
        for x in range(500):
            forma.append(0)        
    if (ocols_list[4] != -1):    
        mon = archivo_excel[f'{ocols_list[4]}'].values 
    else:
        mon = []
        for x in range(500):
            mon.append(0)        
    if (ocols_list[5] != -1):   
        iva_min = archivo_excel[f'{ocols_list[5]}'].values          
    else:
        iva_min = []
        for x in range(500):
            iva_min.append(0)         
    if (ocols_list[6] != -1):      
        iva_bas = archivo_excel[f'{ocols_list[6]}'].values
    else:
        iva_bas = []
        for x in range(500):
            iva_bas.append(0)         
    if (ocols_list[7] != -1):         
        neto_min = archivo_excel[f'{ocols_list[7]}'].values
    else:
        neto_min = []
        for x in range(500):
            neto_min.append(0)       
    if (ocols_list[9] != -1):         
        neto_bas = archivo_excel[f'{ocols_list[8]}'].values
    else:
        neto_bas = []
        for x in range(500):
            neto_bas.append(0)        
    if (ocols_list[9] != -1):         
        tot = archivo_excel[f'{ocols_list[9]}'].values
    else:
        tot = []
        for x in range(500):
            tot.append(0)  


    
    
    f_list= ['Dia, Debe, Haber, Concepto, RUT, Moneda,  Total, CodigoIVA, IVA, Cotizacion, Libro']
    
    with open('ctas.pickle', 'rb') as file:
        # Deserialize the object
        ctas_dict = pickle.load(file)
        ##print(ctas_dict)

    #CUENTAS
    cta_caja = ctas_dict['cta_caja']
    cta_deudores = ctas_dict['cta_deudores']
    cta_ivabas = ctas_dict['cta_ivabas']
    cta_ivamin = ctas_dict['cta_ivamin']
    cta_exent = '0'

    #CODIGOS DE IVA
    cod_ivabas = ctas_dict['cod_ivabas']
    cod_ivamin = ctas_dict['cod_ivamin'] 
    
    ##print(iva_bas)
    for x in range(len(tot)):
        f_1=str(fecha_e[x])
        #print(f_1)
        day = re.search(r"^(\d+).?(\d+).?(\d+).+",f_1)
        #print(day)
        if not day:
            f_2=f"{1:02}"
        else:
            f_2= day.group(3)
        ##print(int(iva_bas[x]))
        ##print(type(iva_bas[x]))
        ##print('mon:', mon[x],'for', forma[x])

        
        if str(mon[x]) == "$" and "Contado" in str(forma[x]):
            ##print('YES')
            f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas},"{tipo[x]} A {num[x]}",,0,{tot[x]},{cod_ivabas},{iva_bas[x]},0,I')
        elif mon[x] == "$" and "Credito" in forma[x]:
            f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas},"{tipo[x]} A {num[x]}",,0,{tot[x]},{cod_ivabas},{iva_bas[x]},0,V')
        elif mon[x] == "U$S" and "Contado" in forma[x]:
            f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas},"{tipo[x]} A {num[x]}",,1,{tot[x]},{cod_ivabas},{iva_bas[x]},0,I')
        elif mon[x] == "U$S" and "Credito" in forma[x]:
            f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas},"{tipo[x]} A {num[x]}",,1,{tot[x]},{cod_ivabas},{iva_bas[x]},0,V')
        elif mon[x] == "$" and "Crédito" in forma[x]:
            f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas},"{tipo[x]} A {num[x]}",,0,{tot[x]},{cod_ivabas},{iva_bas[x]},0,V')
        elif mon[x] == "U$S" and "Crédito" in forma[x]:
            f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas},"{tipo[x]} A {num[x]}",,0,{tot[x]},{cod_ivabas},{iva_bas[x]},0,V')
        
        if iva_bas[x] != 0 and iva_min[x] != 0:  #Chequea si una misma venta esta grabada con IVA basico y minimo a la vez.
            if mon[x] == "UYU" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas}," Venta {tipo[x]} A {num[x]}",0,0,{neto_bas[x]+iva_bas[x]},{cod_ivabas},{iva_bas[x]},0,I')
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivamin}," Venta {tipo[x]} A {num[x]}",0,0,{neto_min[x]+iva_min[x]},{cod_ivamin},{iva_min[x]},0,I')
            elif mon[x] == "UYU" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas}," Venta {tipo[x]} A {num[x]}",0,0,{neto_bas[x]+iva_bas[x]},{cod_ivabas},{iva_bas[x]},0,V')
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivamin}," Venta {tipo[x]} A {num[x]}",0,0,{neto_min[x]+iva_min[x]},{cod_ivamin},{iva_min[x]},0,V')
            elif mon[x] == "USD" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas}," Venta {tipo[x]} A {num[x]}",0,1,{neto_bas[x]+iva_bas[x]},{cod_ivabas},{iva_bas[x]},0,I')
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivamin}," Venta {tipo[x]} A {num[x]}",0,1,{neto_min[x]+iva_min[x]},{cod_ivamin},{iva_min[x]},0,I')
            elif mon[x] == "USD" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas}," Venta {tipo[x]} A {num[x]}",0,1,{neto_bas[x]+iva_bas[x]},{cod_ivabas},{iva_bas[x]},0,V')
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivamin}," Venta {tipo[x]} A {num[x]}",0,1,{neto_min[x]+iva_min[x]},{cod_ivamin},{iva_min[x]},0,V')
        
        elif iva_bas[x] != 0:   #Chequea las compras que solamente esta grabadas con IVA basico.
            if mon[x] == "UYU" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas}," Venta {tipo[x]} A {num[x]}",0,0,{tot[x]},{cod_ivabas},{iva_bas[x]},0,I')
            elif mon[x] == "UYU" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas},"Venta {tipo[x]} A {num[x]}",0,0,{tot[x]},{cod_ivabas},{iva_bas[x]},0,V')
            elif mon[x] == "USD" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas},"Venta {tipo[x]} A {num[x]}",0,1,{tot[x]},{cod_ivabas},{iva_bas[x]},0,I')
            elif mon[x] == "USD" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivabas},"Venta {tipo[x]} A {num[x]}",0,1,{tot[x]},{cod_ivabas},{iva_bas[x]},0,V')

        elif iva_min[x] != 0:   #Chequea las compras que solamente esta grabadas con IVA minimo.
            if mon[x] == "UYU" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivamin}," Venta {tipo[x]} A {num[x]}",0,0,{tot[x]},{cod_ivamin},{iva_min[x]},0,I')
            elif mon[x] == "UYU" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivamin},"Venta {tipo[x]} A {num[x]}",0,0,{tot[x]},{cod_ivamin},{iva_min[x]},0,V')
            elif mon[x] == "USD" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_ivamin},"Venta {tipo[x]} A {num[x]}",0,1,{tot[x]},{cod_ivamin},{iva_min[x]},0,I')
            elif mon[x] == "USD" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_ivamin},"Venta {tipo[x]} A {num[x]}",0,1,{tot[x]},{cod_ivamin},{iva_min[x]},0,V')

        else:   #Chequea las compras no grabadas.
            if mon[x] == "UYU" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_exent}," Venta {tipo[x]} A {num[x]}",0,0,{tot[x]},0,0,0,I')
            elif mon[x] == "UYU" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_exent},"Venta {tipo[x]} A {num[x]}",0,0,{tot[x]},0,0,0,V')
            elif mon[x] == "USD" and forma[x] == "Contado":
                f_list.append(f'\n{f_2},{cta_caja},{cta_exent},"Venta {tipo[x]} A {num[x]}",0,1,{tot[x]},0,0,0,I')
            elif mon[x] == "USD" and forma[x] == "Crédito":
                f_list.append(f'\n{f_2},{cta_deudores},{cta_exent},"Venta {tipo[x]} A {num[x]}",0,1,{tot[x]},0,0,0,V')
        
        
        #if mon[x] == "$" and forma[x] == "Devolución Venta Contado ":
        #    f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas},"{tipo[x]} A {num[x]}",,0,{tot[x]},{cod_ivabas},{iva_bas[x]},0,I')
        #elif mon[x] == "U$S" and forma[x] == "Devolución Venta Contado ":
        #    f_list.append(f'\n{f_2},{cta_caja},{cta_ivabas},"{tipo[x]} A {num[x]}",,1,{tot[x]},{cod_ivabas},{iva_bas[x]},0,I')       


    return f_list

if __name__ == "__main__":
    main()

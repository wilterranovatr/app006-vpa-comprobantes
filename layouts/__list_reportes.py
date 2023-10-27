from tkinter import *
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter.messagebox import showinfo
from tkinter import messagebox
from ttkbootstrap import widgets
import glob
from datetime import datetime
import shutil
from pathlib import Path

class ListaReportes:

    # Iframe de Contenedor Tabla
    form_reportes = ""

    #callback
    callback = None
    
    #Componentes
    dateFechaInicio=None
    dateFechaFin=None
    selectProveedor=None
    tabla=None
    
    #Lista Proveedores
    list_proveedores=["Todos","Alicorp","Frutos y Especias S.A.C.","P&D Andina Alimentos S.A.C.","Casa Grande S.A.A.","Agroindustria Santa María S.A.C.","Perufarma S.A.","Leche Gloria S.A.","Molitalia S.A.","Kimberly-Clark Peru S.R.L."]

    def __init__(self,contenedor,callback = None):
        ##
        self.callback = callback

        self.form_reportes = contenedor
        ##
        self.Center(805,400)
        self.form_reportes.title("Lista de Reportes")
        print("Ingresando a la lista de Reportes ----")

    #region Funcion abrir table
    def Open(self):
        #Focus
        self.form_reportes.focus()
        self.form_reportes.grab_set()

        #Definiendo columnas
        columnas = ('Reportes','Fecha')


        self.tabla = ttk.Treeview(self.form_reportes, columns=columnas,show="headings",style="dark",selectmode="browse")
        self.tabla.place(x=0,y=50,height=350,width=790)

        scroll = ttk.Scrollbar(self.form_reportes,orient="vertical",command=self.tabla.yview)
        scroll.place(x=30+762, y=50, height=350)

        self.tabla.configure(yscrollcommand=scroll.set)

        #Definiendo header
        self.tabla.heading('Reportes',text="Reportes")
        self.tabla.heading('Fecha',text="Fecha de Modificación")
        
        ###
        #
        output_dir = f".\\reports"
        files = glob.glob(f"{output_dir}\*.xlsx")
        for f in files:
            temp_name = f.split("\\")[-1]
            if "~$" not in temp_name:
                temp_y = f.split("\\")[-1].split("-")[-4]
                temp_m = f.split("\\")[-1].split("-")[-3]
                temp_d = f.split("\\")[-1].split("-")[-2]
                self.tabla.insert('',END,values=(f"{temp_name}",f"{temp_y}/{temp_m}/{temp_d}"))
            
        # Funcion para Seleccionar
        def ItemSeleccionado(event):
            for select_i in self.tabla.selection():
                item = self.tabla.item(select_i)
                record = item['values']
                confirm_box = messagebox.askquestion('Confirmar','¿Desea descargar el archivo seleccionado?',icon='info')
                if confirm_box == "yes":
                    path_download=Path.home() / "Downloads"
                    path_report = f"./reports/{record[0]}"
                    shutil.copy(path_report,path_download)
                    messagebox.showinfo('Operación Exitosa',"DESCARGA EXITOSA.\nRevise su caperta de Descargas.")

        self.tabla.bind('<Double-1>',ItemSeleccionado)
        #

        #Cuadro de busqueda
        
        lblFechaInicio = ttk.Label(self.form_reportes,text="Fecha Inicio:")
        self.dateFechaInicio = widgets.DateEntry(self.form_reportes,width=10,bootstyle="secondary")
        
        lblFechaFin = ttk.Label(self.form_reportes,text="Fecha Fin:")
        self.dateFechaFin = widgets.DateEntry(self.form_reportes,width=10,bootstyle="secondary")
        
        lblProveedor = ttk.Label(self.form_reportes,text="Proveedor:")
        self.selectProveedor = ttk.Combobox(self.form_reportes,state="readonly",values=self.list_proveedores,width=30,style="secondary")
        
        btnBuscar = ttk.Button(self.form_reportes,width=10,text="Buscar",style="secondary",command=lambda: self.BuscarReportes())
        

        lblFechaInicio.place(x=10,y=10)
        self.dateFechaInicio.place(x=90,y=10)
        
        lblFechaFin.place(x=210,y=10)
        self.dateFechaFin.place(x=280,y=10)
        
        lblProveedor.place(x=400,y=10)
        self.selectProveedor.place(x=470,y=10)
        
        btnBuscar.place(x=700,y=10)

        ##Funcion Busqueda
        # def KeyPress(event):
        #     valor= inputBuscar.get()
        #     result = list(self.api.consultaUbigeo(valor))

        #     for row in self.tabla.get_children():
        #         self.tabla.delete(row)

        #     for d in result:
        #         self.tabla.insert('',END,values=list(d.values()))

        # inputBuscar.bind_all('<KeyPress>',KeyPress)
        ##


    #endregion

    #region Centrando Frame
    def Center(self,x,y):
        width = x # Ancho de Frame
        height = y # Alto de Frame

        screen_width = self.form_reportes.winfo_screenwidth()  # Ancho de Pantalla
        screen_height = self.form_reportes.winfo_screenheight() # Alto de Pantalla

        # Calculando coordenadas
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)

        self.form_reportes.geometry('%dx%d+%d+%d' % (width, height, x, y))
    #endregion
    
    def BuscarReportes(self):
        fecha_inicio = datetime.strptime(self.dateFechaInicio.entry.get(),"%d/%m/%Y")
        fecha_fin = datetime.strptime(self.dateFechaFin.entry.get(),"%d/%m/%Y")
        proveedor = self.selectProveedor.get()
        output_dir = f".\\reports"
        files = glob.glob(f"{output_dir}\*.xlsx")
        #
        for row in self.tabla.get_children():
                self.tabla.delete(row)
        #       
        for f in files:
            temp_name = f.split("\\")[-1]
            if "~$" not in temp_name:
                temp_y = f.split("\\")[-1].split("-")[-4]
                temp_m = f.split("\\")[-1].split("-")[-3]
                temp_d = f.split("\\")[-1].split("-")[-2]
                temp_date = datetime.strptime(f"{temp_d}/{temp_m}/{temp_y}","%d/%m/%Y")
                if fecha_fin>=temp_date and temp_date >= fecha_inicio:
                    if proveedor != "":
                        if proveedor == "Todos":
                            self.tabla.insert('',END,values=(f"{temp_name}",f"{temp_y}/{temp_m}/{temp_d}"))
                        elif proveedor.lower().split(" ")[0] in temp_name:
                            self.tabla.insert('',END,values=(f"{temp_name}",f"{temp_y}/{temp_m}/{temp_d}"))
                    else:
                        self.tabla.insert('',END,values=(f"{temp_name}",f"{temp_y}/{temp_m}/{temp_d}"))
                
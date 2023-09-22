from tkinter import *
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter.messagebox import showinfo
from ttkbootstrap import widgets
import json

class ListaReportes:

    # Iframe de Contenedor Tabla
    form_reportes = ""

    #callback
    callback = None
    
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
        columnas = ('Reportes','Fecha','Proveedor','Acciones')


        tabla = ttk.Treeview(self.form_reportes, columns=columnas,show="headings",style="dark",selectmode="browse")
        tabla.place(x=0,y=50,height=350)

        scroll = ttk.Scrollbar(self.form_reportes,orient="vertical",command=tabla.yview)
        scroll.place(x=30+762, y=50, height=350)

        tabla.configure(yscrollcommand=scroll.set)

        #Definiendo header
        tabla.heading('Reportes',text="Reportes")
        tabla.heading('Fecha',text="Fecha de Modificación")
        tabla.heading('Proveedor',text="Proveedor")
        tabla.heading('Acciones',text="Acciones")

        ###
        #Definiendo Body
        # o = open('./assets/UbigeoPeru.json')
        # data = json.load(o)
        # for d in data:
        #     tabla.insert('',END,values=list(d.values()))

        # Funcion para Seleccionar
        # def ItemSeleccionado(event):
        #     for select_i in tabla.selection():
        #         item = tabla.item(select_i)
        #         record = item['values']
        #         if len(str(list(record)[3])) == 5:
        #             new_ubigeo = "0"+str(list(record)[3])
        #         else:
        #             new_ubigeo = list(record)[3]
        #         print(new_ubigeo)
        #         self.callback(new_ubigeo,list(record)[0],list(record)[1],list(record)[2])
        #     self.form_reportes.destroy()

        # tabla.bind('<Double-1>',ItemSeleccionado)
        ##

        #Cuadro de busqueda
        
        lblFechaInicio = ttk.Label(self.form_reportes,text="Fecha Inicio:")
        dateFechaInicio = widgets.DateEntry(self.form_reportes,width=10,bootstyle="secondary")
        
        lblFechaFin = ttk.Label(self.form_reportes,text="Fecha Fin:")
        dateFechaFin = widgets.DateEntry(self.form_reportes,width=10,bootstyle="secondary")
        
        lblProveedor = ttk.Label(self.form_reportes,text="Proveedor:")
        selectProveedor = ttk.Combobox(self.form_reportes,state="readonly",values=self.list_proveedores,width=30,style="secondary")
        
        btnBuscar = ttk.Button(self.form_reportes,width=10,text="Buscar",style="secondary")
        

        lblFechaInicio.place(x=10,y=10)
        dateFechaInicio.place(x=90,y=10)
        
        lblFechaFin.place(x=210,y=10)
        dateFechaFin.place(x=280,y=10)
        
        lblProveedor.place(x=400,y=10)
        selectProveedor.place(x=470,y=10)
        
        btnBuscar.place(x=700,y=10)

        ##Funcion Busqueda
        # def KeyPress(event):
        #     valor= inputBuscar.get()
        #     result = list(self.api.consultaUbigeo(valor))

        #     for row in tabla.get_children():
        #         tabla.delete(row)

        #     for d in result:
        #         tabla.insert('',END,values=list(d.values()))

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

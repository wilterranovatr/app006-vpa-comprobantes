from tkinter import *
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from layouts.__list_reportes import ListaReportes
from tkinter import messagebox
from functions.__email_attachment import EmailAttachment
import os
from ttkbootstrap import widgets
from datetime import datetime

class FormMenu:
    
    # Formulario Menu
    form_menu = Tk()
    
    form_proceso = None
    form2_combo = None
    form2_combo2 = None
    form2_combo3 = None
    
    #Definiendo iconos
    logo = PhotoImage(file="./assets/icons/logo.png").subsample(2,2)
    logo2 = PhotoImage(file="./assets/icons/logo3.png").subsample(20)
    
    # Inicializando
    def __init__(self):
        self.form_menu.title("APP - Descarga de Comprobantes de Pago")
        self.Center(self.form_menu,690,450)
        print("Ingresando a formulario de DATOS ----")
        
        #Set image logo
        self.form_menu.iconphoto(False,self.logo)
        
    
    def Open(self):
        
        #Labels
        lblInformacion = ttk.Label(self.form_menu,text="-- Bienvenidos al app de descarga de comprobantes --")
        lblImagen = ttk.Label(self.form_menu,image=self.logo2)
        
        #Button
        btnDescargaCom = ttk.Button(self.form_menu,style="dark",width=35,text="INICIAR PROCESO",command= lambda: self.IniciarProceso())
        btnVerDescargas = ttk.Button(self.form_menu,style="dark",width=35,text="VER REPORTES",command= lambda: self.OpenListReportes())
        btnVerComprobantes = ttk.Button(self.form_menu,style="danger",width=35,text="VER COMPROBANTES DESCARGADOS",command= lambda: self.VerCarpetaComprobantes())
        
        #posicionando
        lblImagen.place(x=200,y=20)
        lblInformacion.place(x=210,y=200)
        btnDescargaCom.place(x=230,y=250)
        btnVerDescargas.place(x=230,y=300)
        btnVerComprobantes.place(x=230,y=350)
        
        #Show
        self.form_menu.mainloop()
    
    #region Centrando Frame
    def Center(self,form,x,y):
        width = x # Ancho de Frame
        height = y # Alto de Frame

        screen_width = form.winfo_screenwidth()  # Ancho de Pantalla
        screen_height = form.winfo_screenheight() # Alto de Pantalla

        # Calculando coordenadas
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)

        form.geometry('%dx%d+%d+%d' % (width, height, x, y))
    #endregion
    
    def IniciarProceso(self):
        self.form_proceso= Toplevel()
        self.form_proceso.title("Selecciona una opción")
        self.Center(self.form_proceso,400,250)
        #Focus
        self.form_proceso.focus()
        self.form_proceso.grab_set()
        ##
        form2_label = ttk.Label(self.form_proceso,text="--Seleccione un proveedor--")
        #self.form2_combo = ttk.Combobox(self.form_proceso,width=33,values=["ALICORP","FRUTOS Y ESPECIAS","P&D ANDINA ALIMENTOS","CASA GRANDE","AGROINDUSTRIA SANTA MARIA","PERUFARMA","LECHE GLORIA","MOLITALIA","KIMBERLY - CLARK PERU"],style="secondary")
        self.form2_combo = ttk.Combobox(self.form_proceso,width=33,values=["ALICORP","MOLITALIA"],style="secondary")
        form2_label2 = ttk.Label(self.form_proceso,text="-- Fecha Inicio --")
        self.form2_combo2 = widgets.DateEntry(self.form_proceso,width=30,bootstyle="secondary")
        form2_label3 = ttk.Label(self.form_proceso,text="-- Fecha Fin --")
        self.form2_combo3 = widgets.DateEntry(self.form_proceso,width=30,bootstyle="secondary")
        form2_btn = ttk.Button(self.form_proceso,width=10,text="ACEPTAR",style="dark",command= lambda: self.ProcesoETL())
        
        form2_label.place(x=120,y=10)
        self.form2_combo.place(x=90,y=30)
        form2_label2.place(x=120,y=70)
        self.form2_combo2.place(x=90,y=90)
        form2_label3.place(x=120,y=130)
        self.form2_combo3.place(x=90,y=150)
        form2_btn.place(x=150,y=200)
        ##
    
    def OpenListReportes(self):
        #self.form_menu.withdraw()
        form2 = Toplevel()
        ListaReportes(form2).Open()
        #FormDescargaTxt(form2,self.form_menu).Open()
    
    def ProcesoETL(self):
        # Fechas seleccionadas
        fec_ini = datetime.strptime(self.form2_combo2.entry.get(),"%d/%m/%Y")
        fec_fin = datetime.strptime(self.form2_combo3.entry.get(),"%d/%m/%Y")
        proveedor = self.form2_combo.get()
        prov_name=""
        if proveedor == "ALICORP":
            prov_name = "alicorp"
        elif proveedor == "FRUTOS Y ESPECIAS":
            prov_name = "frutos_especias"
        elif proveedor == "P&D ANDINA ALIMENTOS":
            prov_name = "p_d_andina_alimentos"
        elif proveedor == "CASA GRANDE":
            prov_name = "casa_grande"
        elif proveedor == "CASA GRANDE":
            prov_name = "casa_grande"
        elif proveedor == "AGROINDUSTRIA SANTA MARIA":
            prov_name = "agroindustria_santa_maria"
        elif proveedor == "PERUFARMA":
            prov_name = "perufarma"
        elif proveedor == "LECHE GLORIA":
            prov_name = "leche_gloria"
        elif proveedor == "MOLITALIA":
            prov_name = "molitalia"
        elif proveedor == "KIMBERLY - CLARK PERU":
            prov_name = "kimberly_clark_peru"
        
        if prov_name != "":
            # Saliendo de Ventana
            self.form_proceso.destroy()
                        
            ## Ejecutando ETL
            etl_function = EmailAttachment()
            etl_function.startProccessETL(prov_name,fec_ini=fec_ini,fec_fin=fec_fin)
            
            # Visualizar reporte
            confirm_box = messagebox.askquestion('Confirmar','¿Desea abrir el reporte creado?',icon='info')
            if confirm_box == "yes":
                path_report = f"./reports/{etl_function.getPathReporte()}.xlsx"
                command = f'start excel "{path_report}"'
                os.system(command)
                
        else:
            messagebox.showerror("Alerta","Seleccione un proveedor para poder continuar.")
            
    def VerCarpetaComprobantes(self):
        path_appdata=os.getenv('APPDATA')
        output_dir = f'{path_appdata}/Comprobantes Terranova/PDF/'
        os.startfile(output_dir)
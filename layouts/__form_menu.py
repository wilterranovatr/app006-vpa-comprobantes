from tkinter import *
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from layouts.__list_reportes import ListaReportes
from tkinter import messagebox

class FormMenu:
    
    # Formulario Menu
    form_menu = Tk()
    
    #Definiendo iconos
    logo = PhotoImage(file="./assets/icons/logo.png").subsample(2,2)
    logo2 = PhotoImage(file="./assets/icons/logo3.png").subsample(20)
    
    # Inicializando
    def __init__(self):
        self.form_menu.title("APP - Descarga de Comprobantes de Pago")
        self.Center(690,400)
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
        
        #posicionando
        lblImagen.place(x=200,y=20)
        lblInformacion.place(x=210,y=200)
        btnDescargaCom.place(x=230,y=250)
        btnVerDescargas.place(x=230,y=300)
        
        #Show
        self.form_menu.mainloop()
    
    #region Centrando Frame
    def Center(self,x,y):
        width = x # Ancho de Frame
        height = y # Alto de Frame

        screen_width = self.form_menu.winfo_screenwidth()  # Ancho de Pantalla
        screen_height = self.form_menu.winfo_screenheight() # Alto de Pantalla

        # Calculando coordenadas
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)

        self.form_menu.geometry('%dx%d+%d+%d' % (width, height, x, y))
    #endregion
    
    def IniciarProceso(self):
        confirm_box = messagebox.askquestion('Confirmar','¿Desea iniciar el proceso de descarga?',icon='info')
        if confirm_box =='yes':
            form2 = Toplevel()
    
    def OpenListReportes(self):
        #self.form_menu.withdraw()
        form2 = Toplevel()
        ListaReportes(form2).Open()
        #FormDescargaTxt(form2,self.form_menu).Open()
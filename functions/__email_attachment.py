from pathlib import Path
import win32com.client
from datetime import datetime, timedelta
import os
import io
import pdfminer 
# from pdfminer.converter import TextConverter
# from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
# from pdfminer.pdfpage import PDFPage
# from pdfminer.layout import LAParams

import camelot

import slate3k as slate

import tabula

class EmailAttachment:
    # Fecha actual
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    
    def __init__(self) -> None:
        pass
    
    def download_attachment(self, email, find, provider):
        # Create output folder
        path_appdata=os.getenv('APPDATA')
        output_dir = Path(f"{path_appdata}/Comprobantes Terranova/Comprobantes/{provider}")
        if not output_dir.is_dir():
            os.mkdir(output_dir)
        
        # Connect to outlook application
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        
        # Get messages
        messages = inbox.Items

        # Filter messages - Optional
        received_dt = datetime.now() - timedelta(days=180)
        received_dt = received_dt.strftime('%m/%d/%Y')
        messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
        messages = messages.Restrict(f"[SenderEmailAddress] = '{email}'")

        for message in messages:
            attachments = message.Attachments
            # Save attachments
            for attachment in attachments:
                attachment_file_name = str(attachment)
                if attachment_file_name.find(find) != -1 and attachment_file_name.find(".pdf") != -1: #Notas de crédito alicorp
                    attachment.SaveAsFile(output_dir / attachment_file_name)
                    
    def descargaxProv(self):
        proveedor_correo = {
            "alicorp":{
                "correo": "dte_20100055237@paperless.pe",
                "razon_social":"ALICORP SAA",
                "ruc":"20100055237",
                "comprobantes":["01","07","08","40"]
            },
            "frutos_especias":{
                "correo": "felectronica@frutosyespecias.com.pe",
                "razon_social":"FRUTOS Y ESPECIAS S.A.C.",
                "ruc":"20207845044",
                "comprobantes":["01","09","07"]
            },
            "p_d_andina_alimentos":{
                "correo": "facturacionelectronica@pdandina.com.pe",
                "razon_social":"P&D ANDINA ALIMENTOS S.A",
                "ruc":"20205922149",
                "comprobantes":["01","07"]
            },
            "casa_grande":{
                "correo": "de_20131823020@azucarperu.com.pe",
                "razon_social":"CASA GRANDE SOCIEDAD ANONIMA ABIERTA",
                "ruc":"20131823020",
                "comprobantes":["01","07"]
            },
            "agroindustria_santa_maria":{
                "correo": "comprobantes@efacturacion.pe",
                "razon_social":"AGROINDUSTRIA SANTA MARIA S.A.C.",
                "ruc":"20100166144",
                "comprobantes":["01","07","40"]
            },
            "perufarma":{
                "correo": "comprobantes@efacturacion.pe",
                "razon_social":"PERUFARMA S A",
                "ruc":"20100052050",
                "comprobantes":["01","09","07"]
            },
            "leche_gloria":{
                "correo": "de_20100190797@gloria.com.pe",
                "razon_social":"LECHE GLORIA SOCIEDAD ANONIMA - GLORIA S.A.",
                "ruc":"20100190797",
                "comprobantes":["01","09","07","08"]
            },
            "molitalia":{
                "correo": "dte_20100035121@paperless.pe",
                "razon_social":"MOLITALIA S.A",
                "ruc":"20100035121",
                "comprobantes":["01","07"]
            },
            "kimberly_clark_peru":{
                "correo": "dte_20100152941@paperless.pe",
                "razon_social":"KIMBERLY-CLARK PERU S.R.L.",
                "ruc":"20100152941",
                "comprobantes":["01","07","08"]
            }
        }
        self.download_attachment(email="dte_20100055237@paperless.pe", find="20100055237_07", provider="alicorp")
    
    def ReadDataPdf(self,rutaPDF):
        
        ### USANDO PDFMINER
        # resource = PDFResourceManager()
        
        # flujo = io.BytesIO()
        
        # convertidor = TextConverter(resource,flujo,codec='utf-8',laparams=LAParams(char_margin=20))
        
        # interprete = PDFPageInterpreter(resource,convertidor)
        
        # with open(rutaPDF,'rb') as archivo_pdf:
        #     for pagina in PDFPage.get_pages(archivo_pdf):
        #         interprete.process_page(pagina)
        
        # contain = flujo.getvalue().decode("utf-8")
        # print(contain)
        
        # with open(rutaTXT,'w',encoding="utf-8") as archivo_txt:
        #     archivo_txt.write(contain)
        
        # flujo.close()
        
        # convertidor.close()
        
        ### USANDO CAMELOT
        # parametros = {
        #     "table_regions": ['0,0,500,100'],  # Especifica el área de la página donde se encuentran las tablas
        #     "strip_text": "\n",  # Elimina saltos de línea en los valores de las celdas
        #     # Agrega más parámetros de configuración según tus necesidades
        # }
        # tables = camelot.read_pdf(rutaPDF,pages="all",flavor="stream",table_threshold=0.9)
        # df = tables[0].df

        # # Imprimir el dataframe
        # print(df)
        
        ### USANDO SLATE
        # with open(rutaPDF,'rb') as f:
        #     doc = slate.PDF(f)
        
        # for page in doc:
        #     print(page.text)
        
        ### USANDO TABULA
        # df = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(20,303,114,476))
        # #df = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(198,2,210,585))
        # #df = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,guess=False)
        # for d in df:
        #     print("Aqui")
        #     print(d)
        
        if "20100055237_01" in rutaPDF:   ##Alicorp  #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(179,28,224,583))
            print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(541,429,635,596))
            #print(costos[0])
            ##  
            
        elif "20100055237_07" in rutaPDF: ##Alicorp  #Nota de Credito
            cabecera = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(28,356,132,584))
            #print(cabecera[0])
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(169,29,212,583))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(528,415,606,583))
            #print(costos[0])
        elif "20100055237_08" in rutaPDF: ##Alicorp  #Nota de Debito
            cabecera = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(29,357,132,585))
            #print(cabecera[0])
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(168,27,213,583))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(529,415,608,586))
            #print(costos[0])
        elif "20100055237_040" in rutaPDF: ##Alicorp  #Comprobante de Percepción Electrónico
            cabecera = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(29,331,131,566))
            #print(cabecera[0])
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(206,30,761,569))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(764,305,794,562))
            #print(costos[0])
        elif "20207845044-01" in rutaPDF: ##Frutos Especies #Factura
            cabecera_y_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(36,325,248,577))
            #print(cabecera_y_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(245,3,811,580))
            #print(costos[0])
        elif "20207845044-07" in rutaPDF: ##Frutos Especies #Nota de Credito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(142,356,218,557))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(240,24,264,570))
            #print(datos_generales_2[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(279,20,807,578))
            #print(costos[0])
        elif "20207845044-09" in rutaPDF: ##Frutos Especies #Guia de Remision Remitente Electrónica
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(271,361,306,559))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(335,20,361,575))
            #print(datos_generales_2[0])
        elif "20205922149-01" in rutaPDF: ##P D Andina #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(102,415,191,559))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(561,398,680,584))
            #print(costos[0])
        elif "20205922149-07" in rutaPDF: ##P D Andina #Nota de Credito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(108,415,145,575))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(659,397,778,591))
            #print(costos[0])
        elif "20131823020-01" in rutaPDF: ##Casa Grande #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(141,40,196,374))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(524,409,634,558))
            #print(costos[0])
        elif "20131823020-07" in rutaPDF: ##Casa Grande #Nota de Credito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(155,41,374,199))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(204,38,229,557))
            #print(datos_generales_2[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(508,406,585,569))
            #print(costos[0])
        elif "20100166144-01" in rutaPDF: ##Agro Industria Santa Maria  #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(101,21,176,290))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(174,23,808,581))
            #print(costos[0])
        elif "20100166144-07" in rutaPDF: ##Agro Industria Santa Maria  #Nota de Credito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(179,20,223,575))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(220,20,808,581))
            #print(costos[0])
        elif "20100166144-40" in rutaPDF: ##Agro Industria Santa Maria  #Comprobante de Percepción
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(114,20,168,576))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(167,21,806,580))
            #print(costos[0])
        elif "20100052050-01" in rutaPDF: ##PeruFarma  #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(147,20,185,809))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(198,514,781,819))
            #print(costos[0])
        elif "20100052050-09" in rutaPDF: ##PeruFarma  #Guia de Remision
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(140,22,216,574))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(342,22,421,574))
            #print(datos_generales_2[0])
        elif "20100052050-07" in rutaPDF: ##PeruFarma  #Notas de Credito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(144,20,172,576))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(432,22,458,382))
            #print(datos_generales_2[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(288,516,369,815))
            #print(costos[0])
        elif "20100190797-01" in rutaPDF: ##Leche Gloria  #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(142,39,189,554))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(407,401,523,575))
            #print(costos[0])
        elif "20100190797-09" in rutaPDF: ##Leche Gloria  #Guia de Remision
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(142,39,188,554))
            #print(datos_generales[0])
        elif "20100190797-07" in rutaPDF: ##Leche Gloria  #Notas de Crédito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(154,40,204,374))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(518,179,553,386))
            #print(datos_generales_2[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(542,402,679,580))
            #print(costos[0])
        elif "20100190797-08" in rutaPDF: ##Leche Gloria #Notas de debito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(154,39,204,374))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(507,180,554,384))
            #print(datos_generales_2[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(547,405,670,569))
            #print(costos[0])
        elif "20100035121-01" in rutaPDF: ##Molitalia  #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(173,514,244,577))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(588,26,633,578))
            #print(costos[0])
        elif "20100035121-07" in rutaPDF: ##Molitalia #Notas de Crédito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(195,20,228,576))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(633,19,669,580))
        elif "20100152941_08" in rutaPDF: ##Kimberly #Nota de Debito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(109,362,171,575))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(278,20,322,306))
            #print(datos_generales_2[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(640,359,732,576))
            #print(costos[0])
        elif "20100152941_01" in rutaPDF: ##Kimberly  #Factura
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(108,359,162,574))
            #print(datos_generales[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(638,369,728,576))
            #print(costos[0])
        elif "20100152941_07" in rutaPDF: ##Kimberly #Notas de Crédito
            datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(109,360,171,557))
            #print(datos_generales[0])
            datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(269,19,325,306))
            #print(datos_generales_2[0])
            costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(643,358,729,580))
            #print(costos[0])
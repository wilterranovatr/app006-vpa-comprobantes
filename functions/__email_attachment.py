from pathlib import Path
import win32com.client
from datetime import datetime, timedelta
import os
import json
import pandas as pd
import openpyxl
import shutil
import tabula
import multiprocessing
import glob
import time as t
from tkinter import messagebox
import requests

class EmailAttachment:
    # Fecha actual
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    
    #Hora actual
    hora_actual = datetime.now().strftime("%H%M%S")
    
    # Datos Reporte
    data = []
    
    # Path Report
    path_report = ""
    
    def __init__(self) -> int:
        pass
        
    def download_attachment(self, email,ruc, find,provider,fec_ini =None, fec_fin = None):
        # Create output folder
        path_appdata=os.getenv('APPDATA')
        nom_folder = fec_fin.strftime("%d-%m-%Y")
        output_dir = Path(f'{path_appdata}/Comprobantes Terranova/{provider}/{nom_folder.split("-")[2]}{nom_folder.split("-")[1]}{nom_folder.split("-")[0]}')
        if not output_dir.is_dir():
            os.makedirs(output_dir)
        else:
            shutil.rmtree(output_dir)
            t.sleep(2)
            os.makedirs(output_dir)
        
        # Connect to outlook application
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        
        # Get messages
        messages = inbox.Items
        # Filter messages - Optional
        received_de = fec_ini.strftime('%d/%m/%Y')
        received_ha = fec_fin + timedelta(days=1)
        received_ha = received_ha.strftime('%d/%m/%Y')
        #received_dt = received_dt.strftime('%m/%d/%Y')
        #print(messages)
        #messages = messages.Restrict(f"[SenderEmailAddress] = '{email}'")
        messages = messages.Restrict(f"[SenderEmailAddress] = '{email}' Or [Subject] = 'COMPROBANTES REENVIADOS | {ruc}'")
        messages = messages.Restrict("[ReceivedTime] >= '" + received_de + "' AND [ReceivedTime] <= '"+received_ha+"'")
        #
        if provider == "agroindustria_santa_maria":
            #messages = messages.Restrict(f"[Body] LIKE '%{provider}%'")
            messages = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%AGROINDUSTRIA SANTA MARIA%'")
            folio = 1
            ## Factura
            message_factura = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%Tipo de documento: Factura%'")
            folio_temp =self.download_attachment_2(messages=message_factura,folio_i=folio,output_dir=output_dir,ruc=ruc,tipo="01")
            ## Nota de Crédito
            message_nc = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%Tipo de documento: Nota de C%'")
            folio_temp_2=self.download_attachment_2(messages=message_nc,folio_i=folio_temp,output_dir=output_dir,ruc=ruc,tipo="07")
            ## Comprobantes
            message_perc = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%comprobante de percepción%'")
            folio_temp_3=self.download_attachment_2(messages=message_perc,folio_i=folio_temp_2,output_dir=output_dir,ruc=ruc,tipo="40")
        elif provider == "perufarma":
            #messages = messages.Restrict(f"[Body] LIKE '%{provider}%'")
            messages = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%PERUFARMA%'")
            folio = 1
            ## Factura
            message_factura = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%Factura%'")
            folio_temp =self.download_attachment_2(messages=message_factura,folio_i=folio,output_dir=output_dir,ruc=ruc,tipo="01")
            ## Nota de Crédito
            message_nc = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%Nota de Credito%'")
            folio_temp_2=self.download_attachment_2(messages=message_nc,folio_i=folio_temp,output_dir=output_dir,ruc=ruc,tipo="07")
            ## Comprobantes
            message_rem = messages.Restrict(f"@SQL=""urn:schemas:httpmail:textdescription"" like '%Guía de remisión%'")
            folio_temp_3=self.download_attachment_2(messages=message_rem,folio_i=folio_temp_2,output_dir=output_dir,ruc=ruc,tipo="09")
        else:  
            #
            #messages = messages.Restrict(f"[Subject] = 'COMPROBANTES REENVIADOS | {ruc}'")
            print("--- TOTAL COMPROBANTES ENCONTRADOR",len(messages))
            folio = 1
            for message in messages:
                attachments = message.Attachments
                # Save attachments
                for attachment in attachments:
                    attachment_file_name = str(attachment)
                    for f in find:
                        if attachment_file_name.find(f) != -1 and attachment_file_name.find(".pdf") != -1:
                            nom_file_final = str(folio)+"_"+attachment_file_name
                            folio = folio + 1
                            attachment.SaveAsFile(output_dir / nom_file_final)
    def download_attachment_2(self,messages,folio_i,output_dir,ruc,tipo):
        folio = folio_i
        messages_temp = []
        names_files = []
        for msg in messages:
            temp_msg = msg.Body.splitlines()
            messages_temp.append([linea for linea in temp_msg if 'PDF' in linea])
            names_files.append([linea for linea in temp_msg if ('Serie y número' in linea or 'Número de Comprobante' in linea)])

        if ruc == "20100052050": ## Perufarma
            messages = list(map(lambda x: x[0].split("<")[1].split(">")[0],messages_temp))
            names = list(map(lambda x: x[0].split(" ")[-2].split("\t")[1] if x[0].split(" ")[-2].split("\t")[0] == "" else x[0].split(" ")[-2].split("\t")[0],names_files))
            print(len(messages))
            print(len(names))
        else:
            names = list(map(lambda x: x[0].split(" ")[-2].split("\t")[0] if x[0].split(" ")[-1] == "" else x[0].split(" ")[-1],names_files))
            messages = list(map(lambda x: x[0].split("<")[-1].split(">")[0],messages_temp))
        #
        for m in range(len(messages)):
            respuesta = requests.get(messages[m], stream=True)
            # Aisla el nombre del archivo PDF de la URL
            nombre_archivo = f"{folio}_{ruc}-{tipo}_{names[m]}.pdf"
            folio = folio +1
            if respuesta.status_code == 200:
                # Guarda en la carpeta especificada
                ruta_archivo = os.path.join(output_dir, nombre_archivo)
                with open(ruta_archivo, 'wb') as archivo_pdf:
                    archivo_pdf.write(respuesta.content)
        return folio

    #region  Main
    def startProccessETL(self,provider, fec_ini = None,fec_fin = None):
        self.data=[]
        # Descargando comprobantes de un proveedor
        print("--- INICIANDO PROCESO DE DESCARGA DE COMPROBANTES DEL PROVEEDOR ",provider)
        try:
            self.descargaxProv(provider,f_ini=fec_ini,f_fin=fec_fin)
        except Exception as e:
            print("--- ERROR - DESCARGA DE COMPROBANTES POR PROVEEDOR: ",e)
            messagebox.showerror("Alerta","No se puede continuar debido a que hay un documento o carpeta abierta.\nPorfavor, cierre para poder continuar.")
        else:
            # Lectura de pdf 
            print("--- EXTRAYENDO Y TRANSFORMANDO DATOS DE LOS COMPROBANTES")
            pool = multiprocessing.Pool() ##Paralelismo
            try:
                #
                path_appdata=os.getenv('APPDATA')
                nom_folder = fec_fin.strftime("%d-%m-%Y")
                print("PROVEEDOR: ",provider)
                output_dir = f'{path_appdata}\Comprobantes Terranova\{provider}\{nom_folder.split("-")[2]}{nom_folder.split("-")[1]}{nom_folder.split("-")[0]}'
                files = glob.glob(f"{output_dir}\*.pdf")
                #
                self.data.append(pool.map(self.ReadDataPdf,files))
                #print(F"SE HA LEIDO UN TOTAL DE {len(self.data)} DOCUMENTOS.")
                #
            except Exception as e:
                pool.close()
                pool.join()
                print(f"--- ERROR - LECTURA DE PDF: ",e)
            else:
                pool.close()
                pool.join()
                try:
                    print("--- PROCESANDO DATOS DE LOS COMPROBANTES")
                    ## Creando Reporte
                    self.createReport(self.data[0],provider)
                except Exception as e:
                    print("--- ERROR - CREACION DE REPORTE: ",e)
                else:
                    print("--- REPORTE CREADO SATISFACTORIAMENTE ---")
    #endregion
                   
    def createReport(self,datos,provider_name):
        ## Probando excel
        seed_cell_excel = 10
        count_cells = 0
        name_report= f"{provider_name}-{self.fecha_actual}-{self.hora_actual}"
        shutil.copy("./assets/model_reporte.xlsx",f"./reports/{name_report}.xlsx")
        data_sheet = openpyxl.load_workbook(f"./reports/{name_report}.xlsx")
        sheet = data_sheet.active
        sheet['C5']= datetime.now()
        for da in datos:
            row_number = count_cells + seed_cell_excel
            count_cells=count_cells+1
            ####
            for d in da:
                sheet[f'B{row_number}'] = d["folio"]
                sheet[f'C{row_number}'] = d["ruc"]
                sheet[f'D{row_number}'] = d["razon_social"]
                sheet[f'E{row_number}'] = d["tipo_comprobante"]
                sheet[f'F{row_number}'] = d["nro_comprobante"]
                sheet[f'G{row_number}'] = d["fecha_emision"]
                sheet[f'H{row_number}'] = "\n".join(d["doc_vinculado"])
                sheet[f'I{row_number}'] = d["igv"]
                sheet[f'J{row_number}'] = d["imp_total"]
                sheet[f'K{row_number}'] = d["imp_percepcion"]
            ####
        data_sheet.save(f"./reports/{name_report}.xlsx")
        
        ##
        self.path_report = name_report
                    
    def descargaxProv(self,provider,f_ini = None, f_fin = None):
        o = open('./assets/data_proveedor.json')
        data_proveedor = json.load(o)
        self.proveedor=data_proveedor[provider]
        ##
        if provider == "alicorp" or provider=="kimberly_clark_peru" or provider=="molitalia":
            temp_finds = f'{self.proveedor["ruc"]}_'+f',{self.proveedor["ruc"]}_'.join(self.proveedor["comprobantes"])
        else:
            temp_finds = f'{self.proveedor["ruc"]}-'+f',{self.proveedor["ruc"]}-'.join(self.proveedor["comprobantes"])
        
        finds = temp_finds.split(",")
        self.download_attachment(email=self.proveedor["correo"],ruc=self.proveedor["ruc"],find=finds,provider=provider,fec_ini=f_ini,fec_fin=f_fin)
    
    def ReadDataPdf(self,rutaPDF):
        data_temp = []
        try:
            if "20100055237_01" in rutaPDF:   ##Alicorp  #Factura
                #
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100055237_01","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(179,28,224,583))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(541,429,635,596))
                #print(costos[0])
                ##
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = df_datos["Unnamed: 3"][3]
                doc_vinculado = ["-"]
                igv = df_costos[df_costos.iloc[0].index[2]][5]
                imp_total = df_costos[df_costos.iloc[0].index[2]][6]
                try:
                    percepcion = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(577,29,642,425))
                    #print(percepcion[0])
                    df_percpecion = percepcion[0]
                    imp_percepcion = df_percpecion.columns.to_numpy()[0].split(" ")[-1]
                except:
                    imp_percepcion = "-"
                ##
            elif "20100055237_07" in rutaPDF: ##Alicorp  #Nota de Credito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100055237_07","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(169,29,212,583))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(215,28,532,594))
                #print(datos_generales_2)
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(528,415,606,583))
                #print(costos[0])
                ##
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = df_datos["Fecha de"][2]
                #
                doc_vinculado = []
                for df in datos_generales_2:
                    temp = df[df["Unnamed: 1"]=="FACTURA ELECTRONICA"][df.iloc[0].index[2]]
                    doc_vinculado.extend(temp.values.tolist())
                #
                #print(doc_vinculado)#
                igv = df_costos[df_costos.iloc[0].index[2]][4]
                imp_total = df_costos[df_costos.iloc[0].index[2]][5]
                imp_percepcion = "-"
                ##
                
            elif "20100055237_08" in rutaPDF: ##Alicorp  #Nota de Debito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100055237_08","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(168,27,213,583))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(215,28,532,594))
                #print(datos_generales_2)
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(529,415,608,586))
                #print(costos[0])
                ##
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTA DE DÉBITO"
                fecha_emision = df_datos["Fecha de"][2]
                #
                doc_vinculado = []
                for df in datos_generales_2:
                    temp = df[df["Unnamed: 1"]=="FACTURA ELECTRONICA"][df.iloc[0].index[2]]
                    doc_vinculado.extend(temp.values.tolist())
                #
                #print(doc_vinculado)#
                igv = df_costos[df_costos.iloc[0].index[2]][4]
                imp_total = df_costos[df_costos.iloc[0].index[2]][5]
                imp_percepcion = "-"
                #print(imp_total,igv)
                ##
            elif "20100055237_040" in rutaPDF: ##Alicorp  #Comprobante de Percepción Electrónico
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100055237_040","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                cabecera = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(29,331,131,566))
                #print(cabecera[0])
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(206,30,761,569))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(764,305,794,562))
                #print(costos[0])
                ##
                df_cabecera= cabecera[0]
                df_costos = costos[0]
                #
                tipo_comprobante="COMPROBANTE DE PERCEPCIÓN E."
                fecha_emision = str(df_cabecera["R.U.C. 20100055237"][3]).split(":")[1]
                #print(fecha_emision)
                #
                doc_vinculado = []
                for df in datos_generales:
                    temp = df[df["Unnamed: 0"]=="FACTURA"]["Unnamed: 1"]
                    doc_vinculado.extend(temp.values.tolist())
                #
                #print(doc_vinculado)#
                igv = "-"
                #imp_total = df_costos.columns.to_numpy()[3]
                imp_total = "-"
                imp_percepcion = df_costos.head().to_numpy()[0][3]
                ##
            elif "20207845044-01" in rutaPDF: ##Frutos Especies #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20207845044-01","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                cabecera_y_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(36,325,248,577))
                #print(cabecera_y_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(245,3,811,580))
                #print(costos[0])
                ##
                df_cabecera= cabecera_y_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = str(df_cabecera["R.U.C.:  20207845044"][2]).split(":")[1]
                #print(fecha_emision)
                #
                doc_vinculado = ["-"]
                try:
                    temp_descripcion = df_costos[["DESCRIPCIÓN P.U."]].dropna(subset=["DESCRIPCIÓN P.U."]).reset_index(drop=True)
                except:
                    temp_descripcion = df_costos[["P.U."]].dropna(subset=["P.U."]).reset_index(drop=True)
                temp_valor = df_costos[["VALOR"]].dropna(subset=["VALOR"]).reset_index(drop=True)
                df_concat = pd.merge(temp_descripcion,temp_valor, left_index=True,right_index=True,how="outer")
                try:
                    igv= round(float(str(df_concat[df_concat["DESCRIPCIÓN P.U."]=="Base imponible:"]["VALOR"].to_numpy()[0]).replace(",","")) * 0.18 ,2)
                except:
                    igv= round(float(str(df_concat[df_concat["P.U."]=="Base imponible:"]["VALOR"].to_numpy()[0]).replace(",","")) * 0.18 ,2)
                imp_total = temp_valor.iloc[-1].to_numpy()[0]
                imp_percepcion = "-"
                ##
            elif "20207845044-07" in rutaPDF: ##Frutos Especies #Nota de Credito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20207845044-07","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(142,356,218,557))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(215,24,719,570))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(279,409,807,578))
                #print(costos[0])
                ##
                df_generales= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = str(df_generales.columns.to_numpy()[1])
                #print(fecha_emision)
                #
                doc_vinculado = []
                for df in datos_generales_2:
                    temp = df[df["Documento que afecta la nota de Crédito:"].str.contains("Nro. de Doc",na=False)]["Documento que afecta la nota de Crédito:"].values.tolist()
                    temp_transf = list(map(lambda x : x.split(":")[2],temp))
                    #print(temp_transf)
                    doc_vinculado.extend(temp_transf)
                #
                #print(doc_vinculado)#
                igv = df_costos[df_costos["P.U.."].str.contains("IGV (18%):",na=False,regex=True)]["VALOR"]
                # except:
                #     igv = df_costos[df_costos["P.U.."].str.contains("IGV (18%):",na=False,regex=True)]["VALOR"] if df_costos[df_costos["P.U.."].str.contains("IGV (18%)",na=False)]["Unnamed: 1"].size == "0" else df_costos[df_costos["P.U.."].str.contains("IGV (18%)",na=False)]["Unnamed: 1"]
                igv = "0.00" if igv.size == 0 else igv.values.tolist()[0]
                
                imp_total = df_costos[df_costos["P.U.."].str.contains("Importe Total:",na=False)]["VALOR"]
                # except:
                #     imp_total = df_costos[df_costos["P.U.."].str.contains("Importe Total:",na=False)]["VALOR"] if df_costos[df_costos["P.U.."].str.contains("Importe Total",na=False)]["Unnamed: 1"].size == "0" else df_costos[df_costos["P.U.."].str.contains("Importe Total",na=False)]["Unnamed: 1"]
                imp_total = "0.00" if imp_total.size == 0 else imp_total.values.tolist()[0]
                imp_percepcion = "-"
                #print(imp_total)
                ##
            elif "20207845044-09" in rutaPDF: ##Frutos Especies #Guia de Remision Remitente Electrónica
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20207845044-09","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(271,361,306,559))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(323,20,684,575))
                #print(datos_generales_2[0])
                ##
                df_generales= datos_generales[0]
                df_generales_2 = datos_generales_2[0]
                #
                tipo_comprobante="GUÍA DE REMISION REMITENTE ELECTRÓNICO"
                fecha_emision = str(df_generales.head().to_numpy()[0][1]).replace(":","")
                #print(fecha_emision)
                #
                temp_vinculado = df_generales_2[df_generales_2["DOCUMENTOS RELACIONADOS:"].str.contains("Factura",na=False)]["DOCUMENTOS RELACIONADOS:"]
                doc_vinculado = list(map(lambda x : x.split(":")[1],temp_vinculado.to_numpy()))
                igv= "-"
                imp_total = "-"
                imp_percepcion = "-"
                #print(imp_total)
                ##
            elif "20205922149-01" in rutaPDF: ##P D Andina #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20205922149-01","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(102,415,191,559))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(540,398,680,584))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = df_datos.columns.to_numpy()[2]
                doc_vinculado = ["-"]
                igv = df_costos[df_costos.iloc[0].index[2]][5]
                imp_total = df_costos[df_costos.iloc[0].index[2]][6]
                imp_percepcion = "-"
                ####
            elif "20205922149-07" in rutaPDF: ##P D Andina #Nota de Credito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20205922149-07","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(108,415,145,575))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(194,13,663,588))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(540,397,814,591))
                #print(costos[0])
                ###
                df_datos= datos_generales[0]
                df_datos_2= datos_generales_2[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = df_datos.columns.to_numpy()[2]
                temp_vinculado = df_datos_2[df_datos_2.iloc[:,1].str.contains("FACTURA",na=False)]
                temp_vinculado_2 = temp_vinculado.iloc[:,1]
                doc_vinculado = list(map(lambda x: x.split("-")[1]+"-"+x.split("-")[2], temp_vinculado_2.to_numpy()))
                #print(doc_vinculado)
                #
                igv = df_costos[df_costos.iloc[0].index[2]][5]
                imp_total = df_costos[df_costos.iloc[0].index[2]][6]
                imp_percepcion = "-"

                ###
            elif "20131823020-01" in rutaPDF: ##Casa Grande #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20131823020-01","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(141,40,196,374))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(524,409,634,558))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = df_datos.columns.to_numpy()[1].split(" ")[1]
                doc_vinculado = ["-"]
                temp_igv = df_costos[df_costos.iloc[:,0].str.contains("I.G.V",na=False)]
                igv = temp_igv.iloc[:,2].to_numpy()[0]
                temp_total = df_costos[df_costos.iloc[:,0].str.contains("Importe Total",na=False)]
                imp_total = temp_total.iloc[:,2].to_numpy()[0]
                #print(imp_total)
                imp_percepcion = "-"
                ####
            elif "20131823020-07" in rutaPDF: ##Casa Grande #Nota de Credito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20131823020-07","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(155,41,374,199))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(202,37,506,557))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(508,406,585,569))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = df_datos.columns.to_numpy()[0].split(" ")[4]
                temp_vinculado = df_datos[df_datos.iloc[:,0].str.contains("F003",na=False)]
                tmp_doc_vinculado = temp_vinculado.iloc[:,0]
                doc_vinculado = list(map(lambda x: x.split(" ")[0], tmp_doc_vinculado.to_numpy()))
                temp_igv = df_costos[df_costos.iloc[:,0].str.contains("I.G.V",na=False)]
                igv = temp_igv.iloc[:,2]
                temp_total = df_costos[df_costos.iloc[:,0].str.contains("Importe Total",na=False)]
                imp_total = temp_total.iloc[:,2]
                #print(imp_total)
                imp_percepcion = "-"
                ####
            elif "20100166144-01" in rutaPDF: ##Agro Industria Santa Maria  #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100166144-01","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(101,21,176,290))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(174,23,808,581))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                temp_emision = df_datos[df_datos.iloc[:,0].str.contains("Emisión",na=False)]
                fecha_emision = temp_emision.iloc[:,1].to_numpy()[0].split(" ")[1]
                doc_vinculado = ["-"]
                igv = df_costos[df_costos["Precio Valor"].str.contains("IGV",na=False)]["Valor"]
                imp_total =df_costos[df_costos["Precio Valor"].str.contains("Importe Total",na=False)]["Valor"]
                imp_percepcion = "-"
                ####
            elif "20100166144-07" in rutaPDF: ##Agro Industria Santa Maria  #Nota de Credito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100166144-07","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(99,21,176,577))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(181,20,719,280))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(220,20,808,581))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_datos_2= datos_generales_2[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = df_datos[df_datos["Cliente"].str.contains("Emisión",na=False)][": TERRANOVA TRADING S.A.C."].to_numpy()[0].split(" ")[1]
                #print(df_datos_2.iloc[1].index[0])
                temp_index = df_datos_2.index[df_datos_2[df_datos_2.iloc[1].index[0]].str.contains("Nro. Doc. Ref.",na=False)].to_list()
                doc_vinculado=[]
                for t in temp_index:
                    temp_list_doc = list(map(lambda x: x.split(" ")[0],list(df_datos_2.iloc[t+1])))
                    doc_vinculado.extend(temp_list_doc)
                #print(doc_vinculado)
                igv = df_costos[df_costos["Precio Valor"].str.contains("IGV",na=False)]["Unnamed: 6"]
                imp_total =df_costos[df_costos["Precio Valor"].str.contains("Importe Total",na=False)]["Unnamed: 6"]
                imp_percepcion = "-"
                ####
            elif "20100166144-40" in rutaPDF: ##Agro Industria Santa Maria  #Comprobante de Percepción
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100166144-40","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(114,20,168,576))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(167,21,806,580))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="COMPROBANTE DE PERCEPCIÓN"
                fecha_emision = df_datos[df_datos["Señor(es)"].str.contains("Emisión",na=False)][": TERRANOVA TRADING S.A.C."].to_numpy()[0].split(" ")[1]
                temp_doc = df_costos[df_costos["Unnamed: 0"]=="FACTURA"]["Comprobante de pago que da Origen a la Percepción"].to_numpy()
                doc_vinculado= list(map(lambda x: x.split(" ")[0],temp_doc))
                imp_percepcion = df_costos[df_costos["Unnamed: 6"].str.contains("Percibido",na=False)]["Importe"].values.tolist()[0]
                #imp_total =df_costos[df_costos["Unnamed: 6"].str.contains("Cobrado",na=False)]["Importe"].values.tolist()[0]
                imp_total = "-"
                igv = "-"
                ####
            elif "20100052050-01" in rutaPDF: ##PeruFarma  #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100052050-01","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(147,20,185,809))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(198,514,781,819))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = df_datos.columns.to_numpy()[1]
                doc_vinculado= ["-"]
                imp_percepcion ="-"
                imp_total =df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("Importe Total",na=False)]["Precio Venta"].values.tolist()[0]
                igv = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IGV",na=False)]["Precio Venta"].values.tolist()[0]
                ####
            elif "20100052050-09" in rutaPDF: ##PeruFarma  #Guia de Remision
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100052050-09","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(140,22,216,574))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(342,22,421,574))
                #print(datos_generales_2[0])
                ####
                df_datos= datos_generales[0]
                df_datos_2= datos_generales_2[0]
                #
                tipo_comprobante="GUIA DE REMISIÓN ELECTRÓNICA"
                fecha_emision = df_datos[df_datos["Unnamed: 0"].str.contains("Emisión")]["Unnamed: 0"].to_numpy()[0].split(":")[1]
                temp_doc= df_datos_2[df_datos_2["Unnamed: 0"].str.contains("Factura",na=False)]["Unnamed: 1"].to_numpy()
                doc_vinculado = list(map(lambda x:x.split(":")[1], temp_doc))
                imp_percepcion ="-"
                imp_total ="-"
                igv = "-"
                ####
            elif "20100052050-07" in rutaPDF: ##PeruFarma  #Notas de Credito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100052050-07","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(144,20,172,576))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(432,22,764,382))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(288,516,764,815))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_datos_2= datos_generales_2[0]
                df_costos=costos[0]
                #
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = df_datos.columns.to_numpy()[1]
                temp_doc= df_datos_2[df_datos_2["Tipo Documento Serie y Número"].str.contains("FACTURA",na=False)]["Tipo Documento Serie y Número"].to_numpy()
                doc_vinculado = list(map(lambda x : x.split(" ")[1], temp_doc))
                imp_percepcion ="-"
                imp_total =df_costos[df_costos["Total Valor de Venta - Operaciones Gravadas:"].str.contains("Total a pagar",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                igv = df_costos[df_costos["Total Valor de Venta - Operaciones Gravadas:"].str.contains("IGV",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                ####
            elif "20100190797-01" in rutaPDF: ##Leche Gloria  #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100190797-01","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(142,39,189,554))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(407,401,600,575))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = df_datos.columns.to_numpy()[1].split(" ")[1]
                doc_vinculado= ["-"]
                imp_percepcion =df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("Percepción",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                #imp_total =df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("MontoTotal",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                igv = df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("IGV",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                ####
                # if str(imp_total) == "0.0" or str(imp_total) == "0.00":
                #     imp_total = df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("Importe Total",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                
                imp_total = df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("Importe Total",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                
            elif "20100190797-09" in rutaPDF: ##Leche Gloria  #Guia de Remision
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100190797-09","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(142,39,188,554))
                #print(datos_generales[0])
                ####
                df_datos= datos_generales[0]
                #
                tipo_comprobante="GUIA DE REMISIÓN ELECTRÓNICA"
                fecha_emision = df_datos.columns.to_numpy()[1].split(" ")[1]
                doc_vinculado= df_datos.columns.to_numpy()[7].split(" ")[2]
                imp_percepcion ="-"
                imp_total ="-"
                igv = "-"
                ####
            elif "20100190797-07" in rutaPDF: ##Leche Gloria  #Notas de Crédito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100190797-07","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(154,40,204,374))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(500,179,709,386))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(542,402,800,580))
                #print(costos[0])
                ###
                df_datos = datos_generales[0]
                df_datos_2 = datos_generales_2[0]
                df_costos = costos[0]
                ###
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = df_datos.columns.to_numpy()[1].split(" ")[1]
                
                try:
                    temp_doc= df_datos_2[df_datos_2[df_datos_2.iloc[0].index[1]].str.contains("Factura",na=False)][df_datos_2.iloc[0].index[1]].to_numpy()
                except:
                    temp_doc= df_datos_2[df_datos_2[df_datos_2.iloc[0].index[0]].str.contains("Factura",na=False)][df_datos_2.iloc[0].index[0]].to_numpy()
                        
                doc_vinculado = list(map(lambda x: x.split(" ")[2],temp_doc))
                
                try: 
                    imp_percepcion =df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("Percepción",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                except:
                    imp_percepcion =df_costos[df_costos[df_costos.iloc[0].index[1]].str.contains("Percepción",na=False)][df_costos.iloc[0].index[3]].to_numpy()[0]
                
                # try:
                #     imp_total =df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("MontoTotal",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                # except:
                #     imp_total =df_costos[df_costos[df_costos.iloc[0].index[1]].str.contains("MontoTotal",na=False)][df_costos.iloc[0].index[3]].to_numpy()[0]
                
                try:
                    igv = df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("IGV",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                except:
                    igv = df_costos[df_costos[df_costos.iloc[0].index[1]].str.contains("IGV",na=False)][df_costos.iloc[0].index[3]].to_numpy()[0]
                ####
                #if "0.0" == str(imp_total) or "0.00" == str(imp_total):
                try:
                    imp_total =df_costos[df_costos[df_costos.iloc[0].index[0]].str.contains("Importe Total",na=False)][df_costos.iloc[0].index[2]].to_numpy()[0]
                except:
                    imp_total =df_costos[df_costos[df_costos.iloc[0].index[1]].str.contains("Importe Total",na=False)][df_costos.iloc[0].index[3]].to_numpy()[0]
            
            elif "20100190797-08" in rutaPDF: ##Leche Gloria #Notas de debito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100190797-08","")).replace(".pdf","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(154,39,204,374))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(507,180,709,384))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(547,405,670,569))
                #print(costos[0])
                ###
                df_datos = datos_generales[0]
                df_datos_2 = datos_generales_2[0]
                df_costos = costos[0]
                ###
                tipo_comprobante="NOTA DE DÉBITO"
                fecha_emision = df_datos.columns.to_numpy()[1].split(" ")[1]
                temp_doc= df_datos_2[df_datos_2["Unnamed: 0"].str.contains("Factura",na=False)]["Unnamed: 0"].to_numpy()
                doc_vinculado = list(map(lambda x: x.split(" ")[2],temp_doc))
                imp_percepcion =df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("Percepción",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                #imp_total =df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("MontoTotal",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                igv = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IGV",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                ####
                #if "0.0" == str(imp_total) or "0.00" == str(imp_total):
                imp_total = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("Importe Total",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                
            elif "20100035121_01" in rutaPDF: ##Molitalia  #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100035121_01","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(173,495,244,577))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(575,19,612,583))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = df_datos["Fecha Emisión"][0]
                doc_vinculado= ["-"]
                imp_percepcion = "-"
                imp_total = df_costos["Importe Total"][0]
                igv = df_costos["I.G.V."][0]
                ####
            elif "20100035121_07" in rutaPDF: ##Molitalia #Notas de Crédito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100035121_07","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(195,20,228,576))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(575,19,612,583))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTAS DE CRÉDITO"
                fecha_emision = df_datos["Fecha Emisión"][0]
                doc_vinculado= df_datos["Doc. Referencia"].to_numpy()
                imp_percepcion = "-"
                imp_total = df_costos["Importe Total"][0]
                igv = df_costos["I.G.V."][0]
                ####
            elif "20100152941_08" in rutaPDF: ##Kimberly #Nota de Debito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100152941_08","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(109,362,171,575))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(278,20,322,306))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(640,359,732,576))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_datos_2= datos_generales_2[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTAS DE DÉBITO"
                fecha_emision = df_datos["Fecha de Vencimiento"][2]
                doc_vinculado = []
                doc_vinculado.append(df_datos_2.columns.to_numpy()[2].split(" - ")[1])
                imp_percepcion = "-"
                imp_total = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IMPORTE TOTAL",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                igv = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IGV",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                ####
            elif "20100152941_01" in rutaPDF: ##Kimberly  #Factura
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100152941_01","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(108,359,162,574))
                #print(datos_generales[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(638,369,728,576))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_costos = costos[0]
                #
                tipo_comprobante="FACTURA"
                fecha_emision = df_datos["Fecha de Vencimiento"][2]
                doc_vinculado = ["-"]
                imp_percepcion = "-"
                imp_total = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IMPORTE TOTAL",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                igv = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IGV",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                ####
            elif "20100152941_07" in rutaPDF: ##Kimberly #Notas de Crédito
                folio = str(rutaPDF.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaPDF.split("\\")[-1].replace(f"{folio}_20100152941_07","")).replace(".pdf","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                datos_generales = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(109,360,171,557))
                #print(datos_generales[0])
                datos_generales_2 = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(278,19,325,306))
                #print(datos_generales_2[0])
                costos = tabula.read_pdf(rutaPDF,pages='all',multiple_tables=True,stream=True,area=(643,358,729,580))
                #print(costos[0])
                ####
                df_datos= datos_generales[0]
                df_datos_2= datos_generales_2[0]
                df_costos = costos[0]
                #
                tipo_comprobante="NOTAS DE CRÉDITO"
                fecha_emision = df_datos["Fecha de Vencimiento"][2]
                doc_vinculado = []
                doc_vinculado.append(df_datos_2.columns.to_numpy()[2].split(" - ")[1])
                imp_percepcion = "-"
                imp_total = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IMPORTE TOTAL",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                igv = df_costos[df_costos[df_costos.iloc[1].index[0]].str.contains("IGV",na=False)][df_costos.iloc[1].index[2]].to_numpy()[0]
                ####
            else:
                folio = "-"
                nro_comprobante = ""
                tipo_comprobante="-"
                fecha_emision = "-"
                doc_vinculado =["-"]
                imp_percepcion = "-"
                imp_total = "-"
                igv = "-"
                
            ##### Cargando datos en un diccionario
            pre_data = {}
            pre_data["folio"] = folio
            pre_data["ruc"] = self.proveedor["ruc"]
            pre_data["razon_social"] = self.proveedor["razon_social"]
            pre_data["tipo_comprobante"] = tipo_comprobante
            pre_data["fecha_emision"] = doc_vinculado
            pre_data["fecha_emision"] = fecha_emision
            pre_data["doc_vinculado"] = doc_vinculado
            pre_data["igv"] = igv
            pre_data["imp_total"] = imp_total
            pre_data["imp_percepcion"] = imp_percepcion
            pre_data["comprobante"] = imp_percepcion
            pre_data["nro_comprobante"] = nro_comprobante
            
            ### Cargando datos al lista principal
            data_temp.append(pre_data)
            return data_temp
        except Exception as e:
            print(f"--- ERROR - LECTURA DE PDF: {rutaPDF} ",e)
            return []
        
        
    def getPathReporte(self):
        return self.path_report
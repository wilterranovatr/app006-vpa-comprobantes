from pathlib import Path
import win32com.client
from datetime import datetime, timedelta
import os
import json
import pandas as pd
import openpyxl
import shutil
import multiprocessing
import glob
import time as t
from tkinter import messagebox
import warnings
import xml.etree.ElementTree as ET

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
        output_dir = Path(f'{path_appdata}/Comprobantes Terranova/XML/{provider}/{nom_folder.split("-")[2]}{nom_folder.split("-")[1]}{nom_folder.split("-")[0]}')
        output_dir_pdf = Path(f'{path_appdata}/Comprobantes Terranova/PDF/{provider}/{nom_folder.split("-")[2]}{nom_folder.split("-")[1]}{nom_folder.split("-")[0]}')
        if not output_dir.is_dir():
            os.makedirs(output_dir)
        else:
            shutil.rmtree(output_dir)
            t.sleep(2)
            os.makedirs(output_dir)
            
        if not output_dir_pdf.is_dir():
            os.makedirs(output_dir_pdf)
        else:
            shutil.rmtree(output_dir_pdf)
            t.sleep(2)
            os.makedirs(output_dir_pdf)
        
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
        #messages = messages.Restrict(f"[Subject] = 'COMPROBANTES REENVIADOS | {ruc}'")
        messages = messages.Restrict("[ReceivedTime] >= '" + received_de + "' AND [ReceivedTime] <= '"+received_ha+"'")
        print("--- TOTAL COMPROBANTES ENCONTRADOR",len(messages))
        folio = 1
        for message in messages:
            attachments = message.Attachments
            # Save attachments
            for attachment in attachments:
                attachment_file_name = str(attachment)
                for f in find:
                    folio = folio + 1
                    if attachment_file_name.find(f) != -1 and attachment_file_name.find(".xml") != -1:
                        nom_file_final = str(folio)+"_"+attachment_file_name
                        attachment.SaveAsFile(output_dir / nom_file_final)
                    if attachment_file_name.find(f) != -1 and attachment_file_name.find(".pdf") != -1:
                        nom_file_final = str(folio)+"_"+attachment_file_name
                        attachment.SaveAsFile(output_dir_pdf / nom_file_final)
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
                output_dir = f'{path_appdata}/Comprobantes Terranova/XML/{provider}/{nom_folder.split("-")[2]}{nom_folder.split("-")[1]}{nom_folder.split("-")[0]}'
                files = glob.glob(f"{output_dir}\*.xml")
                #
                self.data.append(pool.map(self.ReadDataXML,files))
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
                sheet[f'H{row_number}'] = d["moneda_actual"]
                sheet[f'I{row_number}'] = "\n".join(d["doc_vinculado"])
                sheet[f'J{row_number}'] = "\n".join(d["fecha_vinculado"])
                sheet[f'K{row_number}'] = d["motivo"]
                sheet[f'L{row_number}'] = d["imp_total"]
                sheet[f'M{row_number}'] = d["descripcion"]
                sheet[f'N{row_number}'] = d["precio_u"]
                sheet[f'O{row_number}'] = d["item_u"]
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
    
    def ReadDataXML(self,rutaxml):
        #
        data_temp = []
        
        # Names spaces
        ns = {
            'sac':'urn:sunat:names:specification:ubl:peru:schema:xsd:SunatAggregateComponents-1',
            'cac':'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
            'cbc':'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
        }
        tree = ET.parse(rutaxml)
        root = tree.getroot()
        ###
        try:
            if "20100055237_01" in rutaxml:   ##Alicorp  #Factura
                #
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100055237_01","")).replace(".xml","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                
                ##
            elif "20100055237_07" in rutaxml: ##Alicorp  #Nota de Credito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = root.find('cbc:ID', ns).text
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = root.find('cbc:IssueDate', ns).text
                moneda_actual = root.find('cac:TaxTotal', ns).find('cbc:TaxAmount', ns).attrib["currencyID"]
                doc_vinculado =[]
                fecha_vinculado=[]
                for item in root.findall('cac:BillingReference', ns):
                    temp = item.find('cac:InvoiceDocumentReference', ns).find('cbc:ID', ns).text
                    temp_date = item.find('cac:InvoiceDocumentReference', ns).find('cbc:IssueDate', ns).text
                    doc_vinculado.append(temp)
                    fecha_vinculado.append(temp_date)
                    
                motivo = root.find('cac:DiscrepancyResponse', ns).find('cbc:Description', ns).text
                imp_total = root.find('cac:LegalMonetaryTotal', ns).find('cbc:PayableAmount', ns).text
                ##
                count = 0
                for item in root.findall('cac:CreditNoteLine', ns):    
                    if count == 1:
                        break
                    else:    
                        descripcion = item.find('cac:Item', ns).findall('cbc:Description', ns)[0].text.split("@@")[0]
                        precio_u = item.find('cac:Item', ns).findall('cbc:Description', ns)[0].text.split("@@")[-1]
                        item_u = "1"
                        count = count +1
                ##
                
            elif "20100055237_08" in rutaxml: ##Alicorp  #Nota de Debito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100055237_08","")).replace(".xml","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                
                ##
            elif "20100055237_040" in rutaxml: ##Alicorp  #Comprobante de Percepción Electrónico
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100055237_040","")).replace(".xml","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                ##
            elif "20207845044-01" in rutaxml: ##Frutos Especies #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20207845044-01","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ##
            elif "20207845044-07" in rutaxml: ##Frutos Especies #Nota de Credito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20207845044-07","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ##
            elif "20207845044-09" in rutaxml: ##Frutos Especies #Guia de Remision Remitente Electrónica
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20207845044-09","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ##
            elif "20205922149-01" in rutaxml: ##P D Andina #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20205922149-01","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20205922149-07" in rutaxml: ##P D Andina #Nota de Credito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20205922149-07","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ###
            elif "20131823020-01" in rutaxml: ##Casa Grande #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20131823020-01","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20131823020-07" in rutaxml: ##Casa Grande #Nota de Credito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20131823020-07","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100166144-01" in rutaxml: ##Agro Industria Santa Maria  #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100166144-01","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100166144-07" in rutaxml: ##Agro Industria Santa Maria  #Nota de Credito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100166144-07","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100166144-40" in rutaxml: ##Agro Industria Santa Maria  #Comprobante de Percepción
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100166144-40","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100052050-01" in rutaxml: ##PeruFarma  #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100052050-01","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100052050-09" in rutaxml: ##PeruFarma  #Guia de Remision
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100052050-09","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100052050-07" in rutaxml: ##PeruFarma  #Notas de Credito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100052050-07","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100190797-01" in rutaxml: ##Leche Gloria  #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100190797-01","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                #
                
            elif "20100190797-09" in rutaxml: ##Leche Gloria  #Guia de Remision
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100190797-09","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                ####
            elif "20100190797-07" in rutaxml: ##Leche Gloria  #Notas de Crédito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100190797-07","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                #
            
            elif "20100190797-08" in rutaxml: ##Leche Gloria #Notas de debito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100190797-08","")).replace(".xml","")
                #
                nro_comprobante = nro_comprobante[1:]
                #
                #
                
            elif "20100035121_01" in rutaxml: ##Molitalia  #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100035121_01","")).replace(".xml","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                ####
            elif "20100035121_07" in rutaxml: ##Molitalia #Notas de Crédito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = root.find('cbc:ID',ns).text
                tipo_comprobante="NOTA DE CRÉDITO"
                fecha_emision = root.find('cbc:IssueDate', ns).text
                moneda_actual = root.find('cac:TaxTotal', ns).find('cbc:TaxAmount', ns).attrib["currencyID"]
                doc_vinculado = []
                temp_doc = root.find('cac:BillingReference', ns).find('cac:InvoiceDocumentReference', ns).find('cbc:ID', ns).text
                doc_vinculado.append(temp_doc)
                fecha_vinculado = []
                temp_date = root.find('cac:BillingReference', ns).find('cac:InvoiceDocumentReference', ns).find('cbc:IssueDate', ns).text
                fecha_vinculado.append(temp_date)
                
                # for item in root.findall('cac:BillingReference', ns):
                #     temp = item.find('cac:InvoiceDocumentReference', ns).findall('cbc:ID', ns).text
                #     temp_date = item.find('cac:InvoiceDocumentReference', ns).findall('cbc:IssueDate', ns).text
                #     doc_vinculado.append(temp)
                #     fecha_vinculado.append(temp_date)
                    
                motivo = ""
                for i in root[0][1][0][0]:
                    if i.attrib["name"] == "MotivoRefe":
                        motivo = i.text
                        
                imp_total = root.find('cac:LegalMonetaryTotal', ns).find('cbc:PayableAmount', ns).text
                ##
                count = 0
                for item in root.findall('cac:CreditNoteLine', ns):    
                    if count == 1:
                        break
                    else:    
                        descripcion = item.find('cac:Item', ns).findall('cbc:Description', ns)[0].text
                        precio_u = item.find('cac:Item', ns).findall('cbc:Description', ns)[2].text
                        item_u = "1"
                        count = count +1
                ##
                ####
            elif "20100152941_08" in rutaxml: ##Kimberly #Nota de Debito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100152941_08","")).replace(".xml","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                ####
            elif "20100152941_01" in rutaxml: ##Kimberly  #Factura
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100152941_01","")).replace(".xml","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                ####
            elif "20100152941_07" in rutaxml: ##Kimberly #Notas de Crédito
                folio = str(rutaxml.split("\\")[-1].split("_")[0])
                nro_comprobante = str(rutaxml.split("\\")[-1].replace(f"{folio}_20100152941_07","")).replace(".xml","")
                #
                temp_nro_comprobante = nro_comprobante.split("_")[1].split("-")
                rest = 8-len(temp_nro_comprobante[1])
                ceros = ''.join(str("0") for val in range(rest))
                nro_comprobante = f"{temp_nro_comprobante[0]}-{ceros}{temp_nro_comprobante[1]}"  
                #
                ####
            else:
                folio = "-"
                nro_comprobante = ""
                tipo_comprobante="-"
                fecha_emision = "-"
                moneda_actual ="-"
                doc_vinculado =["-"]
                fecha_vinculado=["-"]
                motivo ="-"
                imp_total = "-"
                descripcion = "-"
                precio_u = "-"
                item_u = "-"
                
            ##### Cargando datos en un diccionario
            pre_data = {}
            pre_data["folio"] = folio
            pre_data["ruc"] = self.proveedor["ruc"]
            pre_data["razon_social"] = self.proveedor["razon_social"]
            pre_data["tipo_comprobante"] = tipo_comprobante
            pre_data["fecha_emision"] = fecha_emision
            pre_data["doc_vinculado"] = doc_vinculado
            pre_data["moneda_actual"] = moneda_actual
            pre_data["imp_total"] = imp_total
            pre_data["fecha_vinculado"] = fecha_vinculado
            pre_data["motivo"] = motivo
            pre_data["nro_comprobante"] = nro_comprobante
            pre_data["descripcion"] = descripcion
            pre_data["precio_u"] = precio_u
            pre_data["item_u"] = item_u
            ##
            ### Cargando datos al lista principal
            data_temp.append(pre_data)
            return data_temp
        except Exception as e:
            print(f"--- ERROR - LECTURA DE PDF: {rutaxml} ",e)
            return []
        
        
    def getPathReporte(self):
        return self.path_report
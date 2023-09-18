from pathlib import Path
import win32com.client
from datetime import datetime, timedelta


class EmailAttachment:
    
    def __init__(self) -> None:
        pass
    
    def download_attachment(self, email, find, provider):
        # Create output folder
        output_dir = Path.cwd() / "examples" / "outlook" / "attachments" / provider
        output_dir.mkdir(parents=True, exist_ok=True)
        
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
                "correo": "dte_20100055237@paperless.pe"    
            },
            "frutos_especias":{
                "correo": "felectronica@frutosyespecias.com.pe"    
            },
            "p_d_andina_alimentos":{
                "correo": "facturacionelectronica@pdandina.com.pe"    
            },
            "casa_grande":{
                "correo": "de_20131823020@azucarperu.com.pe"    
            },
            "agroindustria_santa_maria":{
                "correo": "comprobantes@efacturacion.pe"    
            },
            "perufarma":{
                "correo": "comprobantes@efacturacion.pe"    
            },
            "leche_gloria":{
                "correo": "de_20100190797@gloria.com.pe"    
            },
            "molitaria":{
                "correo": "dte_20100035121@paperless.pe"    
            },
            "kimberly_clark_peru":{
                "correo": "dte_20100152941@paperless.pe"    
            }
        }
        self.download_attachment(email="dte_20100055237@paperless.pe", find="20100055237_07", provider="alicorp")
import sys
from layouts.__form_menu import FormMenu
from functions.__email_attachment import EmailAttachment
from multiprocessing import freeze_support


def main()->int:
    ##
    freeze_support()
    # ##
    v= FormMenu()
    v.Open()   
    # pathPDF = r"C:\Users\wilfredoca\AppData\Roaming\Comprobantes Terranova\frutos_especias\20231011\90_20207845044-07-F008-5643.pdf"
    # function = EmailAttachment()
    # function.ReadDataPdf(pathPDF)
    #function.descargaxProv("alicorp")
    # ## Probando excel
    # shutil.copy("./assets/model_reporte.xlsx","./reports/prueba.xlsx")
    # data = openpyxl.load_workbook("./reports/prueba.xlsx")
    # sheet = data.active
    # sheet['C5']= datetime.datetime.now()
    # data.save("./reports/prueba.xlsx")
    ##
    return 0

if __name__=='__main__':
    sys.exit(main())
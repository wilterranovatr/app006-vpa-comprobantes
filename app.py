import sys
from layouts.__form_menu import FormMenu
from functions.__email_attachment import EmailAttachment

def main()->int:
    ##
    ##
    # v= FormMenu()
    # v.Open()
    ##
    pathPDF = "C:/Users/wilfredoca/Downloads/pruebas-comprobantes/alicorp/20100055237_01_F601-39051_74377064.pdf"
    function = EmailAttachment()
    function.ReadDataPdf(pathPDF)
    
    return 0

if __name__=='__main__':
    sys.exit(main())
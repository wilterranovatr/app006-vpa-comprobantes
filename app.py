import sys
from layouts.__form_menu import FormMenu

def main()->int:
    ##
    ##
    v= FormMenu()
    v.Open()
    ##
    return 0

if __name__=='__main__':
    sys.exit(main())
import win32com.client as win32
import os
from datetime import datetime
import formato

def Deshabiiltar_error():
        fecha= "{:%Y_%m_%d}".format(datetime.now())
        directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')

        original_file_path = directorio + "\YRA2_TMOBILE_" + fecha + ".xls"
        modified_file_path = directorio + "\YRA2_TMOBILE_" + fecha + ".xlsx"

        # Create an instance of the Excel application object and open the original file
        excel = win32.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(original_file_path)
        excel.DisplayAlerts = False

        # Your code for modifying the workbook goes here...

        # Save the modified workbook as an XLSX file
        wb.SaveAs(modified_file_path, FileFormat=51) 

        # Close the workbook and exit the Excel application
        wb.Close()
        excel.Quit()

        # Delete the original file
        os.remove(original_file_path)

        # Poner formato al Excel
        formato.estilos()
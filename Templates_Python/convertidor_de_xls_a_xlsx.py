import win32com.client as win32
import os
from datetime import datetime


def Deshabiiltar_error():
        
        print("==============================================================================================================")
        print("====INICIALIZACION DE -DESHABILITAR ERROR-")
        print("==============================================================================================================\n")  

        # ----Toma la fecha actual de hoy
        fecha= "{:%Y_%m_%d}".format(datetime.now())

        # ----Directorio de destino
        print(fecha)
        directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')

        # ----Se definen los paths de los archivos, el archivo original y el archivo al que se quiere convertir
        original_file_path = directorio + "\YRA2_TMOBILE_" + fecha + ".xls"
        modified_file_path = directorio + "\YRA2_TMOBILE_" + fecha + ".xlsx"

        print("========================================================================")
        print("----Modificacion de archivo .xls a .xlsx en proceso")
        print("========================================================================\n")

        # ----Create an instance of the Excel application object and open the original file
        excel = win32.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(original_file_path)
        excel.DisplayAlerts = False

        # ----Save the modified workbook as an XLSX file
        wb.SaveAs(modified_file_path, FileFormat=51) 

        # ----Close the workbook and exit the Excel application
        wb.Close()
        excel.Quit()

        print("========================================================================")
        print("----Modificacion terminada y archivo .xlsx guardado")
        print("Archivo xlsx = " + modified_file_path)
        print("========================================================================\n")

        # ----Delete the original file
        os.remove(original_file_path)

        print("========================================================================")
        print("----Archivo .xls antiguo eliminado")
        print("Archivo xls = " + original_file_path)
        print("========================================================================\n")

        print("==============================================================================================================")
        print("====FINALIZACION DE -DESHABILITAR ERROR-")
        print("==============================================================================================================\n") 

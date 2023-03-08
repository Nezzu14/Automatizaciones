import win32com.client as win32
import os
from datetime import datetime
import openpyxl


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

        #--------------------------------------------------------------------------------------------------------------------
        # <<<<<<<<<SE EJECUTA LA CORRECCION DENTRO DE LOS ARCHIVOS DE REPORTE YRA2 Y DEL ARCHIVO GIC
        correciones_lineas_xlsx(fecha)
        #--------------------------------------------------------------------------------------------------------------------


def correciones_lineas_xlsx(fecha):

        print("==============================================================================================================")
        print("====INICIALIZACION DE -CORRECIONES LINEAS .XLSX-")
        print("==============================================================================================================\n")  

        # ----Se definen los paths de los archivos, el archivo .xlsx
        xlsx_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2') + "\YRA2_TMOBILE_" + fecha + ".xlsx"

        # ----Cargar el archivo Excel
        workbook = openpyxl.load_workbook(xlsx_file_path)

        # ----Seleccionar la hoja que deseas modificar
        sheet = workbook["YRA2_TMOBILE_" + fecha]

        # ----Modificar el valor de una celda
        sheet['AK3'] = 'GIC'

        # ----Eliminar una fila de datos de la hoja de trabajo
        sheet.delete_rows(4, 1) # Elimina la fila 2
        sheet.delete_rows(2, 1) # Elimina la fila 2
        sheet.delete_rows(1, 1) # Elimina la fila 1

        # ----Guardar los cambios en el archivo
        workbook.save(xlsx_file_path)

        print("========================================================================")
        print("Archivo xlsx = " + xlsx_file_path)
        print("----Nombre de la Celda _AK3_ se cambio de -GIC code- a -GIC-")
        print("----Se elimino las filas 1 y 2")
        print("========================================================================\n")

        print("==============================================================================================================")
        print("====FINALIZACION DE -CORRECIONES LINEAS .XLSX-")
        print("==============================================================================================================\n")  



#       """""""""En dado caso que quiera ejecutarlo aca en el archivo:""""""""""
#Deshabiiltar_error()
#correciones_lineas_xlsx()
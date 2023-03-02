import pandas as pd
import clasificador
import os

def estilos(df):
         directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"/DATA"
         writer= pd.ExcelWriter(directorio+"\\"+"CLASIFICADOS"+".xlsx", engine="xlsxwriter")
         
         df.to_excel(writer, index= False)
         
         workbook= writer.book
         worksheet= writer.sheets["Sheet1"]
         
         worksheet.set_zoom(90)
         
         header_format = workbook.add_format(
             {   
                
                 "bold":True,
                 "bg_color": "#8dbdeb"
         
         
             }
         )
         
         for col_num, value in enumerate(df.columns.values):
             worksheet.write(0, col_num, value, header_format)
         
         
         worksheet.set_column("A:Z", 38)
         writer.save()
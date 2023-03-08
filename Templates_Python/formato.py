import pandas as pd
import os
from datetime import datetime

def estilos(df):
         fecha= "{:%Y_%m_%d}".format(datetime.now())
         directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')
         writer= pd.ExcelWriter(directorio+"\YRA2_TMOBILE_" + fecha + ".xlsx", engine="xlsxwriter")
         
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
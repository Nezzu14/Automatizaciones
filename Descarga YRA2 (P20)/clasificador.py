import pandas as pd
import numpy as np
import formato 
import os
import re


def clasificacion(df):
    
    directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'YRA2')

    old_column_names = df.columns
    
    # Crear un diccionario para los nuevos nombres de las columnas
    new_column_names = {}
    
    # Iterar sobre los nombres de las columnas y limpiarlos
    for column_name in old_column_names:
         clean_name = re.sub(r'\W+', ' ', column_name).strip()
         new_column_names[column_name] = clean_name
     
     
     # Reemplazar los nombres de las columnas antiguas por los nuevos
    df = df.rename(columns=new_column_names)
 
    #df["TOP WBS"]= df["WBS Element"].str[:11]
    df["WBS Element"]=df["WBS Element"].str.strip()
    df["WBS Type"]= df["WBS Element"].str[8:10]#corregir
    df['Ord Num'] = df['Ord Num'].apply(lambda x: re.sub(r'\D', '', x))# ELIMINAR CUALQUIR CARACTER NO NUMERICO
    df['Ord Num']=df['Ord Num'].astype(int)#CONVERTIR EN ENTERO LA COLUMNA
    
    condiciones= [
    
    #SVO´s
    (df['Type']=="SVO") & (df['Status'] == 'OPEN'),#SVO if is OPEN Ready to close?, if it's not, justify
    (df['Type']=="SVO") & (df['Status'] == 'CLOSED'),#its ok 
    
    #SO´s 116-117
    (df['Type']=="SO") & (df['Status'] == 'OPEN') & (df['Ord Num']>=1160000000) &  (df['Ord Num']<= 1170000000),#KU's LOG must looks into it for its closing.
    (df['Type']=="SO") & (df['Status'] == 'CLOSED') & (df['Ord Num']>=1160000000) &  (df['Ord Num']<= 1170000000),#It's ok
    
    #SO´s PZ
    (df['Type']=="SO") & (df['Status'] == 'OPEN') & (df['WBS Type']=='PZ') ,#Log Team must looks into it because it's not closed.
    (df['Type']=="SO") & (df['Status'] == 'CLOSED') & (df['WBS Type']=='PZ'),#It's ok
    #SO´s EZ
    (df['Type']=="SO") & (df['Status'] == 'OPEN') & (df['WBS Type']=='EZ'),#Log Team must looks into it because it's not closed.
    (df['Type']=="SO") & (df['Status'] == 'CLOSED') & (df['WBS Type']=='EZ'),#It's ok
    
    #SO´s PD Open
    (df['Type']=="SO") & (df['Status'] == 'OPEN') & (df['WBS Type']=='PD')  & (df['RR Trigger']=="X" ),#Ok, pending on invoice
    (df['Type']=="SO") & (df['Status'] == 'OPEN') & (df['WBS Type']=='PD')  & (df['RR Trigger']!="X" ) & (df['Ord Num']<1129000000),# FPC must looks into it.
    # que pasa con las que son 1129 exacta?
    (df['Type']=="SO") & (df['Status'] == 'OPEN') & (df['WBS Type']=='PD')  & (df['RR Trigger']!="X" ) & (df['Ord Num']>1129000000),#Ok, on going.
    
    
    #SO´s PD Open Rec
    (df['Type']=="SO") & (df['Status'] == 'OPEN RECVBL') & (df['WBS Type']=='PD')  & (df[ 'RR Trigger']=="X" ),#Ok, pending on payment
    
    #SO´s PD Close
    (df['Type']=="SO") & (df['Status'] == 'CLOSED') & (df['WBS Type']=='PD')  & (df['RR Trigger']=="X" ),#It's ok
    (df['Type']=="SO") & (df['Status'] == 'CLOSED') & (df['WBS Type']=='PD')  & (df['RR Trigger']!="X" ),#Log must tell us if: the SO has value, if it's totally cancelled and why they didn't check the RRT
    
    
    
    ]
    
    opciones=[
        
    "Ready to close?, if it's not, justify",#SVOS Open
    
    "It's ok",#SVOS Close
    
    "KU's LOG must looks into it for its closing.",# SO 116-117 OPEN 
    
    "It's ok",# SO 116-117 CLOSE
    
    "Log Team must looks into it because it's not closed.",# PZ OPEN
    
    "It's ok",#PZ CLOSE
    
    "Log Team must looks into it because it's not closed.",#EZ OPEN
    
    "It's ok",#EZ CLOSE
    
    "Ok, pending on invoice",#SO PD OPEN CON RRT
    
    "FPC must looks into it.",#SO OPEN SIN RRT MENOR A 1129
    
    "Ok, on going.",#OPEN SIN RRT MAYOR A 1129
    
    "Ok, pending on payment",#SO PD OPENRCVB CON RRT
    
    "It's ok",# SO PD CLOSE CON RTT
    
    "Log must tell us if: the SO has value, if it's totally cancelled and why they didn't check the RRT"#SO PD CLOSE SIN RRT
    
    
    ]
    
    # df["Comments"]=np.where((df['Type']=="SVO") & (df['Status'] == 'CLOSED'),"It","No")
    df["Comments"]=np.select(condiciones,opciones)
    
    
    df.drop(['WBS Type'],axis= 1)
  
    
    df.to_excel(directorio+"\\"+"CLASIFICADOS"+".xlsx", index=False)
    
    formato.estilos(df)
    
   # formato.formateador(df)



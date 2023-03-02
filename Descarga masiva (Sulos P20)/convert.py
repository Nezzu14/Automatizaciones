import pandas as pd
import os
import clasificador

def transformar_fichero():
     print("========================== ejecutando trasnformar")
     directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"//DATA"
     contenido = os.listdir(directorio)

     for fichero in contenido:
        if fichero[-4:]==".htm"or fichero[-4:]==".html" :
           df= pd.read_html(directorio+"\\"+fichero)[0]
           df.columns = df.iloc[0]
           df= df.drop([0])
           df.to_excel(directorio+"\\"+fichero[:-4]+".xlsx", index=False)


def unir():
    print("========================== ejecutando unir")

    directorio = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+"/DATA"
    contenido = os.listdir(directorio)

    df_completo = pd.DataFrame()
   
    for fichero in contenido:
        print(fichero[-5:])
        if fichero[-5:]==".xlsx":
           df= pd.read_excel(directorio+"\\"+fichero)
           df_completo = df_completo.append(df, ignore_index=True)
           print(df_completo)
    clasificador.clasificacion(df_completo)

  #  df_completo.to_excel(directorio+"\\"+"TERMINADO"+".xlsx")


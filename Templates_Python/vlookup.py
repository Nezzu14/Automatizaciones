import pandas as pd

#define first DataFrame
df1 = pd.read_excel("df1.xlsx")

#define second DataFrame
df2 = pd.read_excel("df2.xlsx")


vlookup_df = pd.merge(df1,
                     df2[['Jugador', 'Trabajo']],
                     on ='Jugador',
                     how ='left')

#view df1
print(vlookup_df)
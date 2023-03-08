import pandas as pd

#create DataFrame
df1 = pd.DataFrame({'year': [2015, 2016, 2017, 2018, 2019, 2020, 2021],
                    'sales': [500, 534, 564, 671, 700, 840, 810]})

df2 = pd.DataFrame({'year': ['2015', '2016', '2017', '2018', '2019', '2020', '2021'],
                    'refunds': [31, 36, 40, 40, 43, 70, 62]})

#attempt to merge two DataFrames
big_df = df1.merge(df2, on='year', how='left')

#view DataFrames
print(df1)
import pandas  as  pd

df = pd.read_excel('info_table.xlsx',skiprows=1)

record_dict = df.loc[2].to_dict()
print(record_dict)
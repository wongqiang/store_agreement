from docxtpl import DocxTemplate
from pathlib import Path
import pandas as pd

current_dir = Path.cwd()
doc = DocxTemplate('template.docx')

df = pd.read_excel('info_table.xlsx',skiprows=1)
for row in range(1,df.shape[0]):
# print(df.shape[0])
    record_dict = df.loc[row].to_dict()
    store_name = record_dict.get('store_name')
    employee_name = record_dict.get('employee_name')

    path = current_dir.joinpath(store_name)
    if not path.exists():
        path.mkdir()

    doc.render(record_dict)
    # doc.save(f'./{store_name}/{store_name}-{employee_name}.docx')
    doc.save(f'./Print/{store_name}-{employee_name}.docx')
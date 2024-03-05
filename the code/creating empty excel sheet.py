import pandas as pd
import  numpy as np
import openpyxl
import xlsxwriter

path=""
ind=['Μετοχή','Μέση ημερήσια απόδοση(%)','Τυπική Απόκλιση','Ελάχιστη ημερήσια απόδοση(%)','Μέγιστη ημερήσια απόδοση(%)','Συντελεστής β']
data=pd.DataFrame(np.array(ind).reshape((6,1)))
print(data)
writer = pd.ExcelWriter(path + 'Αρχείο Μετοχών.xlsx', engine='xlsxwriter')
data.to_excel(writer, index=False,header=False, sheet_name='Sheet1')
workbook = writer.book
worksheet = writer.sheets['Sheet1']
for i, col in enumerate(data.columns):
    width = max(data[col].apply(lambda x: len(str(x))).max(), len(data[col]))
    worksheet.set_column(i, i, width)
writer.close()
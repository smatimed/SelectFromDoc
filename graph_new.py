import pandas as pd
from pandasql import sqldf
import matplotlib.pyplot as plt
import openpyxl
import numpy as np


def my_read_excel(excelDocument:str):
    """Replace pandas read_excel which we cannot make it read a text column in Excel which is like a number as text.
    """
    wb = openpyxl.load_workbook(excelDocument)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        row_data = []
        for cell in row:
            row_data.append(cell)
        data.append(row_data)
    df_result = pd.DataFrame(data[1:], columns=data[0])
    return df_result


df = my_read_excel('ass.xlsx')



# * ---------------------------------------------------------- 1 BAR
dfs = sqldf('select categorie, count(*) as nbre from df group by categorie', globals())

# plt.figure()
plt.figure(figsize=(10, 6))

plt.bar(dfs['CATEGORIE'], dfs['nbre'], label='Nombre', width=0.2)

plt.xlabel('Catégorie')
plt.ylabel('Nombre')
plt.title('Catégorie - Nombre')

plt.legend()

# Rotate x-axis labels for better readability if needed
plt.xticks(rotation=45)   # ?

plt.tight_layout()
plt.show()


# * ---------------------------------------------------------- 2 BARs
dfs = sqldf('select categorie, count(*)*10000 as nbre, sum(salaire_ref) as montant from df group by categorie', globals())

# plt.figure()
plt.figure(figsize=(10, 6))

# Define the width of each bar
bar_width = 0.2  ###

x_positions = np.arange(len(dfs['CATEGORIE']))  # Create equally spaced x positions


plt.bar(x_positions, dfs['nbre'], width=bar_width, label='Nombre')
plt.bar(x_positions + bar_width, dfs['montant'], width=bar_width, label='Montant')

plt.xlabel('Catégorie')
plt.ylabel('Values')
plt.xticks(x_positions + bar_width, dfs['CATEGORIE'])  # Set x-axis labels back to the original non-numeric values
plt.title('Catégorie - Nombre & Montant')
plt.legend()

# Rotate x-axis labels for better readability if needed
plt.xticks(rotation=45)

plt.tight_layout()   # ?
plt.show()

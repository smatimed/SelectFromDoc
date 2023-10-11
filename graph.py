import pandas as pd
from pandasql import sqldf

import matplotlib.pyplot as plt
import matplotlib as mpl
import numpy as np


df = pd.read_excel('FACT.xlsx')

# Count the number of occurrences of each category in 'Sexe'
vCounts = df['NAT'].value_counts()

# Create a bar plot
vCounts.plot(kind='bar')

plt.xlabel('Nature')
plt.ylabel('Count')
plt.title('Distribution of NAT')

plt.show()



# --------------------------------

import matplotlib.pyplot as plt

# Your DataFrame
data = {
    'NAT': ['A70', 'R06', 'R07', 'R08', 'R11', 'R14', 'R15', 'R16', 'R18'],
    'debit': [
        5.553207e+06, 5.472055e+06, 1.962925e+06, 1.474829e+07, 1.440000e+04,
        1.112436e+06, 1.759864e+06, 1.656519e+07, 1.281052e+06
    ],
    'credit': [
        5.545404e+06, 1.396014e+06, 5.885791e+05, 3.881803e+06, 7.200000e+03,
        0.000000e+00, 0.000000e+00, 0.000000e+00, 1.281052e+06
    ],
    'solde': [
        7.803450e+03, 4.076041e+06, 1.374346e+06, 1.086649e+07, 7.200000e+03,
        1.112436e+06, 1.759864e+06, 1.656519e+07, 0.000000e+00
    ]
}

import pandas as pd
df = pd.DataFrame(data)

# Set the figure size
plt.figure(figsize=(10, 6))

# Plot the data
plt.bar(df['NAT'], df['debit'], label='Debit', width=0.2)
plt.bar(df['NAT'], df['credit'], label='Credit', width=0.2, bottom=df['debit'])
plt.bar(df['NAT'], df['solde'], label='Solde', width=0.2, bottom=df['debit'] + df['credit'])

# Add labels and title
plt.xlabel('NAT')
plt.ylabel('Amount')
plt.title('Debit, Credit, and Solde by NAT')

# Add a legend
plt.legend()

# Rotate x-axis labels for better readability if needed
plt.xticks(rotation=45)

# Show the plot
plt.tight_layout()
plt.show()



# --------------------------------------

import matplotlib.pyplot as plt
import numpy as np

# Your DataFrame
data = {
    'NAT': ['A70', 'R06', 'R07', 'R08', 'R11', 'R14', 'R15', 'R16', 'R18'],
    'debit': [
        5.553207e+06, 5.472055e+06, 1.962925e+06, 1.474829e+07, 1.440000e+04,
        1.112436e+06, 1.759864e+06, 1.656519e+07, 1.281052e+06
    ],
    'credit': [
        5.545404e+06, 1.396014e+06, 5.885791e+05, 3.881803e+06, 7.200000e+03,
        0.000000e+00, 0.000000e+00, 0.000000e+00, 1.281052e+06
    ],
    'solde': [
        7.803450e+03, 4.076041e+06, 1.374346e+06, 1.086649e+07, 7.200000e+03,
        1.112436e+06, 1.759864e+06, 1.656519e+07, 0.000000e+00
    ]
}

import pandas as pd
df = pd.DataFrame(data)

# Set the figure size
plt.figure(figsize=(10, 6))

# Define the width of each bar
bar_width = 0.2

# Create x-axis positions for each group of bars
x = np.arange(len(df['NAT']))

# Plot the 'debit' bars
plt.bar(x - bar_width, df['debit'], width=bar_width, label='Debit')

# Plot the 'credit' bars
plt.bar(x, df['credit'], width=bar_width, label='Credit')

# Plot the 'solde' bars
plt.bar(x + bar_width, df['solde'], width=bar_width, label='Solde')

# Add labels and title
plt.xlabel('NAT')
plt.ylabel('Amount')
plt.title('Debit, Credit, and Solde by NAT')

# Add x-axis labels and rotate them for better readability
plt.xticks(x, df['NAT'])
plt.xticks(rotation=45)

# Add a legend
plt.legend()

# Show the plot
plt.tight_layout()
plt.show()

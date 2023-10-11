import pandas as pd
from pandasql import sqldf
import matplotlib.pyplot as plt

d = pd.read_excel('test_file.xlsx')

d2 = sqldf('select sexe, count(*) as nombre, sum(salaire) as montant from d group by sexe', globals())


# ------- 1- Bar Plot for Categorical Data (e.g., 'Sexe' column):
# Count the number of occurrences of each category in 'Sexe'
# sexe_counts = d2['Sexe'].value_counts()
sexe_counts = d['Sexe'].value_counts()

# Create a bar plot
sexe_counts.plot(kind='bar')

plt.xlabel('Sexe')
plt.ylabel('Count')
plt.title('Distribution of Sexe')

plt.show()


# ------- 2- Histogram for Numerical Data (e.g., 'montant' column):
# Create a histogram
plt.hist(d2['montant'], bins=10)  # You can adjust the number of bins

plt.xlabel('Montant')
plt.ylabel('Frequency')
plt.title('Histogram of Montant')

plt.show()


# ------- 3- Scatter Plot for Numerical Data (e.g., 'nombre' vs. 'montant'):
import seaborn as sns

# Create a scatter plot
sns.scatterplot(x='nombre', y='montant', data=d2)

plt.xlabel('Nombre')
plt.ylabel('Montant')
plt.title('Scatter Plot: Nombre vs. Montant')

plt.show()


# ! Make sure you have Matplotlib and Seaborn installed in your Python environment using pip install matplotlib seaborn if you haven't already.
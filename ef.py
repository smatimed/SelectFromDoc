import matplotlib.pyplot as plt

# Example data
data = [45, 30, 15, 10]

# Labels for the sections
labels = ['A', 'B', 'C', 'D']

# Creating explode data
explode = (0.1, 0, 0, 0)  # 'explode' the 1st slice

# Plotting a pie chart
plt.pie(data, explode=explode, labels=labels, autopct='%1.1f%%', shadow=True, startangle=90)

# Aspect ratio to ensure that pie is drawn as a circle
plt.axis('equal')

# Display the pie chart
plt.show()

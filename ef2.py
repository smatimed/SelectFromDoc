import matplotlib.pyplot as plt

# Sample data
x = [1, 2, 3, 4, 5]
y1 = [1, 4, 9, 16, 25]
y2 = [1, 8, 27, 64, 125]

# Plotting the data
plt.plot(x, y1, label='y = x^2')
plt.plot(x, y2, label='y = x^3')

# Adding legend
#plt.legend(loc='upper right', bbox_to_anchor=(1.2, 1))
plt.legend(bbox_to_anchor=(1.2, 1))

# Display the plot
plt.show()

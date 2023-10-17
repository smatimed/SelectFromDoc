import tkinter as tk

def show_context_menu(event):
    context_menu.post(event.x_root, event.y_root)

root = tk.Tk()

# Create a context menu
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Cut")
context_menu.add_command(label="Copy")
context_menu.add_command(label="Paste")

# Create a button
button = tk.Button(root, text="Right-click me")

# Bind the context menu to the button
button.bind("<Button-3>", show_context_menu)  # "<Button-3>" represents the right mouse button click

# Pack the button
button.pack()

# Run the main loop
root.mainloop()

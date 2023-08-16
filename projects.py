
import win32com.client
import tkinter as tk
from tkinter import font

# Create an Outlook application object
outlook = win32com.client.Dispatch("Outlook.Application")

# Get the namespace object
namespace = outlook.GetNamespace("MAPI")

# Get the Tasks folder object
tasks_folder = namespace.GetDefaultFolder(13)

# Specify the category you want to get
category = "Projects"

# Create a list of tasks that have the specified category
category_tasks = []
for task in tasks_folder.Items:
    if task.Categories == category and task.Status != 2:
        category_tasks.append(task)

# Sort the list of tasks by their importance (descending order)
category_tasks = sorted(category_tasks, key=lambda x: x.Importance, reverse=True)

# Create a tkinter window
window = tk.Tk()
window.title("Projects")

# Create a label to show the number of tasks
label = tk.Label(window, text=f"You have {len(category_tasks)} tasks in the category {category}:")
label.pack()

# Create a listbox to show the tasks
listbox = tk.Listbox(window)
list_font = font.Font(family='Helvetica',size=10, weight='normal')
listbox.config(width=200, height=50, font=list_font)
listbox.pack()
# Loop through the tasks and insert them into the listbox
for task in category_tasks:
    # Format the task information
    task_info = f"{task.Subject}"
    listbox.insert("end", task_info) # Insert the item first
    listbox.itemconfig("end", foreground = "red" if task.Importance==2 else "black")

# Start the main loop of the window
window.mainloop()

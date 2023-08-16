import os
from tkinter import END
import win32com.client as win32
import tkinter as tk
from datetime import datetime, timedelta

def create_task():
    outlook = win32.Dispatch('Outlook.Application')
    tasks = task_text.get('1.0', 'end').split('\n')
    for task in tasks:
        if task:
            new_task = outlook.CreateItem(3)
            # Split the task subject by "(" and ")"
            parts = task.split("(")
            # The first part is the subject without the parenthesis
            subject = parts[0].strip()
            new_task.Subject = subject
            # If there are more than one part, it means there is something inside the parenthesis
            if len(parts) > 1:
                # The second part is the body inside the parenthesis, without the closing ")"
                body = parts[1].strip(")")
                new_task.Body = body
            new_task.DueDate = datetime.now() + timedelta(days=1)
            new_task.ReminderSet = True
            new_task.ReminderTime = datetime.now() + timedelta(days=2)
            # Check if the task subject ends with "!"
            if subject.endswith("!"):
                # Set the task priority to high (2)
                new_task.Importance = 2
                new_task.Subject = subject[:-1]
                
            if subject.endswith("?"): 
                new_task.Categories = "Agenda" 
                
            if subject.endswith("*"): 
                new_task.Categories = "Projects" 
                new_task.Subject = subject[:-1]
            
            new_task.Save()

def save_tasks(event):
    create_task()
    #os.startfile("outlook")
    task_text.delete('1.0', END)
    # Update the statusbar with the current time
    statusbar.config(text=f"Tasks saved at {datetime.now().strftime('%H:%M:%S')}")

def close(event):
    root.quit()

root = tk.Tk()
root.geometry('300x250')
root.title('Inbox CTRL+S')

task_text = tk.Text(root, height=5)
task_text.pack(pady=10)
task_text.focus_set()
# Create a label for instructions under the textbox
instructions = tk.Label(root, text="To add an important task, end the subject with '!'\nTo add a a project item use * in the end\nTo add a an Agenda item use?\nTo add a body to the task, use parenthesis '()'\nTo save the tasks, press CTRL+S\n To quit press CTRL+Q")
instructions.pack(pady=10)

# Create a statusbar at the bottom of the window
statusbar = tk.Label(root, text="No tasks saved yet", relief=tk.SUNKEN, anchor=tk.W)
statusbar.pack(side=tk.BOTTOM, fill=tk.X)

root.bind('<Control-s>', save_tasks)
root.bind('<Control-q>', close)
root.mainloop()

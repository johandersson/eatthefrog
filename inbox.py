from tkinter import END
import win32com.client as win32
import tkinter as tk
from datetime import datetime, timedelta


class Inbox(tk.Toplevel): # Inherit from Toplevel instead of Tk
    def __init__(self, master, load_tasks):
        self.load_tasks = load_tasks
        super().__init__(master) # Call the super class constructor with the master window
        self.geometry('600x600')
        self.title('Add several tasks, one on each line')

        self.task_text = tk.Text(self, height=20) # Use self instead of self.root
        self.task_text.pack(pady=10)
        self.task_text.focus_set()
        # Create a label for instructions under the textbox
        self.instructions = tk.Label(self,
                                     text="To add an important task, end the subject with '!'\nTo add a bug end subject with !!\n To add a a project item use * in the end\nTo add a an Agenda item use?\nTo add a body to the task, use parenthesis '()'\nTo save the tasks, press CTRL+S\n To quit press CTRL+Q")
        self.instructions.pack(pady=10)

        # Create a statusbar at the bottom of the window
        self.statusbar = tk.Label(self, text="No tasks saved yet", relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.bind('<Control-s>', self.save_tasks)
        self.bind('<Control-q>', self.close)

        #add save button to save tasks
        self.save_button = tk.Button(self, text="Save tasks", command=self.save_tasks)
        self.save_button.pack(side=tk.RIGHT, padx=10, pady=10)

    def set_focus(self):
        self.task_text.focus_set()

    def create_task(self):
        outlook = win32.Dispatch('Outlook.Application')
        tasks = self.task_text.get('1.0', 'end').split('\n')
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

                if subject.endswith("!") or subject.endswith("A"):
                    # Set the task priority to high (2)
                    new_task.Importance = 2
                    new_task.Subject = subject[:-1]
                    new_task.Categories = "A"

                if subject.endswith("?"):
                    new_task.Categories = "Agenda"

                if subject.endswith("*"):
                    new_task.Categories
                    new_task.Subject = subject[:-1]

                if subject.endswith("!!"):
                    new_task.Categories = "Bugs"
                    new_task.Subject = subject[:-1]

                if subject.endswith("B"):
                    new_task.Categories = "B"
                    new_task.Subject = subject[:-1]

                if subject.endswith("C"):
                    new_task.Categories = "C"
                    new_task.Subject = subject[:-1]

                new_task.Save()

    def save_tasks(self):
        self.create_task()
        # os.startfile("outlook")
        self.task_text.delete('1.0', END)
        # Update the statusbar with the current time
        self.statusbar.config(text=f"Tasks saved at {datetime.now().strftime('%H:%M:%S')}")
        self.load_tasks()
        #close the window
        self.destroy()

    def close(self, event):
        self.destroy() # Use destroy instead of quit to close the popup window

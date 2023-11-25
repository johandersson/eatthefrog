import os
from tkinter import END
import win32com.client as win32
import tkinter as tk
from datetime import datetime, timedelta


class Inbox:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry('600x600')
        self.root.title('Add several tasks, one on each line')

        self.task_text = tk.Text(self.root, height=20)
        self.task_text.pack(pady=10)
        self.task_text.focus_set()
        # Create a label for instructions under the textbox
        self.instructions = tk.Label(self.root,
                                     text="To add an important task, end the subject with '!'\nTo add a bug end subject with !!\n To add a a project item use * in the end\nTo add a an Agenda item use?\nTo add a body to the task, use parenthesis '()'\nTo save the tasks, press CTRL+S\n To quit press CTRL+Q")
        self.instructions.pack(pady=10)

        # Create a statusbar at the bottom of the window
        self.statusbar = tk.Label(self.root, text="No tasks saved yet", relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.root.bind('<Control-s>', self.save_tasks)
        self.root.bind('<Control-q>', self.close)

        #add save button to the bottom right
        self.save_button = tk.Button(self.root, text="Save (CTRL+S)", command=self.create_task)
        self.save_button.pack(side=tk.RIGHT, padx=10, pady=10)


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

    def save_tasks(self, event):
        self.create_task()
        # os.startfile("outlook")
        self.task_text.delete('1.0', END)
        # Update the statusbar with the current time
        self.statusbar.config(text=f"Tasks saved at {datetime.now().strftime('%H:%M:%S')}")

    def close(self, event):
        self.root.quit()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    inbox = Inbox()
    inbox.run()

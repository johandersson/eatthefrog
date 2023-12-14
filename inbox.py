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

        #add radiobuttons to select the priority of the task
        self.priority = tk.StringVar()
        self.priority.set("A")
        self.radiobuttons = tk.Frame(self)
        self.radiobuttons.pack()
        #add label
        self.label = tk.Label(self.radiobuttons, text="Priority:")
        self.high = tk.Radiobutton(self.radiobuttons, text="A", variable=self.priority, value="A")
        self.high.pack(side=tk.LEFT)
        self.medium = tk.Radiobutton(self.radiobuttons, text="B", variable=self.priority, value="B")
        self.medium.pack(side=tk.LEFT)
        self.low = tk.Radiobutton(self.radiobuttons, text="C", variable=self.priority, value="C")
        self.low.pack(side=tk.LEFT)
        #project
        self.project = tk.Radiobutton(self.radiobuttons, text="No category", variable=self.priority, value="No category")
        self.project.pack(side=tk.LEFT)
       

        self.task_text = tk.Text(self, height=20) # Use self instead of self.root
        self.task_text.pack(pady=10)
        self.task_text.focus_set()
        # Create a label for instructions under the textbox

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

                #check radiobuttons for priority and set it as a category
                if self.priority.get() == "A":
                    new_task.Categories = "A"
                elif self.priority.get() == "B":
                    new_task.Categories = "B"
                elif self.priority.get() == "C":
                    new_task.Categories = "C"
                elif self.priority.get() == "No category":
                    new_task.Categories = ""
       
                new_task.Save()

    def save_tasks(self, event=None):
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

#import datetime with today
import datetime

import win32com.client
import tkinter as tk  # you can use tkinter or another library to create a GUI

# set some constants for the drawing
dot_radius = 8  # radius of each dot
dot_gap = 10  # gap between each dot and task
text_gap = 20  # gap between each task and text
line_height = 30  # height of each line
root = tk.Tk()


def close_window(event=None):
    # destroy the root window
    root.destroy()


# bind escape key to close the window


root.bind("<Escape>", close_window)

# make the window full screen from start
# root.attributes('-fullscreen', True)
root.config(bg="white")
#root title "Eat the frog"
root.title("Eat the frog")
# create a frame to display the tasks
frame = tk.Frame(root)
frame.config(width=800, height=600)
frame.pack(fill=tk.BOTH)

# create a canvas to draw the dotted list
canvas = tk.Canvas(frame)
canvas.config(bg="white")
canvas.config(width=700, height=550)
canvas.pack(fill=tk.BOTH)

#add status bar to root
status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)
def draw_tasks():
    # clear
    canvas.delete("all")
    if len(tasks_by_category) == 0:
        # draw a big light green check mark on the canvas
        canvas.create_text(350, 200, text="✓", fill="green", font=("Arial", 100), anchor=tk.CENTER)
        # draw a text under the check mark saying "No tasks"
        canvas.create_text(350, 300, text="No tasks", fill="green", font=("Arial", 30), anchor=tk.CENTER)
        # draw a text under the check mark saying "Press ctrl+n to create a new task"
        canvas.create_text(350, 350, text="Press ctrl+n to create a new task", fill="green", font=("Arial", 15),
                           anchor=tk.CENTER)
        # draw a text under the check mark saying "Press ctrl+r to reload the tasks"
        canvas.create_text(350, 380, text="Press ctrl+r to reload the tasks", fill="green", font=("Arial", 15),
                           anchor=tk.CENTER)
        #press ctrl+p to show project tasks
        canvas.create_text(350, 410, text="Press ctrl+p to show project tasks", fill="green", font=("Arial", 15),
                           anchor=tk.CENTER)
        
    global task, category
    for i, (task, category) in enumerate(tasks_by_category):
        # get the task subject and due date
        subject = task.Subject
        due_date = task.DueDate
        body = task.Body

        # format the task information as a string
        task_info = f"{subject}"

        # calculate the coordinates for drawing
        x1 = dot_radius + 20  # x coordinate of the dot center
        y1 = (i + 1) * line_height  # y coordinate of the dot center
        x2 = x1 + dot_gap  # x coordinate of the task start
        y2 = y1  # y coordinate of the task start
        x3 = x2 + text_gap  # x coordinate of the text start
        y3 = y2  # y coordinate of the text start

        # set the color of the dot according to the category


        if category == "A":
            color = "red"
        elif category == "B":
            color = "yellow"
        elif category == "C":
            color = "green"
        else:
            color = "gold"

        # draw a dot on the canvas with the color
        dot = canvas.create_oval(x1 - dot_radius, y1 - dot_radius, x1 + dot_radius, y1 + dot_radius, fill=color,

                                 outline=color)
        canvas.tag_bind(dot, "<Button-3>", lambda event, task=task: mark_done(event, task))
        # when double clicking the dot open the task in outlook
        canvas.tag_bind(dot, "<Double-Button-1>", lambda event, task=task: task.Display())
        # draw a text with the task information on the canvas with black color and Arial font size 14
        task_text = canvas.create_text(x3, y3, text=task_info, fill="black", font=("Arial", 12), anchor=tk.W)
        if len(body) > 100:
            #strip the body of new lines or blank lines
            body = body.replace("\n", " ").replace("\r", " ").replace("\t", " ")
            canvas.tag_bind(task_text, "<Enter>", lambda event, body=body[:100] + "...": status_bar.config(text=body))
        else:
            canvas.tag_bind(task_text, "<Enter>", lambda event, body=body: status_bar.config(text=body))

        canvas.tag_bind(task_text, "<Leave>", lambda event: status_bar.config(text=""))


        #if task is completed draw light green check mark on the right to the text


        draw_due_date_tasks(task, x3, y3)

        #if task is of category Projects draw a blue star on the right to the text
        #if one of the categories is Projects
        if "Projects" in category:
            #draw a "Projects headline" if this is the first of the tasks with category Projects
            canvas.create_text(x3 + 480, y3, text="★", fill="gold", font=("Arial", 12), anchor=tk.W)


def draw_due_date_tasks(task, x3, y3):
    # if task has due date and due date is in the current year or the next
    if task.DueDate and task.DueDate.year in [datetime.datetime.today().year, datetime.datetime.today().year + 1]:
        #draw due date in small font right under the task text
        canvas.create_text(x3, y3 + 15, text=task.DueDate.strftime("%d/%m/%Y"), fill="grey", font=("Arial", 8),
                           anchor=tk.W)

        #convert due date to a datetime object
        due_date = datetime.datetime.strptime(task.DueDate.strftime("%d/%m/%Y"), "%d/%m/%Y")

        if due_date < datetime.datetime.today():
            canvas.create_text(x3 + 440, y3, text="⚠", fill="red", font=("Arial", 12), anchor=tk.W)
        # if task is due today draw orange exclamation mark on the right to the text
        if due_date == datetime.datetime.today():
            canvas.create_text(x3 + 440, y3, text="⚠", fill="orange", font=("Arial", 12), anchor=tk.W)
        # if task is due tomorrow draw yellow exclamation mark on the right to the text
        if due_date == datetime.datetime.today() + datetime.timedelta(days=1):
            canvas.create_text(x3 + 440, y3, text="⚠", fill="yellow", font=("Arial", 12), anchor=tk.W)
        # if task is due next week draw green exclamation mark on the right to the text
        if due_date == datetime.datetime.today() + datetime.timedelta(days=7):
            canvas.create_text(x3 + 440, y3, text="⚠", fill="green", font=("Arial", 9), anchor=tk.W)

        #if task is complete draw a check mark on the right side of the due date text
        if task.Status == 2:
            canvas.create_text(x3 + 310, y3, text="✓", fill="green", font=("Arial", 12), anchor=tk.W)


def load_tasks(show_projects=False):
    global tasks_by_category, task, category
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    tasks.Sort("[CreationTime]", True)
    tasks = tasks.Restrict("[Complete] = False")
    # create a list to store the tasks by category
    tasks_by_category = []
    # loop through the tasks and check their category
    for task in tasks:
        # get the category of the task
        category = task.Categories
        #if task is category A, B, or C
        if show_projects:
            if category == "A" or category == "B" or category == "C" or "Projects" in category:
                # add task to list
                tasks_by_category.append((task, category))
        else:
            if category == "A" or category == "B" or category == "C":
                #add task to list
                tasks_by_category.append((task, category))


    # sort the list by category in ascending order
    tasks_by_category.sort(key=lambda x: x[1])
    draw_tasks()


load_tasks()


# create a root window for the GUI


def mark_done(event, task):
    if task.Status == 2:
        task.Status = 1
    else:
        task.Status = 2

    item = canvas.find_withtag("current")
    # Change its fill color to gold
    canvas.itemconfig(item, fill="gold")
    #draw a check mark on the right side of the task text
    canvas.create_text(event.x + 310, event.y, text="✓", fill="green", font=("Arial", 12), anchor=tk.W)



    task.Save()
    canvas.itemconfig(item, fill="gold")

    #reload the tasks after 3 seconds

    root.after(1000, load_tasks)


def save_task(subject, category, due_date, popup = None):

    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    # create a new item of type task
    task = tasks.Add(3)
    #set task to high priority
    task.Importance = 2
    # set the subject
    task.Subject = subject
    # set the due date
    # set the category
    task.Categories = category
    #set due date to tomorrow
    #get tomorrow's date
    # set the reminder to tomorrow at 9 AM
    task.ReminderSet = True

    #if due date is today
    if due_date == "Today":
        #set start date to today
        task.StartDate = datetime.datetime.today()
        #set due date to today
        task.DueDate = datetime.datetime.today()
        #set reminder date
        task.ReminderTime = datetime.datetime.today() + datetime.timedelta(days=1)
        task.ReminderTime = task.ReminderTime.replace(hour=17, minute=0, second=0, microsecond=0)

    #if due date is tomorrow
    elif due_date == "Tomorrow":
        #set start date to today
        task.StartDate = datetime.datetime.today()
        #set due date to tomorrow
        task.DueDate = datetime.datetime.today() + datetime.timedelta(days=1)
        #set reminder date
        task.ReminderTime = datetime.datetime.today() + datetime.timedelta(days=1)
        task.ReminderTime = task.ReminderTime.replace(hour=9, minute=0, second=0, microsecond=0)

    #if due date is next week
    elif due_date == "Next Week":
        #set start date to today
        task.StartDate = datetime.datetime.today()
        #set due date to tomorrow
        task.DueDate = datetime.datetime.today() + datetime.timedelta(days=7)
        #set reminder date
        task.ReminderTime = datetime.datetime.today() + datetime.timedelta(days=1)
        # 9 Am
        task.ReminderTime = task.ReminderTime.replace(hour=9, minute=0, second=0, microsecond=0)



    # save the task in the tasks folder
    task.Save()
    #if popup is not None
    if popup:
        #destroy the popup
        popup.destroy()
    load_tasks()


# create a new task in a popup window
def create_new_task_popup():
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Create New Task")
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)

    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)

    # create a label to display the task subject
    subject_label = tk.Label(frame)
    subject_label.config(text="Subject:", bg="white")
    subject_label.grid(row=0, column=0, padx=10, pady=10)

    # create an entry to get the task subject
    subject_entry = tk.Entry(frame)
    subject_entry.config(width=30)
    subject_entry.grid(row=0, column=1, padx=10, pady=10)

    # create a label to display the task due date
    due_date_label = tk.Label(frame)
    due_date_label.config(text="Due Date:", bg="white")
    due_date_label.grid(row=1, column=0, padx=10, pady=10)

    # create a label to display the task category
    category_label = tk.Label(frame)
    category_label.config(text="Category:", bg="white")
    category_label.grid(row=2, column=0, padx=10, pady=10)

    # create a variable to store the category
    category_var = tk.StringVar()
    category_var.set("A")

    # create a radio button for category
    category_radio_button_a = tk.Radiobutton(frame)
    category_radio_button_a.config(text="A", variable=category_var, value="A", bg="white")
    category_radio_button_a.grid(row=2, column=1, padx=10, pady=10)

    # create a radio button for category
    category_radio_button_b = tk.Radiobutton(frame)
    category_radio_button_b.config(text="B", variable=category_var, value="B", bg="white")
    category_radio_button_b.grid(row=2, column=2, padx=10, pady=10)

    # create a radio button for category
    category_radio_button_c = tk.Radiobutton(frame)
    category_radio_button_c.config(text="C", variable=category_var, value="C", bg="white")
    category_radio_button_c.grid(row=2, column=3, padx=10, pady=10)

    #radio buttons for due date today, tomorrow, next week
    # create a variable to store the category
    due_date_var = tk.StringVar()
    due_date_var.set("Today")


    due_date_radio_button_today = tk.Radiobutton(frame)
    due_date_radio_button_today.config(text="Today", variable=due_date_var, value="Today", bg="white")
    due_date_radio_button_today.grid(row=1, column=1, padx=10, pady=10)
    #tomorrow
    due_date_radio_button_tomorrow = tk.Radiobutton(frame)
    due_date_radio_button_tomorrow.config(text="Tomorrow", variable=due_date_var, value="Tomorrow", bg="white")
    due_date_radio_button_tomorrow.grid(row=1, column=2, padx=10, pady=10)
    #next week
    due_date_radio_button_next_week = tk.Radiobutton(frame)
    due_date_radio_button_next_week.config(text="Next Week", variable=due_date_var, value="Next Week", bg="white")
    due_date_radio_button_next_week.grid(row=1, column=3, padx=10, pady=10)







    # create a button to save the task
    save_button = tk.Button(frame)
    save_button.config(text="Save", command=lambda: save_task(subject_entry.get(), category_var.get(), due_date_var.get(), popup),
                       bg="white")


    # create a button to cancel the task
    cancel_button = tk.Button(frame)
    cancel_button.config(text="Cancel", command=popup.destroy, bg="white")
    # show save and cancel buttons at the bottom right corner
    save_button.grid(row=4, column=2, padx=10, pady=10)
    cancel_button.grid(row=4, column=3, padx=10, pady=10)
    #focus on subject entry
    subject_entry.focus_set()
    #bind enter key to save button
    popup.bind("<Return>", lambda event: save_task(subject_entry.get(), category_var.get(), due_date_var.get(), popup))


# loop through the sorted list and draw the tasks on the canvas
draw_tasks()

# define a variable to keep track of how many times the dots have blinked in a minute
# add reload button to root

#bind create new task to ctrl+n
root.bind("<Control-n>", lambda event: create_new_task_popup())
#bind reload to ctrl+r
root.bind("<Control-r>", lambda event: load_tasks())
#bind reload when window is focused or clicked or maximized but not on startup

#create help icon on canvas on the top right corner
help_icon = canvas.create_text(680, 20, text="?", fill="black", font=("Arial", 20), anchor=tk.CENTER)
#bind help icon to open help window
canvas.tag_bind(help_icon, "<Button-1>", lambda event: open_help_window())

#delete frog and text function
def delete_frog_and_text(frog_smiling_face, eat_the_frog_text):
    #delete frog and text
    canvas.delete(frog_smiling_face)
    canvas.delete(eat_the_frog_text)
    #remove blinking loading text
    #reload canvas
    load_tasks()

#function to hide Project tasks
def hide_project_tasks():
    #loop through the tasks and hide the ones with category Projects
    load_tasks(show_projects=True)

#call hide project tasks function when pressing CTRL+H
root.bind("<Control-p>", lambda event: hide_project_tasks())

def open_help_window():
    #create a popup window
    popup = tk.Toplevel(root)
    popup.title("Help")
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)

    #create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)
    #make the frame as big as the popup window
    popup.grid_rowconfigure(0, weight=1)

    #create a label to display the help text
    help_text = tk.Label(frame)
    #display the help text
    help_text.config(text="Press ctrl+n to create a new task\nPress ctrl+r to reload the tasks\nDouble click on a task to open it in Outlook\nRight click on a task to mark it as complete\nHover over a task to see its body\nPress esc to exit the app")
    #pack the label
    #make the frame fill the popup window

    help_text.pack()



# start the main loop of the GUI

root.mainloop()

# import datetime with today
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


root.bind("<Escape>", close_window)

# make the window full screen from start
# root.attributes('-fullscreen', True)
root.config(bg="white")
# root title "Eat the frog"
root.title("Eat the frog")
# make root window full screen
root.state("zoomed")
# create a canvas to draw on
canvas = tk.Canvas(root)
# make the canvas fill the root window
canvas.pack(fill=tk.BOTH, expand=True)

# add status bar to root
status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# add scrollbar to canvas
scrollbar = tk.Scrollbar(canvas)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

#find all lines in body with "Time worked: " somewhere in them stripped from new lines or blank lines or spaces or whatever
def find_time_worked_in_body(task):
    #find all lines in body with "Time worked: " somewhere in them
    lines_with_time_worked = []
    for line in task.Body.splitlines():
        if "Time worked: " in line:
            lines_with_time_worked.append(line)
    return lines_with_time_worked

def convert_seconds_to_string(seconds):
    #if seconds is 0 or less or None or empty string
    if seconds is None or seconds == "" or seconds <= 0:
        #return 00:00:00
        return "00:00:00"
    else:
        #convert the total time worked to a datetime object
        time_worked = datetime.datetime.strptime(str(datetime.timedelta(seconds=seconds)), "%H:%M:%S")
        #convert the time worked to a string of HH:MM:SS
        time_worked = time_worked.strftime("%H:%M:%S")
    return time_worked

#extract all the find_time_worked_in_body lines with "Time worked: " somewhere in them and return the total time worked in seconds
def find_total_time_worked_in_body(task):
    #find all lines in body with "Time worked: " somewhere in them
    lines_with_time_worked = find_time_worked_in_body(task)
    #extract all the find_time_worked_in_body lines with "Time worked: " somewhere in them and return the total time worked in seconds
    total_time_worked = 0
    for line in lines_with_time_worked:
        #extract the time worked from the line
        time_worked = line.split("Time worked: ")[1].split(" ")[0]
        #convert the time worked to a datetime object
        time_worked = datetime.datetime.strptime(time_worked, "%H:%M:%S")
        #convert the time worked to seconds
        time_worked = time_worked.hour * 60 * 60 + time_worked.minute * 60 + time_worked.second
        #add the time worked to the total time worked
        total_time_worked += time_worked
        #convert the total time worked to a datetime object
    return convert_seconds_to_string(total_time_worked)

#convert an int of seconds to a string of HH:MM:SS


def open_timer_window(event, task):
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Timer")
    popup.worked_time = 0
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)
    # make the window full screen from start
    # root.attributes('-fullscreen', True)
    popup.config(bg="white")
    # root title "Eat the frog"
    popup.title("Timer")
    # make root window full screen
    popup.state("zoomed")
    #make popup white background
    popup.config(bg="white")
    #maximize popup and make it not and really fullscreen without the taskbar
    popup.wm_state('zoomed')
    popup.overrideredirect(True)

    # create a canvas to draw on
    canvas = tk.Canvas(popup)
    # make the canvas fill the root window
    canvas.pack(fill=tk.BOTH, expand=True)
    # add status bar to root
    status_bar = tk.Label(popup, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    # add scrollbar to canvas
    # create a label to display the task subject
    subject_label = tk.Label(canvas)
    subject_label.config(text=task.Subject, bg="white")
    subject_label.pack()
    # create a label to display the task due date
    due_date_label = tk.Label(canvas)
    due_date_label.config(text=task.DueDate.strftime("%d/%m/%Y"), bg="white")
    due_date_label.pack()
    # create a label to display the task category
    category_label = tk.Label(canvas)
    category_label.config(text=task.Categories, bg="white")
    category_label.pack()
    # create a label to display the task body
    body_label = tk.Label(canvas)
    body_label.config(text=task.Body, bg="white")
    body_label.pack()
    # create a label to display the timer
    timer_label = tk.Label(canvas)
    timer_label.config(text="25:00", bg="white", font=("Arial", 100))
    timer_label.pack()
    #make all labels white background
    subject_label.config(bg="white")
    due_date_label.config(bg="white")
    category_label.config(bg="white")
    body_label.config(bg="white")
    timer_label.config(bg="white")
    #make subject label with big font
    subject_label.config(font=("Arial", 30))
    #make canvas white background
    canvas.config(bg="white")
    #make popup white background
    popup.config(bg="white")
    #on press escape, close this popup and focus on root
    popup.bind("<Escape>", lambda event: close_popup_and_save_time_in_task(event, popup, task))

    #start the timer
    start_timer(timer_label, popup)
    #save worked time in a variable

def close_popup_and_save_time_in_task(event, popup, task):
    #save worked time in a variable
    worked_time = popup.worked_time
    #close popup
    popup.destroy()
    #focus root
    root.focus_set()
    #save worked time in task
    task.ReminderTime = task.ReminderTime + datetime.timedelta(seconds=worked_time)
    #append to task body the worked time HH:MM:SS format, and then the todays date
    task.Body = task.Body + "\nTime worked: " + str(datetime.timedelta(seconds=worked_time)) + " " + datetime.datetime.today().strftime("%d/%m/%Y")
    #update main window tasks
    load_tasks()
    #save task
    task.Save()

def start_timer(timer_label, popup):
    #start the timer
    #create a variable to store the time
    time = 25 * 60
    #create a function to update the timer
    def update_timer():
        #update the time
        nonlocal time
        #if time is 0
        if time == 0:
            #destroy the popup
            popup.destroy()
            #return
            return
        #calculate the minutes and seconds
        minutes = time // 60
        seconds = time % 60
        #if seconds is less than 10
        if seconds < 10:
            #add a 0 before the seconds
            seconds = "0" + str(seconds)
        #update the timer label
        #if time is up, show a happy frog that says "Time's up!"
        if time <= 1:
            popup_canvas = popup.winfo_children()[0]
            popup_canvas.create_text(350, 200, text="ðŸ˜Š", fill="green", font=("Arial", 100),
                                     anchor=tk.CENTER)
            time_up_text = popup_canvas.create_text(350, 300, text="Time's up!", fill="green", font=("Arial", 30),
                                              anchor=tk.CENTER)


        timer_label.config(text=f"{minutes}:{seconds}")
        #decrement the time
        time -= 1
        popup.worked_time += 1
        #call the update timer function after 1 second
        popup.after(1000, update_timer)

    update_timer()

def draw_tasks():
    # clear
    canvas.delete("all")
    if len(tasks_by_category) == 0:
        # draw a big light green check mark on the canvas
        canvas.create_text(350, 200, text="âœ“", fill="green", font=("Arial", 100), anchor=tk.CENTER)
        # draw a text under the check mark saying "No tasks"
        canvas.create_text(350, 300, text="No tasks", fill="green", font=("Arial", 30), anchor=tk.CENTER)
        # draw a text under the check mark saying "Press ctrl+n to create a new task"
        canvas.create_text(350, 350, text="Press ctrl+n to create a new task", fill="green", font=("Arial", 15),
                           anchor=tk.CENTER)
        # draw a text under the check mark saying "Press ctrl+r to reload the tasks"
        canvas.create_text(350, 380, text="Press ctrl+r to reload the tasks", fill="green", font=("Arial", 15),
                           anchor=tk.CENTER)
        # press ctrl+p to show project tasks
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
        #when mouse wheel clicking the dot open a timer window
        canvas.tag_bind(dot, "<Button-2>", lambda event, task=task: open_timer_window(event, task))
        # draw a text with the task information on the canvas with black color and Arial font size 14
        task_text = canvas.create_text(x3, y3, text=task_info, fill="black", font=("Arial", 12), anchor=tk.W)
        # when double clicking the task text open the task in outlook
        canvas.tag_bind(task_text, "<Double-Button-1>", lambda event, task=task: task.Display())
        #when mouse wheel clicking the task text open a timer window
        canvas.tag_bind(task_text, "<Button-2>", lambda event, task=task: open_timer_window(event, task))

        #if task has Time worked: in the body draw that time worked in small font right under the task text, next to the due date
        if find_total_time_worked_in_body(task) != "00:00:00":
            #draw time worked in small font right under the task text, next to the due date
            canvas.create_text(x3 + 440, y3 + 15, text=find_total_time_worked_in_body(task), fill="grey", font=("Arial", 8),
                               anchor=tk.W)

        # if task is completed draw light green check mark on the right to the text

        draw_due_date_tasks(task, x3, y3)

        # if task is of category Projects draw a blue star on the right to the text
        # if one of the categories is Projects
        if "Projects" in category:
            # draw a "Projects headline" if this is the first of the tasks with category Projects
            canvas.create_text(x3 + 480, y3, text="â˜…", fill="gold", font=("Arial", 12), anchor=tk.W)

        # if task is drawn out of the canvas
        if y3 > 800:
            # show the scrollbar
            canvas.config(scrollregion=canvas.bbox("all"))
            # bind the scrollbar to the canvas
            canvas.config(yscrollcommand=scrollbar.set)
            scrollbar.config(command=canvas.yview)
            # bind the mouse wheel to the scrollbar
            canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))
            # bind the up and down arrow keys to the scrollbar
            canvas.bind_all("<Up>", lambda event: canvas.yview_scroll(-1, "units"))
            canvas.bind_all("<Down>", lambda event: canvas.yview_scroll(1, "units"))
            # when dragging scrollbar with mouse scroll the canvas
            scrollbar.bind("<B1-Motion>", lambda event: canvas.yview_moveto(event.y))
        # else hide the scrollbar
        else:
            canvas.config(scrollregion=(0, 0, 0, 0))


def draw_due_date_tasks(task, x3, y3):
    # if task has due date and due date is in the current year or the next
    if task.DueDate and task.DueDate.year in [datetime.datetime.today().year, datetime.datetime.today().year + 1]:
        # draw due date in small font right under the task text
        canvas.create_text(x3, y3 + 15, text=task.DueDate.strftime("%d/%m/%Y"), fill="grey", font=("Arial", 8),
                           anchor=tk.W)

        # convert due date to a datetime object
        due_date = datetime.datetime.strptime(task.DueDate.strftime("%d/%m/%Y"), "%d/%m/%Y")

        if due_date < datetime.datetime.today() and task.Status != 2:
            canvas.create_text(x3 + 440, y3, text="âš ", fill="red", font=("Arial", 12), anchor=tk.W)
        # if task is due today draw orange exclamation mark on the right to the text
        if due_date == datetime.datetime.today() and task.Status != 2:
            canvas.create_text(x3 + 440, y3, text="âš ", fill="orange", font=("Arial", 12), anchor=tk.W)
        # if task is due tomorrow draw yellow exclamation mark on the right to the text
        if due_date == datetime.datetime.today() + datetime.timedelta(days=1) and task.Status != 2:
            canvas.create_text(x3 + 440, y3, text="âš ", fill="yellow", font=("Arial", 12), anchor=tk.W)
        # if task is due next week draw green exclamation mark on the right to the text
        if due_date == datetime.datetime.today() + datetime.timedelta(days=7) and task.Status != 2:
            canvas.create_text(x3 + 440, y3, text="âš ", fill="green", font=("Arial", 9), anchor=tk.W)

        # if task is complete draw a check mark on the right side of the due date text
        if task.Status == 2:
            canvas.create_text(x3 + 310, y3, text="âœ“", fill="green", font=("Arial", 12), anchor=tk.W)


def load_tasks(show_projects=False, show_tasks_finished_today=False, show_all_tasks=False,
               show_only_this_category=None):
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
    if show_tasks_finished_today:
        #restrict tasks to tasks finished today
        tasks = tasks.Restrict("[Complete] = True AND [DateCompleted] >= '" + datetime.datetime.today().strftime("%d/%m/%Y") + "'")
    # if task are not finished today but not show all tasks
    elif not show_all_tasks:
        # filter tasks by not completed tasks
        tasks = tasks.Restrict("[Complete] = False")

    # create a list to store the tasks by category
    tasks_by_category = []
    # loop through the tasks and check their category
    for task in tasks:
        # get the category of the task
        category = task.Categories
        if show_only_this_category:
            # only append the category in the string show_only_this_category
            if category == show_only_this_category:
                tasks_by_category.append((task, category))
                continue

        # if show all tasks is true
        if show_all_tasks:
            # add task to list
            tasks_by_category.append((task, category))
            continue
        # if task is category A, B, or C
        if show_projects:
            if category == "A" or category == "B" or category == "C" or "Projects" in category:
                # add task to list
                tasks_by_category.append((task, category))
        else:
            if show_only_this_category is None:
                if category == "A" or category == "B" or category == "C":
                    # add task to list
                    tasks_by_category.append((task, category))
                # add task to list

    # sort the list by category in ascending order
    tasks_by_category.sort(key=lambda x: x[1])
    draw_tasks()
    return tasks_by_category


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
    # draw a check mark on the right side of the task text
    canvas.create_text(event.x + 310, event.y, text="âœ“", fill="green", font=("Arial", 12), anchor=tk.W)

    task.Save()
    canvas.itemconfig(item, fill="gold")

    # reload the tasks after 3 seconds

    root.after(3000, load_tasks)


def save_task(subject, category, due_date, popup=None):
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
    # set task to high priority
    task.Importance = 2
    # set the subject
    task.Subject = subject
    # set the due date
    # set the category
    task.Categories = category
    # set due date to tomorrow
    # get tomorrow's date
    # set the reminder to tomorrow at 9 AM
    task.ReminderSet = True

    # if due date is today
    if due_date == "Today":
        # set start date to today
        task.StartDate = datetime.datetime.today()
        # set due date to today
        task.DueDate = datetime.datetime.today()
        # set reminder date
        task.ReminderTime = datetime.datetime.today() + datetime.timedelta(days=1)
        task.ReminderTime = task.ReminderTime.replace(hour=17, minute=0, second=0, microsecond=0)

    # if due date is tomorrow
    elif due_date == "Tomorrow":
        # set start date to today
        task.StartDate = datetime.datetime.today()
        # set due date to tomorrow
        task.DueDate = datetime.datetime.today() + datetime.timedelta(days=1)
        # set reminder date
        task.ReminderTime = datetime.datetime.today() + datetime.timedelta(days=1)
        task.ReminderTime = task.ReminderTime.replace(hour=9, minute=0, second=0, microsecond=0)

    # if due date is next week
    elif due_date == "Next Week":
        # set start date to today
        task.StartDate = datetime.datetime.today()
        # set due date to tomorrow
        task.DueDate = datetime.datetime.today() + datetime.timedelta(days=7)
        # set reminder date
        task.ReminderTime = datetime.datetime.today() + datetime.timedelta(days=1)
        # 9 Am
        task.ReminderTime = task.ReminderTime.replace(hour=9, minute=0, second=0, microsecond=0)

    # save the task in the tasks folder
    task.Save()
    # if popup is not None
    if popup:
        # destroy the popup
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

    # radio buttons for due date today, tomorrow, next week
    # create a variable to store the category
    due_date_var = tk.StringVar()
    due_date_var.set("Today")

    due_date_radio_button_today = tk.Radiobutton(frame)
    due_date_radio_button_today.config(text="Today", variable=due_date_var, value="Today", bg="white")
    due_date_radio_button_today.grid(row=1, column=1, padx=10, pady=10)
    # tomorrow
    due_date_radio_button_tomorrow = tk.Radiobutton(frame)
    due_date_radio_button_tomorrow.config(text="Tomorrow", variable=due_date_var, value="Tomorrow", bg="white")
    due_date_radio_button_tomorrow.grid(row=1, column=2, padx=10, pady=10)
    # next week
    due_date_radio_button_next_week = tk.Radiobutton(frame)
    due_date_radio_button_next_week.config(text="Next Week", variable=due_date_var, value="Next Week", bg="white")
    due_date_radio_button_next_week.grid(row=1, column=3, padx=10, pady=10)

    # create a button to save the task
    save_button = tk.Button(frame)
    save_button.config(text="Save",
                       command=lambda: save_task(subject_entry.get(), category_var.get(), due_date_var.get(), popup),
                       bg="white")

    # create a button to cancel the task
    cancel_button = tk.Button(frame)
    cancel_button.config(text="Cancel", command=popup.destroy, bg="white")
    # show save and cancel buttons at the bottom right corner
    save_button.grid(row=4, column=2, padx=10, pady=10)
    cancel_button.grid(row=4, column=3, padx=10, pady=10)
    # focus on subject entry
    subject_entry.focus_set()
    # bind enter key to save button
    popup.bind("<Return>", lambda event: save_task(subject_entry.get(), category_var.get(), due_date_var.get(), popup))


# loop through the sorted list and draw the tasks on the canvas
draw_tasks()

# define a variable to keep track of how many times the dots have blinked in a minute
# add reload button to root

# bind create new task to ctrl+n
root.bind("<Control-n>", lambda event: create_new_task_popup())
# bind reload to ctrl+r
root.bind("<Control-r>", lambda event: load_tasks())
# bind reload when window is focused or clicked or maximized but not on startup

# create help icon on canvas on the top right corner
help_icon = canvas.create_text(680, 20, text="?", fill="black", font=("Arial", 20), anchor=tk.CENTER)
# bind help icon to open help window
canvas.tag_bind(help_icon, "<Button-1>", lambda event: open_help_window())


# add generate html file button to root

# delete frog and text function
def delete_frog_and_text(frog_smiling_face, eat_the_frog_text):
    # delete frog and text
    canvas.delete(frog_smiling_face)
    canvas.delete(eat_the_frog_text)
    # remove blinking loading text
    # reload canvas
    load_tasks()


# function to hide Project tasks
def hide_project_tasks():
    # loop through the tasks and hide the ones with category Projects
    load_tasks(show_projects=True)


# call hide project tasks function when pressing CTRL+H
root.bind("<Control-p>", lambda event: hide_project_tasks())
root.bind("<Control-t>", lambda event: load_tasks(show_tasks_finished_today=True))
root.bind("<Control-a>", lambda event: load_tasks(show_only_this_category="A"))
root.bind("<Control-b>", lambda event: load_tasks(show_only_this_category="B"))
root.bind("<Control-c>", lambda event: load_tasks(show_only_this_category="C"))


def show_only_this_category(category):
    load_tasks(show_only_this_category=category)


# make function to generate a html file with the finished tasks with dots and finished dates
def generate_html_file():
    # while html file is being generated show a blinking loading text on root window
    # create a label to display the task subject


    #show file dialog to save the html file
    import tkinter.filedialog
    #get the file path
    file_path = tkinter.filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML files", "*.html")])
    #if file path is empty
    if file_path == "":
        #return
        return

    # open the html file with utf-8 encoding
    html_file = open(file_path, "w", encoding="utf-8")
    #if open was successful
    if not html_file:
        #show error message
        tk.messagebox.showerror("Error", "Could not open file")

    # write the html file header
    html_file.write("<html><head><title>Finished Tasks</title></head><body>")
    # load nice google font
    html_file.write("<link href='https://fonts.googleapis.com/css?family=Roboto' rel='stylesheet'>")
    # make the font of the html file Roboto
    html_file.write("<style>body {font-family: 'Roboto', sans-serif;}</style>")
    # make that font the default font for the html file
    # loop through the tasks and write the html file
    for task, category in load_tasks(show_all_tasks=True):
        # if task is completed
        # write the task subject and finished date
        # if task is category A, B, or C, draw a dot with the color of the category
        if task.Status == 2 and task.DateCompleted is not None:
            # get the task subject
            subject = task.Subject
            # get the task finished date
            finished_date = task.DateCompleted
            # get the task category
            category = task.Categories
            # write the task subject and finished date
            # if task is finished draw a big green checkmark before the dot
            generate_dots_and_subjects(category, finished_date, html_file, subject, task)

    # write the html file footer
    html_file.write("</body></html>")
    # remove blinking loading text
    # reload canvas
    load_tasks()
    # close the html file
    html_file.close()
    # open the html file
    import webbrowser
    #open the html file in the default browser
    webbrowser.open(file_path)


def generate_dots_and_subjects(category, finished_date, html_file, subject, task):
    if task.Status == 2:
        html_file.write("<span style='color:green; font-size:40px'>âœ“</span> ")
    if category == "A":
        # write big dot and then subject and finished date
        html_file.write(
            "<span style='color:red; font-size:40px'>â—</span> " + " " + subject + " - " + finished_date.strftime(
                "%d/%m/%Y") + "<br>")
    elif category == "B":
        html_file.write(
            "<span style='color:yellow; font-size:40px'>â—</span> " + " " + subject + " - " + finished_date.strftime(
                "%d/%m/%Y") + "<br>")
    elif category == "C":
        html_file.write(
            "<span style='color:green; font-size:40px'>â—</span>" + " " + subject + " - " + finished_date.strftime(
                "%d/%m/%Y") + "<br>")
    else:
        html_file.write(
            "<span style='color:gold; font-size:40px'>â—</span>" + " " + subject + " - " + finished_date.strftime(
                "%d/%m/%Y") + "<br>")
    # if the task has a body print it in small grey nice formatted letters under the task subject with space before and after
    if task.Body is not None:
        html_file.write(
            "<span style='color:grey; font-size:10px; margin-left: 20px; margin-right: 20px;'>" + task.Body + "</span><br>")


def open_help_window():
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Help")
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)

    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)
    # make the frame as big as the popup window
    popup.grid_rowconfigure(0, weight=1)

    # create a label to display the help text
    help_text = tk.Label(frame)
    # display the help text
    help_text.config(
        text="Press ctrl+n to create a new task\nPress ctrl+r to reload the tasks\nDouble click on a task to open it in Outlook\nRight click on a task to mark it as complete\nHover over a task to see its body\nPress esc to exit the app")
    # pack the label
    # make the frame fill the popup window

    help_text.pack()


# add generate html file button to ctrl+g

# add file menu to root
menu_bar = tk.Menu(root)
# add file menu to menu bar
file_menu = tk.Menu(menu_bar, tearoff=0)
# add file menu to menu bar
menu_bar.add_cascade(label="File", menu=file_menu)
# add generate html file to file menu
file_menu.add_command(label="Finished report", command=generate_html_file)
# add exit to file menu
file_menu.add_command(label="Exit", command=root.destroy)
#add show only category A to file menu
file_menu.add_command(label="Show only category A", command=lambda: show_only_this_category("A"))
#add show only category B to file menu
file_menu.add_command(label="Show only category B", command=lambda: show_only_this_category("B"))
#add show only category C to file menu
file_menu.add_command(label="Show only category C", command=lambda: show_only_this_category("C"))
#add file menu to show_all_tasks
file_menu.add_command(label="Show all tasks", command=lambda: load_tasks(show_all_tasks=True))
# add file menu to root
root.config(menu=menu_bar)
#on focus or click or maximize reload the tasks
root.bind("<FocusIn>", lambda event: load_tasks())

root.mainloop()

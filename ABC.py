# import datetime with today
import datetime
import os
import sqlite3
from tkinter import messagebox, filedialog, simpledialog

import win32com.client
import tkinter as tk  # you can use tkinter or another library to create a GUI

# set some constants for the drawing
dot_radius = 8  # radius of each dot
dot_gap = 10  # gap between each dot and task
text_gap = 20  # gap between each task and text
line_height = 30  # height of each line
root = tk.Tk()
# import Inbox class from inbox.py
from inbox import Inbox
# import Note class from note.py
from note import Note


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


def create_filter_buttons():
    frame = tk.Frame(root)
    frame.config(bg="white")
    frame.pack(side=tk.TOP)
    a_filter_button = tk.Button(frame)
    a_filter_button.config(text="A", bg="white", command=lambda: load_tasks(show_only_this_category="A"))
    a_filter_button.pack(side=tk.LEFT)
    b_filter_button = tk.Button(frame)
    b_filter_button.config(text="B", bg="white", command=lambda: load_tasks(show_only_this_category="B"))
    b_filter_button.pack(side=tk.LEFT)
    c_filter_button = tk.Button(frame)
    c_filter_button.config(text="C", bg="white", command=lambda: load_tasks(show_only_this_category="C"))
    c_filter_button.pack(side=tk.LEFT)
    all_button = tk.Button(frame)
    all_button.config(text="All", bg="white", command=lambda: load_tasks())
    all_button.pack(side=tk.LEFT)
    projects_button = tk.Button(frame)
    projects_button.config(text="Projects", bg="white", command=lambda: load_tasks(show_only_this_category="Projects"))
    projects_button.pack(side=tk.LEFT)



#add three buttons on top of the canvas
create_filter_buttons()
# add a button to create a new note


#add small notes button next to
# create a canvas to draw on
canvas = tk.Canvas(root)
# make the canvas fill the root window
canvas.pack(fill=tk.BOTH, expand=True)
#above canvas put a small button that loads the inbox


# add status bar to root
status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# add scrollbar to canvas
scrollbar = tk.Scrollbar(canvas)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)


def convert_seconds_to_string(seconds):
    # if seconds is 0 or less or None or empty string
    if seconds is None or seconds == "" or seconds <= 0:
        # return 00:00:00
        return "00:00:00"
    else:
        # convert the total time worked to a datetime object
        time_worked = datetime.datetime.strptime(str(datetime.timedelta(seconds=seconds)), "%H:%M:%S")
        # convert the time worked to a string of HH:MM:SS
        time_worked = time_worked.strftime("%H:%M:%S")
    return time_worked

def set_estimated_work(event, popup, timer_label, task):
    #ask user to set estimated work with simpledialog
    estimated_work = simpledialog.askinteger("Set estimated work", "How many minutes do you think this task will take?")
    #if user pressed cancel
    if estimated_work is None:
        #return
        return

    #set task TotalWork to estimated work
    task.TotalWork = estimated_work

def open_timer_window(event, task_to_start_timer_on):
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
    # make popup white background
    popup.config(bg="white")
    # maximize popup and make it not and really fullscreen without the taskbar
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
    subject_label.config(text=task_to_start_timer_on.Subject, bg="white")
    subject_label.pack()
    # create a label to display the task due date
    due_date_label = tk.Label(canvas)
    due_date_label.config(text=task_to_start_timer_on.DueDate.strftime("%d/%m/%Y"), bg="white")
    due_date_label.pack()
    # create a label to display the task category
    category_label = tk.Label(canvas)
    category_label.config(text=task_to_start_timer_on.Categories, bg="white")
    category_label.pack()
    # create a label to display the task body
    body_label = tk.Label(canvas)
    body_label.config(text=task_to_start_timer_on.Body, bg="white")
    body_label.pack()
    # create a label to display the timer
    timer_label = tk.Label(canvas)
    timer_label.config(text="25:00", bg="white", font=("Arial", 100))
    timer_label.pack()
    # create a button to stop the timer
    stop_button = tk.Button(canvas)
    stop_button.config(text="Stop", bg="white", command=lambda: close_popup_and_save_time_in_task(event, popup, task_to_start_timer_on))
    stop_button.pack()

    # make all labels white background
    subject_label.config(bg="white")
    due_date_label.config(bg="white")
    category_label.config(bg="white")
    body_label.config(bg="white")
    timer_label.config(bg="white")
    # make subject label with big font
    subject_label.config(font=("Arial", 30))
    # make canvas white background
    canvas.config(bg="white")
    # make popup white background
    popup.config(bg="white")
    # on press escape, close this popup and focus on root
    popup.bind("<Escape>", lambda event: close_popup_and_save_time_in_task(event, popup, task_to_start_timer_on))

    # if task has no TotalWork, ask if user wants to set it, TotalWork is estimated work
    if task_to_start_timer_on.TotalWork is None or task_to_start_timer_on.TotalWork == 0:
        #print a small text on the canvas that says "Click here to set estimated work" that looks like a link, under the timer
        # draw a button under the timer that says "Set estimated work"
        set_estimated_work_button = tk.Button(canvas)
        set_estimated_work_button.config(text="Set estimated work", bg="white")
        #bind button to set_estimated_work function
        set_estimated_work_button.bind("<Button-1>", lambda event: set_estimated_work(event, popup, timer_label, task_to_start_timer_on))
        set_estimated_work_button.pack()
    else:
        pass



    start_timer(timer_label, popup)
    #focus on popup






# set task ActualWork to worked time in minutes, and add to it if there is already a value
def set_task_actual_work(task, worked_time):
    # if task ActualWork is None
    if task.ActualWork is None:
        # set task ActualWork to worked time in minutes
        task.ActualWork = worked_time
    else:
        # add worked time in minutes to task ActualWork
        task.ActualWork += worked_time

    return task.ActualWork



# caluclate percentage of task done
def calculate_percentage_of_task_done(task):
    #check for division by zero
    if task.TotalWork == 0:
        return 0
    # calculate percentage of task done

    percentage = int(task.ActualWork / task.TotalWork * 100)
    if percentage > 100:
        percentage = 100
        #calculate how much time was over estimated in hours
        over_estimated_time = str(datetime.timedelta(minutes=task.ActualWork - task.TotalWork))
        #add info to task body that task took more time than estimated
        task.Body += "\n\nThis task took more time than estimated:" + over_estimated_time

    return percentage


def close_popup_and_save_time_in_task(event, popup, task):
    # save worked time in a variable
    worked_time = popup.worked_time
    # convert worked time in seconds to minutes
    worked_time_minutes = worked_time // 60
    # if worked time is 0
    if worked_time_minutes < 1:
        # show message box to ask if user wants to save time in task
        if messagebox.askyesno("Save time in task?", "You worked less than 1 minute on the task.\nDo you want to save time in task?\nIt will be saved as one minute in Outlook."):
            # set worked time to 1 minute
            worked_time_minutes = 1

    task.ActualWork = set_task_actual_work(task, worked_time_minutes)
    task.PercentComplete = calculate_percentage_of_task_done(task)
    # close popup and save time in task
    popup.destroy()
    # focus root
    root.focus_set()

    # update main window tasks
    load_tasks()
    # save task
    task.Save()


# create a popup to search notes
def search_notes_popup():
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Search Notes")
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)

    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)

    # create a label to display the note subject
    subject_label = tk.Label(frame)
    subject_label.config(text="Subject:", bg="white")
    subject_label.grid(row=0, column=0, padx=10, pady=10)

    # create an entry to get the note subject
    subject_entry = tk.Entry(frame)
    subject_entry.config(width=30)
    subject_entry.grid(row=0, column=1, padx=10, pady=10)

    # create a label to display the note body
    body_label = tk.Label(frame)
    body_label.config(text="Search in body:", bg="white")
    body_label.grid(row=1, column=0, padx=10, pady=10)

    # create an entry to get the note body
    body_entry = tk.Entry(frame)
    body_entry.config(width=30)
    body_entry.grid(row=1, column=1, padx=10, pady=10)

    # create search results listbox
    search_results_listbox = tk.Listbox(frame, width=50, height=10)
    search_results_listbox.grid(row=3, column=0, padx=10, pady=10)
    # make search results listbox nice font
    search_results_listbox.config(font=("Tahoma", 10))
    # bind list box double click to open note in outlook
    search_results_listbox.bind("<Double-Button-1>", lambda event: open_task_in_outlook(event, search_results_listbox))

    # create a search button and bind it to search_notes function
    search_button = tk.Button(frame)
    search_button.config(text="Search", bg="white",
                         command=lambda: search_notes(subject_entry.get(), body_entry.get(), search_results_listbox,
                                                      popup))
    search_button.grid(row=2, column=0, padx=10, pady=10)


def search_notes(subject, body, search_results_listbox, popup):
    # display hour glass cursor while searching
    # create an Outlook application object
    loading_icon(popup)
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for notes
    notes_folder = namespace.GetDefaultFolder(12)
    # get all the notes in the folder
    notes = notes_folder.Items
    search_results = []
    # loop through the notes and check their subject or body
    for note in notes:
        # get the category of the note
        # if note subject contains subject
        # is subject is not empty and in note subject
        if subject != "" and subject.casefold() in note.Subject:
            # add note to list
            search_results.append(note)
        if body != "" and body.casefold() in note.Body:
            # add note to list
            search_results.append(note)

    # sort the list by category in ascending order
    # update search results listbox
    search_results_listbox.delete(0, tk.END)
    for note in search_results:
        search_results_listbox.insert(tk.END, note.Subject)

    search_results_listbox.search_results = search_results
    # if no results found display message
    if len(search_results) == 0:
        search_results_listbox.insert(tk.END, "No results found")

    # display normal cursor after searching
    reset_loading_icon(popup)
    return search_results


# create popup with a search box that searches for tasks by subject and body
def search_tasks_popup():
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Search Tasks")
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

    # create a label to display the task body
    body_label = tk.Label(frame)
    body_label.config(text="Search in body:", bg="white")
    body_label.grid(row=1, column=0, padx=10, pady=10)

    # create an entry to get the task body
    body_entry = tk.Entry(frame)
    body_entry.config(width=30)
    body_entry.grid(row=1, column=1, padx=10, pady=10)

    # create a search button
    search_button = tk.Button(frame)
    search_button.config(text="Search", bg="white",
                         command=lambda: search_tasks(subject_entry.get(), body_entry.get(), search_results_listbox,
                                                      popup))
    search_button.grid(row=2, column=0, padx=10, pady=10)

    # list box to display search results
    search_results_listbox = tk.Listbox(frame, width=50, height=10)
    search_results_listbox.grid(row=3, column=0, padx=10, pady=10)
    # make search results listbox nice font
    search_results_listbox.config(font=("Tahoma", 10))

    # when pressing enter in subject entry search for tasks
    subject_entry.bind("<Return>",
                       lambda event: search_tasks(subject_entry.get(), body_entry.get(), search_results_listbox, popup))
    # when pressing enter in body entry search for tasks
    body_entry.bind("<Return>",
                    lambda event: search_tasks(subject_entry.get(), body_entry.get(), search_results_listbox, popup))

    # focus subject entry on load
    subject_entry.focus_set()

    # when clickin task in search results listbox open task in outlook
    search_results_listbox.bind("<Double-Button-1>", lambda event: open_task_in_outlook(event, search_results_listbox))


def open_task_in_outlook(event, search_results_listbox):
    # get task from current index
    index = search_results_listbox.curselection()
    # If there is an item selected
    if index and index[0] >= 0 and index[0] < len(search_results_listbox.search_results):
        # Get the corresponding task object from the high priority tasks list
        task = search_results_listbox.search_results[index[0]]
        task.Display()
    else:
        return None


def search_tasks(subject, body, search_results_listbox, popup):
    # display hour glass cursor while searching
    loading_icon(popup)
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    search_results = []
    # loop through all the tasks and check their subject or body
    for task_item in tasks:
        # get the category of the task
        # if task subject contains subject
        # is subject is not empty and in task subject, ignore case
        if subject != "" and subject.casefold() in task_item.Subject.casefold():
            # add task to list
            search_results.append(task_item)
        #is body is not empty and in task body, ignore case
        if body != "" and body.casefold() in task_item.Body.casefold():
            # add task to list
            search_results.append(task_item)

    # sort the list by category in ascending order
    # update search results listbox
    search_results_listbox.delete(0, tk.END)
    for task_item in search_results:
        # if task is done draw a green check mark on the right side of the task text
        if task_item.Status == 2:
            search_results_listbox.insert(tk.END, task_item.Subject + " ✓")
        else:
            search_results_listbox.insert(tk.END, task_item.Subject)

    search_results_listbox.search_results = search_results
    # if no results found display message
    if len(search_results) == 0:
        search_results_listbox.insert(tk.END, "No results found")

    reset_loading_icon(popup)
    return search_results


def start_timer(timer_label, popup):
    # start the timer
    # create a variable to store the time
    time = 25 * 60

    # create a function to update the timer
    def update_timer():
        # update the time
        nonlocal time
        # if time is 0
        if time == 0:
            # destroy the popup
            popup.destroy()
            # return
            return
        # calculate the minutes and seconds
        minutes = time // 60
        seconds = time % 60
        # if seconds is less than 10
        if seconds < 10:
            # add a 0 before the seconds
            seconds = "0" + str(seconds)
        # update the timer label
        # if time is up, show a happy frog that says "Time's up!"
        if time <= 1:
            popup_canvas = popup.winfo_children()[0]
            popup_canvas.create_text(350, 200, text="😊", fill="green", font=("Arial", 100),
                                     anchor=tk.CENTER)
            time_up_text = popup_canvas.create_text(350, 300, text="Time's up!", fill="green", font=("Arial", 30),
                                                    anchor=tk.CENTER)

        timer_label.config(text=f"{minutes}:{seconds}")
        # decrement the time
        time -= 1
        popup.worked_time += 1
        # call the update timer function after 1 second
        popup.after(1000, update_timer)

    update_timer()


def save_body(body_text, task, popup):
    # get body text from text widget
    body = body_text.get(1.0, tk.END)
    # set task body to body text
    task.Body = body
    # save task
    task.Save()
    # destroy popup
    popup.destroy()


def show_task_body(event, task):
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Task body")
    popup.config(bg="white")
    popup.geometry("600x400")
    # popup.resizable(False, False)
    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)

    # create a text widget to display the task body
    body_text = tk.Text(frame)
    body_text.config(width=50, height=10)
    body_text.grid(row=0, column=0, padx=10, pady=10)
    # insert task body into text widget
    body_text.insert(tk.END, task.Body)

    # show

    # add a save button under the text widget
    save_button = tk.Button(frame)
    save_button.config(text="Save", bg="white", command=lambda: save_body(body_text, task, popup))
    save_button.grid(row=1, column=0, padx=10, pady=10)


def draw_tasks():
    # clear
    canvas.delete("all")
    # display hour glass cursor while loading tasks
    loading_icon()
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

        # draw a box on the left side of the task with the color of the category, without border
        check_box = canvas.create_rectangle(x1 - dot_radius, y1 - dot_radius, x1 + dot_radius, y1 + dot_radius, fill=color,
                                outline="")

        #when clicking the checkbox draw a check mark inside it
        canvas.tag_bind(check_box, "<Button-1>", lambda event, task=task: mark_done(event, task, check_box))

        # mark done on left click
        canvas.tag_bind(check_box, "<Button-1>", lambda event, task=task: mark_done(event, task))
        # when double clicking the dot open the task in outlook
        canvas.tag_bind(check_box, "<Double-Button-1>", lambda event, task=task: task.Display())
        # when mouse wheel clicking the dot open a timer window
        canvas.tag_bind(check_box, "<Button-2>", lambda event, task=task: open_timer_window(event, task))
        # draw a text with the task information on the canvas with black color and Arial font size 14
        task_text = canvas.create_text(x3, y3, text=task_info, fill="black", font=("Arial", 12), anchor=tk.W)
        # when double clicking the task text open the task in outlook
        canvas.tag_bind(task_text, "<Double-Button-1>", lambda event, task=task: task.Display())
        # when mouse wheel clicking the task text open a timer window
        canvas.tag_bind(task_text, "<Button-2>", lambda event, task=task: open_timer_window(event, task))
        # when left clicking the task text open a popup window to show the body of the task

        #if task is done draw a green check mark inside the box
        if task.Status == 2:
            check_mark = canvas.create_text(x1, y1, text="✓", fill="light green", font=("Arial", 12), anchor=tk.CENTER)
            #bind check mark to mark_done function
            canvas.tag_bind(check_mark, "<Button-1>", lambda event, task=task: mark_done(event, task, check_mark))

        draw_tasks_with_icons(task, x3, y3)

        # if task is of category Projects draw a blue star on the right to the text
        # if one of the categories is Projects
        if "Projects" in category:
            # draw a "Projects headline" if this is the first of the tasks with category Projects
            canvas.create_text(x3 + 480, y3, text="★", fill="gold", font=("Arial", 12), anchor=tk.W)

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

    # display normal cursor after loading tasks
    reset_loading_icon()


def loading_icon(window=root):
    window.config(cursor="watch")
    window.update()


def reset_loading_icon(window=root):
    window.config(cursor="")
    window.update()


# make function that exports all tasks with actual work to an excel file but does not open excel
def export_tasks_to_excel():
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    # create a list to store the tasks by category
    tasks_by_category = []
    # loop through the tasks and check their category
    for task in tasks:
        # get the category of the task
        category = task.Categories
        if task.ActualWork:
            # add task to list
            tasks_by_category.append((task, category))
            continue

    # sort the list by category in ascending order
    tasks_by_category.sort(key=lambda x: x[1])
    # create excel file
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    # create a new workbook
    workbook = excel.Workbooks.Add()
    # create a new worksheet
    worksheet = workbook.Worksheets.Add()
    # set worksheet name to "Tasks"
    worksheet.Name = "Tasks"
    # set worksheet header
    worksheet.Cells(1, 1).Value = "Subject"
    worksheet.Cells(1, 2).Value = "Actual Work"
    worksheet.Cells(1, 3).Value = "Total Work"
    worksheet.Cells(1, 4).Value = "Percent Complete"
    worksheet.Cells(1, 5).Value = "Due Date"
    # loop through tasks and add them to excel file
    for i, (task, category) in enumerate(tasks_by_category):
        # get the task subject and due date
        subject = task.Subject
        due_date = task.DueDate

        # add task to list
        # if task is category A, B, or C, make subject cell background red, yellow, or green
        if category == "A":
            worksheet.Cells(i + 2, 1).Interior.ColorIndex = 3
        elif category == "B":
            worksheet.Cells(i + 2, 1).Interior.ColorIndex = 6
        elif category == "C":
            worksheet.Cells(i + 2, 1).Interior.ColorIndex = 4

        # add task to excel file
        worksheet.Cells(i + 2, 1).Value = subject
        worksheet.Cells(i + 2, 2).Value = task.ActualWork
        worksheet.Cells(i + 2, 3).Value = task.TotalWork
        worksheet.Cells(i + 2, 4).Value = task.PercentComplete
        worksheet.Cells(i + 2, 5).Value = due_date
        # change cell width to fit text
        worksheet.Cells(i + 2, 1).ColumnWidth = subject.__len__() + 5
        worksheet.Cells(i + 2, 2).ColumnWidth = 15
        worksheet.Cells(i + 2, 3).ColumnWidth = 15
        worksheet.Cells(i + 2, 4).ColumnWidth = 15
        worksheet.Cells(i + 2, 5).ColumnWidth = 15

    # save excel file
    # close excel file
    # workbook.Close()
    # open excel file
    excel.Visible = True
    # quit excel
    # excel.Quit()


def draw_tasks_with_icons(task, x3, y3):
    # if task has body text
    if task.Body != "":
        # draw a paperclip icon on the right side of the task text
        paperclip = canvas.create_text(x3 + 400, y3, text="📎", fill="grey", font=("Arial", 12), anchor=tk.W)
        # when double clicking the paperclip icon open the body popup
        canvas.tag_bind(paperclip, "<Double-Button-1>", lambda event, task=task: show_task_body(event, task))

    # if task has due date and due date is in the current year or the next
    if task.DueDate and task.DueDate.year in [datetime.datetime.today().year, datetime.datetime.today().year + 1]:
        due_date = datetime.datetime.strptime(task.DueDate.strftime("%d/%m/%Y"), "%d/%m/%Y")
        # draw due date in small font right under the task text, if the task is due, make it red
        if due_date < datetime.datetime.today():
            text_color = "red"
        else:
            text_color = "grey"

        canvas.create_text(x3, y3 + 15, text=task.DueDate.strftime("%d/%m/%Y"), fill=text_color, font=("Arial", 7),
                           anchor=tk.W)

        # draw actual work next to due date
        canvas.create_text(x3 + 100, y3 + 15, text="Acutal: " + convert_seconds_to_string(task.ActualWork * 60),
                           fill="grey", font=("Arial", 7),
                           anchor=tk.W)
        # draw estimated work next to actual work
        canvas.create_text(x3 + 200, y3 + 15, text="Estimated: " + convert_seconds_to_string(task.TotalWork * 60),
                           fill="grey", font=("Arial", 7),
                           anchor=tk.W)

        # draw percentage of task done next to estimated work, if it is 100% make it green
        if task.PercentComplete == 100:
            text_color = "green"
        else:
            text_color = "grey"

        canvas.create_text(x3 + 300, y3 + 15, text="Done: " + str(task.PercentComplete) + "%",
                           fill=text_color, font=("Arial", 7),
                           anchor=tk.W)

        # if task is complete draw a check mark on the right side of the due date text
        if task.Status == 2:
            canvas.create_text(x3 + 310, y3, text="✓", fill="green", font=("Arial", 12), anchor=tk.W)


# make function to export tasks on canvas sql database
def export_tasks_to_sqlite():
    # file dialog for user to choose file name and location
    file = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite database", "*.db")])
    #
    #check if file exists already
    if os.path.exists(file):
        # do you want to overwrite file?
        overwrite_file = messagebox.askyesno("Overwrite file?", "Do you want to overwrite file?")
        if overwrite_file:
            # overwrite file
            pass
        else:
            # return
            return

    # create database
    conn = sqlite3.connect(file)
    # create cursor
    c = conn.cursor()

    # create table if not exists
    c.execute("""CREATE TABLE IF NOT EXISTS tasks (
                subject text,
                actual_work integer,
                total_work integer,
                percent_complete integer,
                due_date text,
                body text,
                category text
                )""")

    # loop through tasks and add them to sqlite database
    for task, category in tasks_by_category:
        # get the task subject and due date
        subject = task.Subject
        due_date = task.DueDate
        actual_work = task.ActualWork
        total_work = task.TotalWork
        percent_complete = task.PercentComplete
        body = task.Body
        category = task.Categories
        # add task to sqlite database
        # convert pywin32 datetime to sqlite datetime
        due_date = datetime.datetime.strptime(due_date.strftime("%d/%m/%Y"), "%d/%m/%Y")
        c.execute("INSERT INTO tasks VALUES (:subject, :actual_work, :total_work, :percent_complete, :due_date, :body, :category)",
                  {"subject": subject, "actual_work": actual_work, "total_work": total_work,
                   "percent_complete": percent_complete, "due_date": due_date, "body": body, "category": category})

    # commit changes
    conn.commit()
    # close connection
    conn.close()


# make function to import tasks from sqlite database
def import_tasks_from_sqlite():
    # create database
    # file dialog for user to choose file name and location
    file = filedialog.askopenfilename(defaultextension=".db", filetypes=[("SQLite database", "*.db")])
    #connect to database
    conn = sqlite3.connect(file)

    # create cursor
    c = conn.cursor()
    # get tasks from database
    c.execute("SELECT * FROM tasks")
    # loop through tasks and add them to outlook
    for task in c.fetchall():
        # create an Outlook application object
        outlook = win32com.client.Dispatch("Outlook.Application")
        # get the namespace object
        namespace = outlook.GetNamespace("MAPI")
        # get the default folder for tasks
        tasks_folder = namespace.GetDefaultFolder(13)
        # get all the tasks in the folder
        tasks = tasks_folder.Items
        # create a new item of type task
        imported_task = tasks.Add(3)
        # set task to high priority
        imported_task.Importance = 2
        # set the subject
        imported_task.Subject = task[0]
        # set the due date
        # set the category
        # set due date to tomorrow
        # get tomorrow's date
        # set the reminder to tomorrow at 9 AM
        imported_task.ReminderSet = True
        imported_task.ActualWork = task[1]
        imported_task.TotalWork = task[2]
        imported_task.PercentComplete = task[3]
        imported_task.DueDate = task[4]
        imported_task.Body = task[5]
        imported_task.Categories = task[6]
        imported_task.Save()

    # commit changes
    conn.commit()
    # close connection
    conn.close()
    load_tasks()


def load_tasks(show_projects=False, show_tasks_finished_today=False, show_all_tasks=False,
               show_only_this_category=None):
    # clear status bar
    status_bar.config(text="")
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
        # restrict tasks to tasks finished today
        tasks = tasks.Restrict(
            "[Complete] = True AND [DateCompleted] >= '" + datetime.datetime.today().strftime("%d/%m/%Y") + "'")
        # show message in status bar that only tasks finished today are shown
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
                # show message in status bar that only this category is shown
                status_bar.config(text="Showing only " + category + " tasks")
                tasks_by_category.append((task, category))
                continue

        # if show all tasks is true
        if show_all_tasks:
            # show message in status bar that all tasks are shown
            status_bar.config(text="Showing all tasks")
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
                    if show_tasks_finished_today:
                        status_bar.config(text="Showing A, B, and C tasks finished today")
                    else:
                        status_bar.config(text="Showing A, B, and C tasks")
                    # add task to list
                    tasks_by_category.append((task, category))
                # add task to list

    # sort the list by category in ascending order
    tasks_by_category.sort(key=lambda x: x[1])
    draw_tasks()
    return tasks_by_category


load_tasks()


# create a root window for the GUI


def mark_done(event, task, check_mark=None):
    if task.Status == 2:
        task.Status = 1
    else:
        task.Status = 2

    item = canvas.find_withtag("current")
    #frame the item with a black border
    #if task is not already done
    if task.Status == 2:
        #draw a small green check mark a little bit bigger then the item, inside the item
        canvas.create_text(canvas.bbox(item)[0] + 1, canvas.bbox(item)[1] + 6, text="✓", fill="light green",
                           font=("Arial", 18), anchor=tk.W)
    else:
        #delete the check mark from item
        #remove check_mark
        canvas.delete(check_mark)


    task.Save()

    # reload the tasks after 3 seconds

    root.after(3000, load_tasks)


def save_task(subject, category, due_date, create_calendar_event, popup=None):
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

    # check if create calendar event is checked
    if create_calendar_event:
        create_new_calendar_event_based_on_task(category, due_date, namespace, subject)

    # if due date is today
    set_due_date_on_task(due_date, task)

    # save the task in the tasks folder
    task.Save()
    # if popup is not None
    if popup:
        # destroy the popup
        popup.destroy()
    load_tasks()


def create_new_calendar_event_based_on_task(category, due_date, namespace, subject):
    # create a new calendar event
    # get the default folder for calendar events
    calendar_folder = namespace.GetDefaultFolder(9)
    # get all the calendar events in the folder
    calendar_events = calendar_folder.Items
    # create a new item of type calendar event
    calendar_event = calendar_events.Add(1)
    # set the subject
    calendar_event.Subject = subject
    # if color of task is red, set calendar event to red
    # save the calendar event in the calendar folder
    # if category is A, set calendar event to red
    if category == "A":
        calendar_event.Categories = "A"
    # if category is B, set calendar event to yellow
    elif category == "B":
        calendar_event.Categories = "B"
    # if category is C, set calendar event to green
    elif category == "C":
        calendar_event.Categories = "C"
    # if task has due date, set calendar event due date to task due date
    set_date_on_calendar_event(calendar_event, due_date)


def set_due_date_on_task(due_date, task):
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


def set_date_on_calendar_event(calendar_event, due_date):
    if due_date:
        # if due date is today
        if due_date == "Today":
            # set start date to today
            calendar_event.Start = datetime.datetime.today()
            # set due date to today
            calendar_event.AllDayEvent = True
            # set reminder date
            calendar_event.ReminderSet = True
        # if due date is tomorrow, set the start date to tomorror and whole day event
        elif due_date == "Tomorrow":
            # set due date to tomorrow
            calendar_event.Start = datetime.datetime.today() + datetime.timedelta(days=1)
            calendar_event.AllDayEvent = True
            # set reminder date
            calendar_event.ReminderSet = True
            # set reminder time to 9 AM tomorrow
            calendar_event.ReminderMinutesBeforeStart = 15
        # if due date is next week
        elif due_date == "Next Week":
            # set start date to next monday
            # caluclate how many days until next monday
            days_until_next_monday = 7 - datetime.datetime.today().weekday()
            # make calendar event the whole day on next monday
            next_monday = datetime.datetime.today() + datetime.timedelta(days=days_until_next_monday)
            calendar_event.Start = next_monday
            calendar_event.AllDayEvent = True
            # set reminder date
            calendar_event.ReminderSet = True
            # set reminder to next_monday
            calendar_event.ReminderMinutesBeforeStart = 15
    calendar_event.Save()


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

    # checkbox to create a new calendar event with the same subject as the task and the same color as the category
    create_calendar_event_checkbox = tk.Checkbutton(frame)
    create_calendar_event_checkbox.config(text="Create calendar event", bg="white")
    create_calendar_event_checkbox.grid(row=3, column=0, padx=10, pady=10)
    # variable to store the checkbox value
    create_calendar_event_var = tk.IntVar()
    create_calendar_event_checkbox.config(variable=create_calendar_event_var)

    # create a button to save the task
    save_button = tk.Button(frame)
    save_button.config(text="Save",
                       command=lambda: save_task(subject_entry.get(), category_var.get(), due_date_var.get(),
                                                 create_calendar_event_var.get(), popup),
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

    file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML files", "*.html")])
    # if file path is empty
    if file_path == "":
        # return
        return

    # open the html file with utf-8 encoding
    html_file = open(file_path, "w", encoding="utf-8")
    # if open was successful
    if not html_file:
        # show error message
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
    # open the html file in the default browser
    webbrowser.open(file_path)


def generate_dots_and_subjects(category, finished_date, html_file, subject, task):
    if task.Status == 2:
        html_file.write("<span style='color:green; font-size:40px'>✓</span> ")
    if category == "A":
        # write big dot and then subject and finished date
        html_file.write(
            "<span style='color:red; font-size:40px'>●</span> " + " " + subject + " - " + finished_date.strftime(
                "%d/%m/%Y") + "<br>")
    elif category == "B":
        html_file.write(
            "<span style='color:yellow; font-size:40px'>●</span> " + " " + subject + " - " + finished_date.strftime(
                "%d/%m/%Y") + "<br>")
    elif category == "C":
        html_file.write(
            "<span style='color:green; font-size:40px'>●</span>" + " " + subject + " - " + finished_date.strftime(
                "%d/%m/%Y") + "<br>")
    else:
        html_file.write(
            "<span style='color:gold; font-size:40px'>●</span>" + " " + subject + " - " + finished_date.strftime(
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
    popup.geometry("600x600")
    popup.resizable(False, False)

    # create white canvas that fills the popup window
    canvas = tk.Canvas(popup)
    canvas.config(bg="white")
    # make canvas fill height and width of popup window
    canvas.pack(fill=tk.BOTH, expand=True)

    # draw the help text
    canvas.create_text(300, 50, text="Help", fill="black", font=("Arial", 20), anchor=tk.CENTER)
    canvas.create_text(300, 100, text="Press ctrl+n to create a new task", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 130, text="Press ctrl+r to reload the tasks", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 160, text="Press ctrl+p to show project tasks", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 190, text="Press ctrl+t to show tasks finished today", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 220, text="Press ctrl+a to show only category A tasks", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 250, text="Press ctrl+b to show only category B tasks", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 280, text="Press ctrl+c to show only category C tasks", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 310, text="Press ctrl+f to search tasks", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 340, text="Press i to load Inbox", fill="black", font=("Arial", 12), anchor=tk.CENTER)
    canvas.create_text(300, 370, text="Press ctrl+g to generate html file", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 400, text="Right click dot to finish task", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 430, text="Double click task text to open it in Outlook", fill="black", font=("Arial", 12),
                       anchor=tk.CENTER)
    canvas.create_text(300, 460, text="Mouse wheel click task text or dot to start working timer", fill="black",
                       font=("Arial", 12),
                       anchor=tk.CENTER)


def load_inbox():
    # create new window
    inbox = Inbox()


# add generate html file button to ctrl+g

def load_note_input():
    # create new window
    note_input_window = tk.Toplevel(root)
    note = Note(note_input_window)
    note.run()
    # focus note window
    note.focus_note_content()


def add_menus():
    global menu_bar
    # add file menu to root
    menu_bar = tk.Menu(root)
    # add file menu to menu bar
    file_menu = tk.Menu(menu_bar, tearoff=0)
    # add file menu to menu bar
    menu_bar.add_cascade(label="File", menu=file_menu)
    #add sub menu for exporting stuff
    export_menu = tk.Menu(file_menu, tearoff=0)
    # add export menu to file menu
    file_menu.add_cascade(label="Export", menu=export_menu)
    # add export tasks with actual work to excel file
    export_menu.add_command(label="Tasks with actual work to excel", command=export_tasks_to_excel)
    # export tasks to sqlite database
    export_menu.add_command(label="Tasks to sqlite database", command=export_tasks_to_sqlite)
    # import tasks from sqlite database
    export_menu.add_command(label="Tasks from sqlite database", command=import_tasks_from_sqlite)
    # add generate html file to file menu
    export_menu.add_command(label="HTML-file with done tasks", command=lambda: generate_html_file())

    #add separator
    file_menu.add_separator()
    # add exit to file menu

    #Add new menu for showing tasks
    show_tasks_menu = tk.Menu(menu_bar, tearoff=0)
    #add show tasks menu to menu bar
    menu_bar.add_cascade(label="Show tasks", menu=show_tasks_menu)
    # add show tasks menu to file menu
    # add show all tasks to show tasks menu
    show_tasks_menu.add_command(label="Show all tasks", command=lambda: load_tasks(show_all_tasks=True))
    # add show tasks finished today to show tasks menu
    show_tasks_menu.add_command(label="Show tasks finished today", command=lambda: load_tasks(show_tasks_finished_today=True))
    # add show only category A to show tasks menu
    show_tasks_menu.add_command(label="Show only category A", command=lambda: load_tasks(show_only_this_category="A"))
    # add show only category B to show tasks menu
    show_tasks_menu.add_command(label="Show only category B", command=lambda: load_tasks(show_only_this_category="B"))
    # add show only category C to show tasks menu
    show_tasks_menu.add_command(label="Show only category C", command=lambda: load_tasks(show_only_this_category="C"))
    # add show only category C to show tasks menu
    show_tasks_menu.add_command(label="Show only project tasks", command=lambda: load_tasks(show_projects=True))

    # add file menu to load Inbox class
    file_menu.add_command(label="Load Inbox", command=lambda: load_inbox())
    # add file menu to load Note class
    file_menu.add_command(label="Load Note", command=lambda: load_note_input())
    file_menu.add_command(label="Exit", command=root.destroy)
    # bind CTRL+F to load search tasks
    root.bind("<Control-f>", lambda event: search_tasks_popup())
    # add new search menu
    search_menu = tk.Menu(menu_bar, tearoff=0)
    # add search menu to menu bar
    menu_bar.add_cascade(label="Search", menu=search_menu)
    # on click on search menu cascade open search tasks popup
    search_menu.add_command(label="Search tasks", command=lambda: search_tasks_popup())
    # add command to search for notes
    search_menu.add_command(label="Search notes", command=lambda: search_notes_popup())
    # add help menu
    help_menu = tk.Menu(menu_bar, tearoff=0)
    # add help menu to menu bar
    menu_bar.add_cascade(label="Help", menu=help_menu)
    # on click on help menu cascade open help window
    help_menu.add_command(label="Help", command=lambda: open_help_window())



add_menus()

# bind i to inbox
root.bind("i", lambda event: load_inbox())

# add file menu to root
root.config(menu=menu_bar)
# on focus or click or maximize reload the tasks
root.bind("<FocusIn>", lambda event: load_tasks())

root.mainloop()

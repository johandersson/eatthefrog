# import datetime with today
import datetime
import os
from tkinter import messagebox, filedialog, simpledialog, ttk
import tkinter as tk
import win32com.client
from inbox import Inbox
from tkcalendar import DateEntry

import xml.etree.ElementTree as ET

# set some constants for the drawing
dot_radius = 8  # radius of each dot
dot_gap = 10  # gap between each dot and task
text_gap = 20  # gap between each task and text
line_height = 30  # height of each line
root = tk.Tk()
no_date = datetime.datetime.strptime("01/01/4501", "%d/%m/%Y")
current_filter = None
# import Note class from note.py
from note import Note

tasks_by_category = []
task = None
category = None


def close_window(event=None):
    # destroy the root window
    root.destroy()


root.bind("<Escape>", close_window)

# make the window full screen from start
# root.attributes('-fullscreen', True)
root.config(bg="white")
root.title("Eat the frog (A-B-C)")
root.geometry("800x800")
tab_control = ttk.Notebook(root)

today_tab = ttk.Frame(tab_control)
tab_control.add(today_tab, text="Today")
tab_control.pack(expand=1, fill="both")
# on focus of tab Today, load tasks with category Today
# load a+b+c tasks
today_tab.bind("<FocusIn>", lambda event: load_tasks_in_correct_tab())
# bind click on tab to load tasks
# create separate tab for A


inbox_tab = ttk.Frame(tab_control)
# when clicking inbox_tab focus on inbox input field
inbox_tab.bind("<Button-1>", lambda event: capture_text.focus_set())
tab_control.add(inbox_tab, text="To do")
tab_control.pack(expand=1, fill="both")
# on focus of tab Inbox, load tasks with category Inbox
inbox_tab.bind("<FocusIn>", lambda event: load_tasks_in_correct_tab())
# separate tag for inbox Note

tasks_finished_today_tab = ttk.Frame(tab_control)
tab_control.add(tasks_finished_today_tab, text="Done today")
tab_control.pack(expand=1, fill="both")
# on focus of tab Tasks finished today, load tasks with category Tasks finished today
tasks_finished_today_tab.bind("<FocusIn>", lambda event: load_tasks_in_correct_tab())

tab_control.pack(expand=1, fill="both")
today_canvas = tk.Canvas(today_tab)
# when today_canvas gets focus, load tasks with category Today
today_canvas.bind("<FocusIn>", lambda event: load_tasks_in_correct_tab())
# when inbox canvas gets focus, load tasks with category Inbox

flagged_email_tab = ttk.Frame(tab_control)
tab_control.add(flagged_email_tab, text="Flagged emails")
tab_control.pack(expand=1, fill="both")
# on focus of tab Flagged email, load tasks with category Flagged email
flagged_email_tab.bind("<FocusIn>", lambda event: load_tasks_in_correct_tab())

flagged_email_canvas = tk.Canvas(flagged_email_tab)
# when flagged_email_canvas gets focus, load tasks with category Flagged email
flagged_email_canvas.bind("<FocusIn>", lambda event: load_tasks_in_correct_tab())
# make canvas fill the root window
flagged_email_canvas.pack(fill=tk.BOTH, expand=True)


# add a button below the canvas in a frame to create a new task
# create a frame to hold the widgets


def create_new_task(task_subject):
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    # new task
    new_task = outlook.CreateItem(3)
    # set task subject to task_subject
    new_task.Subject = task_subject
    # set task category to current_filter
    new_task.Categories = ""
    # set task start date to today
    return new_task


# function fo find all tasks with due date today
def find_tasks_with_due_date_today_and_set_to_a():
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    # loop through all the tasks and check their subject or body
    # new list to hold tasks with due date today
    tasks_with_due_date_today = []
    for task_item in tasks:
        # if task has due date today
        # convert task due date to a comparable format
        task_due_date = task_item.DueDate.strftime("%d/%m/%Y")
        today = datetime.datetime.today().strftime("%d/%m/%Y")
        if task_due_date == today:
            # if task has no category, set it to "A"
            if task_item.Categories == "":
                task_item.Categories = "A"


def create_new_task_from_entry(event, task_subject):
    # if task subject is not empty
    if task_subject != "":
        # create a new task
        new_task_from_entry = create_new_task(task_subject)
        new_task_from_entry.Save()
        # save task
        # load tasks in correct tab
        load_tasks_in_correct_tab()
        # delete text from capture text
        capture_text.delete(0, tk.END)


def add_inbox_input_field():
    global capture_text
    # add a textfield under the plus button that also fills x
    # add text label above text field
    capture_text_label = tk.Label(inbox_tab)
    capture_text_label.config(
        text="Write your thoughts and press enter to add to list of tasks for categorizing later:", bg="white")
    capture_text_label.pack(fill=tk.X)
    capture_text = tk.Entry(inbox_tab)
    capture_text.pack(fill=tk.X)
    # make capture text white background
    capture_text.config(bg="white")
    # make capture text font big
    capture_text.config(font=("Arial", 18))
    # make capture text placeholder

    # make capture text placeholder gray
    capture_text.config(fg="gray")

    capture_text.bind("<Return>", lambda event: create_new_task_from_entry(event, capture_text.get()))


add_inbox_input_field()

# make the canvas fill the root window
today_canvas.pack(fill=tk.BOTH, expand=True)
# add canvas to today tab
today_canvas.pack(fill=tk.BOTH, expand=True)
# add plus button under canvas


inbox_canvas = tk.Canvas(inbox_tab)
# make the canvas fill the root window
inbox_canvas.pack(fill=tk.BOTH, expand=True)
# add canvas to inbox tab
inbox_canvas.pack(fill=tk.BOTH, expand=True)
inbox_canvas.bind("<FocusIn>", lambda event: load_tasks_in_correct_tab())
tasks_finished_today_canvas = tk.Canvas(tasks_finished_today_tab)
# make the canvas fill the root window
tasks_finished_today_canvas.pack(fill=tk.BOTH, expand=True)
# add canvas to root
tasks_finished_today_canvas.pack(fill=tk.BOTH, expand=True)

# add status bar to root
status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# add scrollbar to canvas
scrollbar = tk.Scrollbar(inbox_canvas)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
# add scrollbar to today canvas
today_scrollbar = tk.Scrollbar(today_canvas)
today_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
# add scrollbar to a canvas

prev_item = None  # Keep track of the previous item


def highlight_text(event, canvas_x, task_text):
    global prev_item  # Keep track of the previous item
    # change font to underline
    # set cursor to hand
    canvas_x.itemconfig(task_text, fill="blue", font=("Arial", 12, "underline"))
    if prev_item and prev_item != task_text:  # If there was a previous item and it is different from the current item
        # ensure that the previous item is a text
        if canvas_x.type(prev_item) == "text":
            # Restore its color and font
            canvas_x.itemconfig(prev_item, fill="black", font=("Arial", 12, "normal"))
    prev_item = task_text  # Update the previous item


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
    # check if tasks already has TotalWork
    total_work = task.TotalWork
    if total_work is not None and total_work != 0:
        # show simple diaglog with the total work already set in input box, let the user change it
        estimated_work = simpledialog.askinteger("Set estimated work",
                                                 "How many minutes do you think this task will take?",
                                                 initialvalue=total_work)

    # ask user to set estimated work with simpledialog
    estimated_work = simpledialog.askinteger("Set estimated work", "How many minutes do you think this task will take?")
    # if user pressed cancel
    if estimated_work is None:
        # return
        return

    # set task TotalWork to estimated work
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
    # root title "Eat the frog (A-B-C)"
    popup.title("Timer")
    # make root window full screen
    popup.state("zoomed")
    # make popup white background
    popup.config(bg="white")
    # maximize popup and make it not and really fullscreen without the taskbar
    popup.wm_state('zoomed')
    # popup.overrideredirect(True)
    # disable x button
    popup.protocol("WM_DELETE_WINDOW", lambda: close_popup_and_save_time_in_task(event, popup, task_to_start_timer_on))

    # create a canvas to draw on
    canvas = tk.Canvas(popup)
    # make the canvas fill the root window
    canvas.pack(fill=tk.BOTH, expand=True)
    # make canvas white
    canvas.config(bg="white")
    # add status bar to root
    status_bar = tk.Label(popup, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    # add scrollbar to canvas
    # create a label to display the task subject
    subject_label = tk.Label(canvas)
    subject_label.config(text=task_to_start_timer_on.Subject, bg="white")
    # make subject label left center
    subject_label.pack()
    # make subject label fill x
    # subject_label.pack(fill=tk.X)
    # create a colored rectangle next to the subject label
    # get color from category
    color = get_color_code_from_category(task_to_start_timer_on.Categories)

    subject_label.pack(pady=20)

    timer_label = tk.Label(canvas)
    timer_label.config(text="25:00", bg="white", font=("Segoe UI", 50, "bold"))
    timer_label.pack()
    # create a button to stop the timer

    # make all labels white background
    subject_label.config(bg="white")

    timer_label.config(bg="white")

    # calculate x and y of the timer label

    # make a label with a rectangle under the subject label
    rectange_label = tk.Label(canvas)
    rectangle_char = "▬"
    rectange_label.config(text=rectangle_char)
    # make rectangle char big
    rectange_label.config(font=("Arial", 50))
    # make rectangle char the color of the category
    rectange_label.config(fg=color)
    # make background color also the color of the category
    rectange_label.config(bg=color)
    rectange_label.pack()

    # make subject label with big font
    subject_label.config(font=("Segoe UI", 40))
    # make subject label wrap
    subject_label.config(wraplength=1200)
    # calculate how many characters can fit in one line
    # make canvas white background
    canvas.config(bg="white")
    # make space between labels and the buttons under it
    rectange_label.pack(pady=20)
    stop_button = tk.Button(canvas)
    # set font of stop button to Segoe UI
    stop_button.config(font=("Segoe UI", 10))
    stop_button.config(text="Stop", bg="white",
                       command=lambda: close_popup_and_save_time_in_task(event, popup, task_to_start_timer_on))
    stop_button.pack()
    # make popup white background
    popup.config(bg="white")
    # on press escape, close this popup and focus on root
    popup.bind("<Escape>", lambda event: close_popup_and_save_time_in_task(event, popup, task_to_start_timer_on))

    # make space between labels and the buttons under it

    start_timer(timer_label, popup)
    # focus on popup


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
    # check for division by zero
    if task.TotalWork == 0:
        return 0
    # calculate percentage of task done

    percentage = int(task.ActualWork / task.TotalWork * 100)
    if percentage > 100:
        percentage = 100
        # calculate how much time was over estimated in hours
        over_estimated_time = str(datetime.timedelta(minutes=task.ActualWork - task.TotalWork))
        # add info to task body that task took more time than estimated
        task.Body += "\n\nThis task took more time than estimated:" + over_estimated_time

    return percentage


def close_popup_and_save_time_in_task(event, popup, task):
    # save worked time in a variable
    worked_time = popup.worked_time
    # convert worked time in seconds to minutes
    worked_time_minutes = worked_time // 60
    # if worked time is 0
    if worked_time_minutes < 1:
        worked_time_minutes = 1

    task.ActualWork = set_task_actual_work(task, worked_time_minutes)
    task.PercentComplete = calculate_percentage_of_task_done(task)
    # close popup and save time in task
    popup.destroy()
    # focus root
    root.focus_set()

    # update main window tasks
    load_tasks_in_correct_tab()
    # save task
    task.Save()


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
    # add label over listbox that says: "Search results, double click to open task in Outlook"

    search_results_listbox = tk.Listbox(frame, width=50, height=10)
    search_results_listbox.grid(row=3, column=0, padx=10, pady=10)
    # make search results listbox nice font
    search_results_listbox.config(font=("Tahoma", 10))

    # add search results label under listbox
    search_results_label = tk.Label(frame)
    search_results_label.config(text="", bg="white")
    search_results_label.grid(row=4, column=0, padx=10, pady=10)

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
    if subject == "" and body == "":
        # ask if the user wants to search for all tasks the choices "All finished tasks" and "All tasks" or "Cancel"
        search_all_tasks = messagebox.askyesnocancel("Search all tasks?",
                                                     "Do you want to search for all tasks?\n\nClick 'Yes' to search for all tasks.\nClick 'No' to search for all finished tasks.\nClick 'Cancel' to cancel.")
        if search_all_tasks is None:
            # close popup
            popup.destroy()
            # return
            return
        elif search_all_tasks:
            # search for all tasks
            tasks = tasks_folder.Items
        else:
            # search for all finished tasks
            tasks = tasks_folder.Items.Restrict("[Status] = 2")

        for task_item in tasks:
            search_results.append(task_item)

    else:
        for task_item in tasks:
            # get the category of the task
            # if task subject contains subject
            # is subject is not empty and in task subject, ignore case
            if subject != "" and subject.casefold() in task_item.Subject.casefold():
                # add task to list
                search_results.append(task_item)
            # is body is not empty and in task body, ignore case
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
        return None

    # display how many result in search results label, find label by searching for it in frame
    search_results_label = popup.winfo_children()[0].winfo_children()[6]
    search_results_label.config(
        text="Results: " + str(len(search_results)) + " tasks. Double click to open task in Outlook.")

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


def create_calendar_event(task, date, popup):
    # create a calendar event from task
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    # loop through all the tasks and check their subject or body
    # new calendar event
    calendar_event = outlook.CreateItem(1)
    # set calendar event subject to task subject
    calendar_event.Subject = task.Subject
    # set calendar event body to task body
    calendar_event.Body = task.Body
    # set to full day event
    calendar_event.AllDayEvent = True
    # compare task.StartDate to datetime.datetime(4501, 1, 1, 0, 0)
    # set date
    # convert date to fit pywin32
    outlook_date = date.strftime("%m/%d/%Y")
    calendar_event.Start = outlook_date
    # set category to task category
    calendar_event.Categories = task.Categories
    # save calendar event
    calendar_event.Save()
    # convert calendar_event to string date to display in messagebox
    event_date_str = calendar_event.Start.strftime("%d/%m/%Y")
    # info dialog to say that the calendar event with subject and date was created and ask if user wants to open it
    # in outlook
    if messagebox.askyesno("Calendar event created",
                           "Calendar event with subject '" + calendar_event.Subject + "' and date '" + event_date_str + "' was created.\n\nDo you want to open it in Outlook?"):
        # open calendar event in outlook
        calendar_event.Display()
    # close popup
    popup.destroy()


def create_calendar_event_from_task(task_to_create_event_from):
    # open popup window to choose a date
    popup = tk.Toplevel(root)
    popup.title("Choose date")
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)

    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)
    date_picker = DateEntry(frame)
    date_picker.config(date_pattern="dd/mm/yyyy")
    date_picker.pack()
    # save picked date in variable

    # create a button to create calendar event
    create_calendar_event_button = tk.Button(frame)
    create_calendar_event_button.config(text="Create calendar event", bg="white",
                                        command=lambda: create_calendar_event(task_to_create_event_from,
                                                                              date_picker.get_date(), popup))
    create_calendar_event_button.pack()
    color = get_color_code_from_category(task_to_create_event_from.Categories)
    canvas = tk.Canvas(frame)
    # fill frame with canvas
    # put canvas on top of frame above the labels and buttons and date picker
    canvas.pack(fill=tk.BOTH, expand=True)

    # draw a quadratic shape next to the label with the color of the task category
    canvas.create_rectangle(0, 0, 50, 50, fill=color, outline=color)
    # pack the canvas so that it fills the frame

    # draw subject next to color
    canvas.create_text(60, 25, text=task_to_create_event_from.Subject, anchor=tk.W)
    # if task has a Body draw that under the subject, but limit it to 50 characters and add ... at the end if it is
    # longer
    if task_to_create_event_from.Body != "":
        canvas.create_text(60, 50, text=task_to_create_event_from.Body[:50] + "...", anchor=tk.W)
    # make canvas white
    canvas.config(bg="white")


def delete_task(event, task):
    # show popup to with delete icon and ask if user wants to delete task with subject
    if messagebox.askyesno("Delete task?", "Are you sure you want to delete task with subject:\n'" + task.Subject + "'",
                           icon="warning"):
        task.Delete()
        load_tasks_in_correct_tab()


def change_priority_to_low(task):
    task.Importance = 0
    task.Save()
    load_tasks_in_correct_tab()


def change_priority_to_normal(task):
    task.Importance = 1
    task.Save()
    load_tasks_in_correct_tab()


def change_priority_to_high(task):
    task.Importance = 2
    task.Save()
    load_tasks_in_correct_tab()


def set_task_due_date(task, date, popup):
    # set task due date to date
    # convert date to fit pywin32
    outlook_date = date.strftime("%m/%d/%Y")
    task.DueDate = outlook_date
    task.Save()
    popup.destroy()
    load_tasks_in_correct_tab()


def add_due_date(event, task):
    # open a popup window with a date picker
    popup = tk.Toplevel(root)
    popup.title("Choose due date")
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)

    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)
    date_picker = DateEntry(frame)
    date_picker.config(date_pattern="dd/mm/yyyy")
    date_picker.pack()
    # save picked date in variable

    # create a button to create calendar event
    add_due_date_button = tk.Button(frame)
    add_due_date_button.config(text="Add due date", bg="white",
                               command=lambda: set_task_due_date(task, date_picker.get_date(), popup))
    add_due_date_button.pack()
    color = get_color_code_from_category(task.Categories)
    canvas = tk.Canvas(frame)
    # fill frame with canvas
    # put canvas on top of frame above the labels and buttons and date picker
    canvas.pack(fill=tk.BOTH, expand=True)

    # draw a quadratic shape next to the label with the color of the task category
    canvas.create_rectangle(0, 0, 50, 50, fill=color, outline=color)
    # pack the canvas so that it fills the frame

    # draw subject next to color
    canvas.create_text(60, 25, text=task.Subject, anchor=tk.W)
    # if task has a Body draw that under the subject, but limit it to 50 characters and add ... at the end if it is
    # longer
    if task.Body != "":
        canvas.create_text(60, 50, text=task.Body[:50] + "...", anchor=tk.W)
    # make canvas white
    canvas.config(bg="white")


def action_menu_popup(event, task_to_change):
    # highlight_text(event, canvas_x, task_text)
    # create popup menu
    popup_menu = tk.Menu(root, tearoff=0)
    # set title of popup men
    # add separator
    # get subject of task and cut if it is longer than 30 characters
    subject = task_to_change.Subject
    if len(subject) > 30:
        subject = subject[:30] + "..."

    # add a first item in the popup that acts as a title with bold label
    popup_menu.add_command(label=subject, font=("Arial", 9, "bold"))
    # disable the first item so that it can't be clicked
    popup_menu.entryconfig(0, state=tk.DISABLED)
    # add a separator
    popup_menu.add_separator()

    # add categories to popup menu as submenu
    categories_menu = tk.Menu(popup_menu, tearoff=0)
    categories_menu.add_command(label="A", command=lambda: change_category(task_to_change, "A"))
    categories_menu.add_command(label="B", command=lambda: change_category(task_to_change, "B"))
    categories_menu.add_command(label="C", command=lambda: change_category(task_to_change, "C"))
    # if is on inbox tab
    current_tab = tab_control.tab(tab_control.select(), "text")
    if current_tab == "To do":
        popup_menu.add_cascade(label="Add task to today", menu=categories_menu)
    else:
        popup_menu.add_cascade(label="Change priority", menu=categories_menu)
    # add choice to change priority
    # add submenu to the above change priority

    # add menu to create calendar event from task
    popup_menu.add_command(label="Create calendar event based on this task",
                           command=lambda: create_calendar_event_from_task(task_to_change))
    priority_menu = tk.Menu(popup_menu, tearoff=0)
    priority_menu.add_command(label="Low", command=lambda: change_priority_to_low(task_to_change))
    priority_menu.add_command(label="Normal", command=lambda: change_priority_to_normal(task_to_change))
    priority_menu.add_command(label="High", command=lambda: change_priority_to_high(task_to_change))

    # add option to rename task Subject
    popup_menu.add_command(label="Rename task", command=lambda: rename_task(event, task_to_change))
    # delete task
    popup_menu.add_command(label="Delete task", command=lambda: delete_task(event, task_to_change))
    # move single task to inbox
    if task_to_change.Categories != "":
        popup_menu.add_command(label="Move to To do", command=lambda: move_single_task_to_inbox(task_to_change))
    else:
        popup_menu.add_cascade(label="Change Outlook priority", menu=priority_menu)
        # add choice to add due date
        popup_menu.add_command(label="Add due date", command=lambda: add_due_date(event, task_to_change))
    # add menu to start a timer on the task
    popup_menu.add_command(label="Start timer", command=lambda: open_timer_window(event, task_to_change))

    # add menu to open task in outlook
    popup_menu.add_command(label="Open task in Outlook", command=lambda: task_to_change.Display())

    # display popup menu
    popup_menu.tk_popup(event.x_root, event.y_root)


def change_category(task, category):
    # set task category to category
    task.Categories = category
    # save task
    task.Save()
    # update tasks
    load_tasks_in_correct_tab()


def draw_tasks(canvas_to_draw_on):
    loading_icon()
    canvas_to_draw_on.delete("all")
    # as a first object before the tasks draw a + sign to the right
    # if we are on the today tab

    if len(tasks_by_category) == 0:
        selected_tab = tab_control.tab(tab_control.select(), "text")
        if selected_tab == "Done today":
            canvas_to_draw_on.create_text(350, 70, text="🐸", fill="green", font=("Segoe UI", 50), anchor=tk.CENTER)
            canvas_to_draw_on.create_text(350, 200, text="No tasks marked done today", fill="green",
                                          font=("Segoe UI", 30),
                                          anchor=tk.CENTER)
            # small text in Segoe UI with the same color that says: If you flag an email it will show up here
            canvas_to_draw_on.create_text(350, 250,
                                          text="Yet. Mark a task done by clicking the colored box left to the subject.",
                                          fill="green",
                                          font=("Segoe UI", 10),
                                          anchor=tk.CENTER)
            reset_loading_icon()
            return
        canvas_to_draw_on.create_text(350, 200, text="✓", fill="green", font=("Arial", 100), anchor=tk.CENTER)
        if current_filter is None or current_filter == "" or tab_control.tab(tab_control.select(),
                                                                             "text") != "Done today":
            canvas_to_draw_on.create_text(350, 300, text="No tasks", fill="green", font=("Arial", 30), anchor=tk.CENTER)
        else:

            canvas_to_draw_on.create_text(350, 300, text="No " + current_filter + " tasks", fill="green",
                                          font=("Arial", 30),
                                          anchor=tk.CENTER)
        # draw a text under the check mark saying "Press ctrl+n to create a new task"
        canvas_to_draw_on.create_text(350, 350, text="Press Ctrl+N to create one or several new tasks", fill="green",
                                      font=("Arial", 15),
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

        color = get_color_code_from_category(category)

        # if category is something else than A, B or C or empty
        if color == "#FFC0CB" and category != "":
            # draw a small text under the task text saying the category
            canvas_to_draw_on.create_text(x3, y3 + 12, text="Category:" + category, fill="black", anchor=tk.W,
                                          font=("Arial", 8))

        # draw a box on the left side of the task if it is the today view otherwhise a circle
        if tab_control.tab(tab_control.select(), "text") == "Today" or tab_control.tab(tab_control.select(),
                                                                                       "text") == "Done today":
            check_box = canvas_to_draw_on.create_rectangle(x1 - dot_radius, y1 - dot_radius, x1 + dot_radius,
                                                           y1 + dot_radius, fill=color, outline=color)
        else:
            check_box = canvas_to_draw_on.create_oval(x1 - dot_radius, y1 - dot_radius, x1 + dot_radius,
                                                      y1 + dot_radius, fill=color, outline=color)

        # set checkbox cursor to hand but only when hovering over the checkbox
        canvas_to_draw_on.tag_bind(check_box, "<Enter>", lambda event: canvas_to_draw_on.config(cursor="hand2"))
        canvas_to_draw_on.tag_bind(check_box, "<Leave>", lambda event: canvas_to_draw_on.config(cursor="arrow"))
        # when clicking the checkbox draw a check mark inside it
        canvas_to_draw_on.tag_bind(check_box, "<Button-1>", lambda event, task=task: mark_done(event, task, check_box))

        # mark done on left click
        canvas_to_draw_on.tag_bind(check_box, "<Button-1>", lambda event, task=task: mark_done(event, task))
        # when double clicking the dot open the task in outlook
        canvas_to_draw_on.tag_bind(check_box, "<Double-Button-1>", lambda event, task=task: task.Display())
        # draw a text with the task information on the canvas with black color and Segoe UI font size 12 if it is installed otherwise Arial
        task_text = canvas_to_draw_on.create_text(x3, y3, text=task_info, fill="black", anchor=tk.W,
                                                  font=("Segoe UI", 12))


        canvas_to_draw_on.tag_bind(task_text, "<Button-3>", lambda event, task=task: action_menu_popup(event, task))

        # when mouse wheel clicking the task text open a timer window
        canvas_to_draw_on.tag_bind(task_text, "<Button-2>", lambda event, task=task: open_timer_window(event, task))
        # when mouse over the text call the highlight_task function

        # if task is done draw a green check mark inside the box
        if task.Status == 2:
            check_mark = canvas_to_draw_on.create_text(x1, y1, text="✓", fill="light green", font=("Arial", 12),
                                                       anchor=tk.CENTER)
            # bind check mark to mark_done function
            canvas_to_draw_on.tag_bind(check_mark, "<Button-1>",
                                       lambda event, task=task: mark_done(event, task, check_mark))

        # if task is drawn out of the canvas
        if y3 > 800:
            # show the scrollbar
            canvas_to_draw_on.config(scrollregion=canvas_to_draw_on.bbox("all"))
            # bind the scrollbar to the canvas
            canvas_to_draw_on.config(yscrollcommand=scrollbar.set)
            scrollbar.config(command=canvas_to_draw_on.yview)
            # bind the mouse wheel to the scrollbar
            canvas_to_draw_on.bind_all("<MouseWheel>",
                                       lambda event: canvas_to_draw_on.yview_scroll(int(-1 * (event.delta / 120)),
                                                                                    "units"))
            # bind the up and down arrow keys to the scrollbar
            canvas_to_draw_on.bind_all("<Up>", lambda event: canvas_to_draw_on.yview_scroll(-1, "units"))
            canvas_to_draw_on.bind_all("<Down>", lambda event: canvas_to_draw_on.yview_scroll(1, "units"))
            # when dragging scrollbar with mouse scroll the canvas
            scrollbar.bind("<B1-Motion>", lambda event: canvas_to_draw_on.yview_moveto(event.y))
        # else hide the scrollbar
        else:
            canvas_to_draw_on.config(scrollregion=(0, 0, 0, 0))

    reset_loading_icon()


def get_color_code_from_category(task_category):
    if task_category == "A":
        color = "#ED2939"
    elif task_category == "B":
        color = "#FFFD74"
    elif task_category == "C":
        color = "#32de84"
    else:
        color = "#FFC0CB"
    return color


def loading_icon(window=root):
    window.config(cursor="watch")
    window.update()


def reset_loading_icon(window=root):
    window.config(cursor="")
    window.update()


def draw_tasks_with_icons(canvas_to_draw_on, task, x3, y3):
    # if task has due date and due date is in the current year or the next
    if task.DueDate and task.DueDate.year in [datetime.datetime.today().year, datetime.datetime.today().year + 1] \
            and task.Categories != "A" and task.Categories != "B" and task.Categories != "C":
        due_date = datetime.datetime.strptime(task.DueDate.strftime("%d/%m/%Y"), "%d/%m/%Y")
        # draw due date in small font right under the task text, if the task is due, make it red
        if due_date < datetime.datetime.today():
            text_color = "red"
        else:
            text_color = "grey"

        canvas_to_draw_on.create_text(x3, y3 + 15, text=task.DueDate.strftime("%d/%m/%Y"), fill=text_color,
                                      font=("Arial", 7),
                                      anchor=tk.W)

        # if task is complete draw a check mark on the right side of the due date text
        if task.Status == 2:
            canvas_to_draw_on.create_text(x3 + 310, y3, text="✓", fill="green", font=("Arial", 12), anchor=tk.W)


def load_tasks(canvas_to_draw_on, show_tasks_finished_today=False, show_all_tasks=False,
               show_only_this_category=None, draw=True, first_time_loading_tasks=False, show_all_finished_tasks=False):
    # clear status bar
    status_bar.config(text="")
    global tasks_by_category, task, category, current_filter
    current_filter = show_only_this_category
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    tasks.Sort("[CreationTime]", True)
    today = datetime.date.today().strftime("%m/%d/%Y")

    if show_tasks_finished_today:
        # restrict tasks to tasks finished today
        tasks = tasks.Restrict(
            "[Complete] = True AND [DateCompleted] >= '" + today + "'")
        # show message in status bar that only tasks finished today are shown
    # if task are not finished today but not show all tasks
    if show_all_finished_tasks:
        # restrict tasks to tasks finished today
        tasks = tasks.Restrict(
            "[Complete] = True")
        # show message in status bar that only tasks finished today are shown

    elif not show_all_tasks and not show_tasks_finished_today:
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
            status_bar.config(text="Showing all tasks, including completed tasks")
            # add task to list
            tasks_by_category.append((task, category))
            continue
        if show_only_this_category is None:
            if category == "A" or category == "B" or category == "C":
                if show_tasks_finished_today:
                    status_bar.config(text="Showing A, B, and C tasks finished today")
                else:
                    status_bar.config(text="Showing A, B, and C tasks")
                # add task to list
                tasks_by_category.append((task, category))
        else:
            if show_only_this_category == "No category" and category != "A" and category != "B" and category != "C":
                status_bar.config(text="Showing tasks that are not prioritized with A-B-C")
                tasks_by_category.append((task, category))
            # add task to list

    # sort the list by category in ascending order
    tasks_by_category.sort(key=lambda x: x[1])
    # inside category sort by priority

    if show_only_this_category == "No category":
        # sort tasks by priority in outlook, high priority first
        tasks_by_category.sort(key=lambda x: x[0].Importance, reverse=True)
    # check how many tasks are in the list and show a warning if there are more than 20 tasks, and then ask the user
    # if he wants to see the warning next time and save his anser in a config fil

    if draw:
        canvas_to_draw_on = get_canvas_to_draw_on()
        draw_tasks(canvas_to_draw_on)

    return tasks_by_category


def mark_email_as_completed(email):
    # mark email as completed
    email.FlagStatus = 1
    email.Save()
    draw_flagged_emails(get_canvas_to_draw_on())


def unflag_email(email):
    # unflag email
    email.FlagStatus = 0
    email.Save()
    draw_flagged_emails(get_canvas_to_draw_on())


def create_task_from_email(email):
    # create a new task
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    task_from_email = outlook.CreateItem(3)
    # set task subject to email subject
    task_from_email.Subject = email.Subject
    # set task body to email body
    task_from_email.Body = email.Body
    # set task categories to email categories
    task_from_email.Categories = email.Categories
    # save task
    task_from_email.Save()
    # ask user if he also wants to remove the flag from the email
    if messagebox.askyesno("Remove flag from email?",
                           "A task was created in To do. Do you also want to remove the flag from the email?"):
        unflag_email(email)
    load_tasks_in_correct_tab()


def uncomplete(email):
    # mark email as not completed
    email.FlagStatus = 0
    email.Save()
    draw_flagged_emails(get_canvas_to_draw_on())


def add_popup_menu_to_flag(event, email):
    popup_menu = tk.Menu(root, tearoff=0)
    popup_menu.add_command(label="Unflag email", command=lambda: unflag_email(email))
    popup_menu.add_command(label="Mark as completed", command=lambda: mark_email_as_completed(email))
    # add popup menu item to create task from Email subject
    popup_menu.add_command(label="Create task from email subject", command=lambda: create_task_from_email(email))
    # display popup menu
    popup_menu.tk_popup(event.x_root, event.y_root)


def draw_flagged_emails(canvas_to_draw_on):
    loading_icon()
    # clear
    canvas_to_draw_on.delete("all")
    # find emails that have been flagged
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for emails
    emails_folder = namespace.GetDefaultFolder(6)
    # get all the emails in the folder
    emails = emails_folder.Items
    flagged_emails = []
    for flagged_email in emails:
        if flagged_email.FlagRequest != "":
            flagged_emails.append(flagged_email)
    # sort the list by category in ascending order
    flagged_emails.sort(key=lambda x: x.FlagRequest)

    if len(flagged_emails) == 0:
        canvas_to_draw_on.create_text(350, 70, text="🐸", fill="green", font=("Segoe UI", 50), anchor=tk.CENTER)
        canvas_to_draw_on.create_text(350, 200, text="No flagged emails", fill="green", font=("Segoe UI", 30),
                                      anchor=tk.CENTER)
        # small text in Segoe UI with the same color that says: If you flag an email it will show up here
        canvas_to_draw_on.create_text(350, 250, text="If you flag an email in your Outlook inbox it will show up here",
                                      fill="green",
                                      font=("Segoe UI", 10),
                                      anchor=tk.CENTER)
        reset_loading_icon()
        return

    for i, flagged_email in enumerate(flagged_emails):
        # get the email subject and due date
        subject = flagged_email.Subject
        # draw a flag with the color of the flagged email and then the subject and bind it to open the email in outlook and left align the text
        flag = canvas_to_draw_on.create_text(50, (i + 1) * line_height, text="🚩", fill="red", font=("Segoe UI", 12),
                                      anchor=tk.W)
        subject_text = canvas_to_draw_on.create_text(70, (i + 1) * line_height, text=subject, fill="black", font=("Segoe UI", 12),
                                      anchor=tk.W)
        # if email is completed draw a check mark next to the flag
        if flagged_email.FlagStatus == 1:
            canvas_to_draw_on.create_text(30, (i + 1) * line_height, text="✓", fill="green", font=("Segoe UI", 12),
                                          anchor=tk.W)

        #make variable that holds both flag and text for bind to clicking
        # when double clicking the dot open the task in outlook avoid late binding


        #make variable that holds both flag and text for bind to clicking
        # when double clicking the dot open the task in outlook avoid late binding
        canvas_to_draw_on.tag_bind(subject_text, "<Double-Button-1>", lambda event, email=flagged_email: email.Display())
        # when right clicking the dot open a popup menu
        canvas_to_draw_on.tag_bind(subject_text, "<Button-3>",
                                    lambda event, email=flagged_email: add_popup_menu_to_flag(event, email))
        canvas_to_draw_on.tag_bind(flag, "<Double-Button-1>",
                                   lambda event, email=flagged_email: email.Display())
        # when right clicking the dot open a popup menu
        canvas_to_draw_on.tag_bind(flag, "<Button-3>",
                                   lambda event, email=flagged_email: add_popup_menu_to_flag(event, email))

    reset_loading_icon()


# function to change cursor to a hand when any task is hovered over
def change_cursor_to_hand(event):
    # get type of widget that the mouse is over
    widget = event.widget.find_withtag(tk.CURRENT)
    # if the widget is a task text
    if widget and widget[0] == 2:
        # change cursor to hand
        root.config(cursor="hand2")
    else:
        root.config(cursor="")


# bind hover over task_text to change_cursor_to_hand function


def get_canvas_to_draw_on():
    canvas_to_draw_on = None
    chosen_tab = tab_control.index(tab_control.select())
    if chosen_tab == 0:
        canvas_to_draw_on = today_canvas
    elif chosen_tab == 1:
        canvas_to_draw_on = inbox_canvas
    elif chosen_tab == 2:
        canvas_to_draw_on = tasks_finished_today_canvas
    elif chosen_tab == 3:
        canvas_to_draw_on = flagged_email_canvas
    return canvas_to_draw_on


# function to move tasks to inbox with a popup window with a listbox where you can multi select tasks and then click a button to move them to the inbox
def move_tasks_to_inbox(tasks_by_category):
    # create popup window
    popup = tk.Toplevel(root)
    popup.title("Move tasks to Todo")
    popup.config(bg="white")
    popup.geometry("600x400")
    popup.resizable(False, False)
    # set font of popup window to be Segoe UI

    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    # make frame as big as the window
    frame.pack(fill=tk.BOTH, expand=True)
    label = tk.Label(frame)
    label.config(text="Select one or multiple tasks to move to To do by clicking them in the list", bg="white")
    label.pack()

    # create a listbox to hold the tasks
    listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE)
    listbox.config(bg="white", fg="black")
    # make listbox as big as the window except for the buttons under it
    listbox.pack(fill=tk.BOTH, expand=True)
    # set font of listbox to be Segoe UI
    listbox.config(font=("Segoe UI", 12))

    # create a scrollbar for the listbox
    scrollbar = tk.Scrollbar(listbox)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # bind scrollbar to listbox
    listbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox.yview)

    # add tasks to listbox
    for task, category in tasks_by_category:
        listbox.insert(tk.END, task.Subject)

    # create a button to move tasks to inbox
    move_to_inbox_button = tk.Button(frame)
    move_to_inbox_button.config(text="Move tasks to To do", bg="white",
                                command=lambda: move_tasks_to_inbox_command(listbox, popup))
    move_to_inbox_button.pack()

    # create a button to close the popup
    close_button = tk.Button(frame)
    # make space between buttons
    close_button.config(pady=10)

    close_button.config(text="Cancel", bg="white", command=popup.destroy)
    close_button.pack()
    load_tasks_in_correct_tab()


# move tasks to inbox command
def move_tasks_to_inbox_command(listbox, popup):
    # get selected tasks
    selected_tasks = listbox.curselection()
    # create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")
    # get the namespace object
    namespace = outlook.GetNamespace("MAPI")
    # get the default folder for tasks
    tasks_folder = namespace.GetDefaultFolder(13)
    # get all the tasks in the folder
    tasks = tasks_folder.Items
    # loop through selected tasks
    for task in selected_tasks:
        # get task
        task = tasks_by_category[task][0]
        # set category to no category
        task.Categories = ""
        # save task
        task.Save()
    # destroy popup
    # Message info dialog to show that tasks were moved to inbox and how many
    popup.destroy()
    # reload tasks
    load_tasks_in_correct_tab()
    messagebox.showinfo("Tasks moved to To do", str(len(selected_tasks)) + " tasks were moved to To do")


def move_single_task_to_inbox(task):
    task.Categories = ""
    task.Save()
    load_tasks_in_correct_tab()
    # Message box with title "Task moved to inbox" and text "Task with subject x was moved to To do"
    messagebox.showinfo("Task moved to To do",
                        "Task with subject:\n" + "'" + task.Subject + "'" + "\nwas moved to To do")


# create a root window for the GUI


def mark_done(event, task, check_mark=None):
    canvas_to_draw_on = get_canvas_to_draw_on()
    # if canvas to draw on is the todo view
    if canvas_to_draw_on != inbox_canvas:
        if task.Status == 2:
            task.Status = 1
        else:
            task.Status = 2

        item = canvas_to_draw_on.find_withtag("current")
        # frame the item with a black border
        # if task is not already done
        if task.Status == 2:
            # draw a small green check mark a little bit bigger then the item, inside the item
            canvas_to_draw_on.create_text(canvas_to_draw_on.bbox(item)[0] + 1, canvas_to_draw_on.bbox(item)[1] + 6,
                                          text="✓", fill="light green",
                                          font=("Arial", 18), anchor=tk.W)
        else:
            # delete the check mark from item
            # remove check_mark
            canvas_to_draw_on.delete(check_mark)

        task.Save()
        # call load_tasks_in_correct_tab function after 300 milliseconds
        canvas_to_draw_on.after(3000, load_tasks_in_correct_tab)
    else:
        # move task to today with category A
        task.Categories = "A"
        task.Save()
        load_tasks_in_correct_tab()


def save_task(subject, category, due_date=None, create_calendar_event=False, date_var=None, popup=None):
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
    if category != "No category":
        task.Categories = category
    # set due date to tomorrow
    # get tomorrow's date
    # check if create calendar event is checked
    if create_calendar_event:
        create_new_calendar_event_based_on_task(category, date_var, namespace, subject)
    set_due_date_on_task(due_date, task)

    # if task subject is empty, ask user for subject line
    if task.Subject == "":
        # you didn't add any subject line, do you want to add one?
        # info dialog, please add subject line
        messagebox.showinfo("Empty subject", "Please add subject line")
        if popup:
            # focus subject entry
            # find subject entry in popup
            popup.focus_set()
    else:
        # save the task in the tasks folder
        task.Save()
        load_tasks_in_correct_tab()
        if popup:
            # destroy the popup
            popup.destroy()


def load_tasks_in_correct_tab():
    chosen_tab = tab_control.index(tab_control.select())
    if chosen_tab == 0:
        load_tasks(canvas_to_draw_on=today_canvas)
    elif chosen_tab == 1:
        load_tasks(canvas_to_draw_on=inbox_canvas, show_only_this_category="No category")
    elif chosen_tab == 2:
        load_tasks(canvas_to_draw_on=tasks_finished_today_canvas, show_tasks_finished_today=True)
    else:
        draw_flagged_emails(canvas_to_draw_on=flagged_email_canvas)


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


    # if due date is tomorrow
    elif due_date == "Tomorrow":
        # set start date to today
        task.StartDate = datetime.datetime.today()
        # set due date to tomorrow
        task.DueDate = datetime.datetime.today() + datetime.timedelta(days=1)

    # if due date is next week
    elif due_date == "Next Week":
        # set start date to today
        task.StartDate = datetime.datetime.today()
        # set due date to tomorrow
        task.DueDate = datetime.datetime.today() + datetime.timedelta(days=7)


def set_date_on_calendar_event(calendar_event, due_date):
    if due_date:
        # if due date is DateTime
        if type(due_date) == datetime.date:
            # covert datetime.date to pywin32 datetime
            due_date = datetime.datetime.combine(due_date, datetime.datetime.min.time())

            # set due date to local timezone
            calendar_event.Start = due_date
            # all day
            calendar_event.AllDayEvent = True
            # set reminder date
            calendar_event.ReminderSet = True
            # set reminder time to 9 AM tomorrow
            calendar_event.ReminderMinutesBeforeStart = 15

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
def create_new_task_popup(default_category="No category"):
    if default_category is None:
        default_category = "No category"
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Add new task to Today")
    popup.config(bg="white")
    popup.geometry("400x500")
    ##set font on popup to Segoe UI

    # create a frame to hold the widgets
    frame = tk.Frame(popup)
    frame.config(bg="white")
    frame.pack(fill=tk.BOTH)

    # create a label to display the task subject
    subject_label = tk.Label(frame)
    subject_label.config(text="Subject:", bg="white", font=("Segoe UI", 10))
    subject_label.grid(row=0, column=0, padx=10, pady=10)

    # create an entry to get the task subject
    subject_entry = tk.Entry(frame)
    subject_entry.config(width=40, bg="white", font=("Segoe UI", 10))
    subject_entry.grid(row=0, column=1, padx=10, pady=10)

    # make category dropdown next to subject entry
    category_label = tk.Label(frame)
    category_label.config(text="Category:", bg="white", font=("Segoe UI", 10))
    category_label.grid(row=1, column=0, padx=10, pady=10)

    # create a variable to hold the category
    category_var = tk.StringVar()
    # set the default category
    category_var.set("A")
    # create a dropdown menu to select the category
    category_dropdown = tk.OptionMenu(frame, category_var, "A", "B", "C")
    category_dropdown.config(width=10, bg="white", font=("Segoe UI", 10))
    category_dropdown.grid(row=1, column=1, padx=10, pady=10)

    # create a button to save the task
    save_button = tk.Button(frame)
    save_button.config(text="Save",
                       command=lambda: save_task(subject_entry.get(), category_var.get(), "Today",
                                                 popup=popup),
                       bg="white")

    # create a button to cancel the task
    cancel_button = tk.Button(frame)
    cancel_button.config(text="Cancel", command=popup.destroy, bg="white")
    # show save and cancel buttons at the bottom right corner
    save_button.grid(row=4, column=1)
    cancel_button.grid(row=4, column=2)
    # align buttons to the left
    frame.grid_columnconfigure(0, weight=1)
    # focus on subject entry
    subject_entry.focus_set()
    # bind enter key to save button
    popup.bind("<Return>", lambda event: save_task(subject_entry.get(), category_var.get(), due_date_var.get(), popup))


def switch_to_tab(tab_index):
    tab_control.select(tab_index)
    load_tasks_in_correct_tab()


canvas = get_canvas_to_draw_on()
root.bind("<Control-r>", lambda event: load_tasks_in_correct_tab())
# bind ctrl+´m to move tasks to inbox
root.bind("<Alt-m>", lambda event: move_tasks_to_inbox(tasks_by_category))

root.bind("<Control-t>", lambda event: switch_to_tab(5))
# bind Control+A to only show A tasks on today tab
root.bind("<Control-a>", lambda event: load_tasks(canvas_to_draw_on=today_canvas, show_only_this_category="A"))
# bind Control+B to only show B tasks on today tab
root.bind("<Control-b>", lambda event: load_tasks(canvas_to_draw_on=today_canvas, show_only_this_category="B"))
# bind Control+C to only show C tasks on today tab
root.bind("<Control-c>", lambda event: load_tasks(canvas_to_draw_on=today_canvas, show_only_this_category="C"))
# bind Control+N to load_inbox
root.bind("<Control-n>", lambda event: load_inbox())


def generate_html_file_to_print():
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

    # write the html file header with Todo list for today and today's date as title
    html_file.write("<html><head><title>Todo list</title></head><body>")
    # load nice google font
    html_file.write("<link href='https://fonts.googleapis.com/css?family=Roboto' rel='stylesheet'>")
    # make the font of the html file Roboto
    html_file.write("<style>body {font-family: 'Roboto', sans-serif;}</style>")
    # make that font the default font for the html file
    # loop through the tasks and write the html file
    # write Todo list and date in 12 december 2020 format
    html_file.write("<h1>Todo list</h1>")
    html_file.write("<h2>" + datetime.datetime.today().strftime("%d %B %Y") + "</h2>")

    for task, category in load_tasks(canvas_to_draw_on=None, draw=False):
        # if task is completed
        # write the task subject and finished date
        # if task is category A, B, or C, draw a dot with the color of the category

        if task.Status != 2 and task.DateCompleted != no_date:
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
    # add print javascript button
    # center button
    html_file.write("<center>")
    html_file.write("<button onclick='window.print()'>Print</button>")
    html_file.write("</center>")
    html_file.write("</body></html>")
    # close the html file
    html_file.close()
    # send html_file to printer, ask user if he wants to print it
    # open the html file
    import webbrowser
    # open the html file in the default browser
    webbrowser.open(file_path)


def generate_dots_and_subjects(category, finished_date, html_file, subject, task):
    if task.Status == 2:
        html_file.write("<span style='color:green; font-size:40px'>✓</span>")

    html_file.write(generate_html_task_line(task))
    # if the task has a body print it in small grey nice formatted letters under the task subject with space before and after
    if task.Body is not None:
        html_file.write(
            "<span style='color:grey; font-size:10px; margin-left: 20px; margin-right: 20px;'>" + task.Body + "</span><br>")


def generate_html_task_line(task):
    color = get_color_code_from_category(task.Categories)
    # if task has no date
    if task.DateCompleted is None or task.DateCompleted != datetime.datetime.strptime("01/01/4501", "%d/%m/%Y"):
        # return task without date with color from category
        return "<span style='color:" + color + "; font-size:40px'>●</span>" + " " + task.Subject + " (" + task.Categories + ")<br>"
    else:
        # return task with date with color from category
        return "<span style='color:" + color + "; font-size:40px'>●</span>" + " " + task.Subject + " - " + finished_date.strftime(
            "%d/%m/%Y") + "<br>"


def open_licenses_file():
    # open licenses.txt file with complete file path and utf-8 encoding
    os.startfile(os.path.join(os.path.dirname(__file__), "licenses.txt"))


def open_help_window():
    # open help window where a uneditable text widget shows the help text, the widget is flat and shows no border and is scrollable
    # create a popup window
    popup = tk.Toplevel(root)
    popup.title("Help")
    popup.config(bg="white")
    popup.geometry("800x500")
    popup.resizable(False, False)

    # read text from help.txt file with complete file path and utf-8 encoding
    help_text_to_show = open(os.path.join(os.path.dirname(__file__), "help.txt"), "r", encoding="utf-8").read()

    # create a canvas that fills the frame
    help_canvas = tk.Canvas(popup)
    help_canvas.config(bg="white")
    help_canvas.pack(fill=tk.BOTH, expand=True)

    # for each line in the help.txt file, draw a text on the canvas with Sego UI font and size 10
    for line in help_text_to_show.splitlines():
        # create new text object for each line and increase the y coordinate by 20 for each line
        help_canvas.create_text(10, 10 + 20 * help_text_to_show.splitlines().index(line), text=line,
                                font=("Segoe UI", 12), anchor=tk.NW)

    # create a scrollbar for the canvas
    scrollbar = tk.Scrollbar(help_canvas)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    # bind scrollbar to canvas
    help_canvas.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=help_canvas.yview)

    # create a button to close the popup
    close_button = tk.Button(help_canvas)
    close_button.config(text="Close", command=popup.destroy, bg="white")
    # show save and cancel buttons at the bottom right corner
    close_button.pack(side=tk.BOTTOM, padx=10, pady=10)

    licenses_button = tk.Button(help_canvas)
    licenses_button.config(text="Licenses", command=open_licenses_file, bg="white")
    licenses_button.pack(side=tk.BOTTOM, padx=10, pady=10)


def load_inbox():
    # create new inbox window with root as parent
    # create new inbox object
    inbox_window = Inbox(root, load_tasks_in_correct_tab)


# add generate html file button to ctrl+g

def load_note_input():
    # create new window
    note = Note()


# open message input box to rename a task Subject
def rename_task(event, task):
    # input box to rename task
    input_box = tk.simpledialog.askstring("Rename task", "Rename task", initialvalue=task.Subject)
    # if input box is not empty
    if input_box != "":
        # set task subject to input box
        task.Subject = input_box
        # save the task
        task.Save()
        # load tasks in correct tab
        load_tasks_in_correct_tab()
    else:
        # show error message
        messagebox.showerror("Error", "Please add a new subject")


def rename_task_subject(task, new_subject, popup):
    # if new subject is not empty
    if new_subject != "":
        # set task subject to new subject
        task.Subject = new_subject
        # save the task
        task.Save()
        # destroy the popup
        popup.destroy()
        # load tasks in correct tab
        load_tasks_in_correct_tab()
    else:
        # show error message
        messagebox.showerror("Error", "Please add a new subject")


def init_categories():
    # Create an Outlook application object
    outlook = win32com.client.Dispatch("Outlook.Application")

    # Get the default namespace
    namespace = outlook.GetNamespace("MAPI")

    # Get the categories collection
    categories = namespace.Categories

    # Define a list of category names and colors
    category_list = [
        ("A", 1),  # Red
        ("B", 4),  # Yellow
        ("C", 5),  # Green
    ]

    # Define a function to check and create categories
    def check_and_create_categories(category_list):
        # Loop through the category list
        for category_name, category_color in category_list:
            # Check if the category exists
            category = categories.Item(category_name)
            if category is None:
                # Create a new category
                categories.Add(category_name, category_color)

    # Call the function
    check_and_create_categories(category_list)


init_categories()


def add_menus():
    global menu_bar
    # add file menu to root
    menu_bar = tk.Menu(root)
    # add file menu to menu bar
    file_menu = tk.Menu(menu_bar, tearoff=0)
    # add file menu to menu bar
    menu_bar.add_cascade(label="File", menu=file_menu)

    # add exit to file menu

    # add file menu to load Inbox class
    file_menu.add_command(label="Add one or several tasks", command=lambda: load_inbox())
    file_menu.add_command(label="Move tasks to To do",
                          command=lambda: move_tasks_to_inbox(load_tasks(get_canvas_to_draw_on(), draw=False)))
    file_menu.add_command(label="Exit", command=root.destroy)
    # add choice to move tasks to inbox
    # bind CTRL+F to load search tasks
    root.bind("<Control-f>", lambda event: search_tasks_popup())
    search_menu = tk.Menu(menu_bar, tearoff=0)
    # add search menu to menu bar
    menu_bar.add_cascade(label="Search", menu=search_menu)
    # on click on search menu cascade open search tasks popup
    search_menu.add_command(label="Search tasks", command=lambda: search_tasks_popup())
    # add command to search for notes
    # add help menu
    # print menu
    print_menu = tk.Menu(menu_bar, tearoff=0)
    # add print menu to menu bar
    menu_bar.add_cascade(label="Print", menu=print_menu)
    # Add print option to print menu
    print_menu.add_command(label="Create Today todo list as html for printing",
                           command=lambda: generate_html_file_to_print())

    help_menu = tk.Menu(menu_bar, tearoff=0)
    # add help menu to menu bar
    menu_bar.add_cascade(label="Help", menu=help_menu)
    # on click on help menu cascade open help window
    help_menu.add_command(label="Help", command=lambda: open_help_window())


add_menus()

# add file menu to root
root.config(menu=menu_bar)
find_tasks_with_due_date_today_and_set_to_a()
load_tasks_in_correct_tab()
root.mainloop()

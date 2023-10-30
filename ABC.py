import win32com.client
import tkinter as tk  # you can use tkinter or another library to create a GUI

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
    # get the task category
    category = task.Categories

    # check if the category is A, B, or C
    if category in ["A", "B", "C"]:
        # append the task and its category to the list
        tasks_by_category.append((task, category))

# sort the list by category in ascending order
tasks_by_category.sort(key=lambda x: x[1])

# create a root window for the GUI
root = tk.Tk()


def close_window(event=None):
    # destroy the root window
    root.destroy()


# bind escape key to close the window
root.bind("<Escape>", close_window)

# make the window full screen from start
#root.attributes('-fullscreen', True)
root.config(bg="white")

# create a frame to display the tasks
frame = tk.Frame(root)
frame.config(width=800, height=600)
frame.pack(fill=tk.BOTH)

# create a canvas to draw the dotted list
canvas = tk.Canvas(frame)
canvas.config(bg="white")
canvas.config(width=700, height=550)
canvas.pack(fill=tk.BOTH)

# set some constants for the drawing
dot_radius = 8  # radius of each dot
dot_gap = 10  # gap between each dot and task
text_gap = 20  # gap between each task and text
line_height = 30  # height of each line


def mark_done(event, task):
    if task.Status == 2:
        task.Status = 1
    else:
        task.Status = 2

    item = canvas.find_withtag("current")
    # Change its fill color to gold
    canvas.itemconfig(item, fill="gold")

    task.Save()


# loop through the sorted list and draw the tasks on the canvas
for i, (task, category) in enumerate(tasks_by_category):
    # get the task subject and due date
    subject = task.Subject
    due_date = task.DueDate

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
    else:
        color = "green"

    # draw a dot on the canvas with the color
    dot = canvas.create_oval(x1 - dot_radius, y1 - dot_radius, x1 + dot_radius, y1 + dot_radius, fill=color,

                             outline=color)
    canvas.tag_bind(dot, "<Button-3>", lambda event, task=task: mark_done(event,task))
    # draw a text with the task information on the canvas with black color and Arial font size 14
    canvas.create_text(x3, y3, text=task_info, fill="black", font=("Arial", 12), anchor=tk.W)

# define a variable to keep track of how many times the dots have blinked in a minute
blink_count = 0

# define a variable to keep track of whether to toggle or not
toggle_flag = False

# define a variable to keep track of whether it is startup or not
startup_flag = True


# define a function that will toggle the color of the red dots once every minute until they blink twice
def toggle_red_dots():
    global blink_count, toggle_flag, startup_flag

    # check if we need to toggle or not
    if toggle_flag:
        # get all the items on the canvas
        items = canvas.find_all()

        # loop through the items and check their color
        for item in items:
            # get the color of the item
            color = canvas.itemcget(item, "fill")

            # if the color is red, change it to white
            if color == "red":
                canvas.itemconfig(item, fill="white", outline="white")

            # if the color is white, change it back to red
            elif color == "white":
                canvas.itemconfig(item, fill="red", outline="red")

        # increment the blink count by 1
        blink_count += 1

        # toggle the flag to False
        toggle_flag = False

        # check if it is startup or not
        if startup_flag:
            # schedule the function to run again after 0.5 seconds
            root.after(10, toggle_red_dots)

        # if it is not startup, schedule the function to run again after one minute
        else:
            # schedule the function to run again after one minute
            root.after(1000, toggle_red_dots)

    # if we don't need to toggle, just wait for another minute
    else:
        # check if the blink count has reached 2 in a minute
        if blink_count < 2:
            # toggle the flag to True
            toggle_flag = True

            # schedule the function to run again after one minute
            root.after(1000, toggle_red_dots)

        # if the blink count has reached 2 in a minute, reset it to 0 and wait for another minute
        else:
            # reset the blink count to 0
            blink_count = 0

            # set the startup flag to False
            startup_flag = False

            # schedule the function to run again after one minute
            root.after(1000, toggle_red_dots)


# call the function for the first time
toggle_red_dots()

# start the main loop of the GUI
root.mainloop()

import win32com.client
import tkinter as tk
from tkinter import font
import time
# Create an Outlook application object
outlook = win32com.client.Dispatch("Outlook.Application")

# Get the namespace object
namespace = outlook.GetNamespace("MAPI")

# Get the Tasks folder object
tasks_folder = namespace.GetDefaultFolder(13)

# Specify the importance level you want to get
importance_level = 2

# Create a list of tasks that have a high priority
high_priority_tasks = []
for task in tasks_folder.Items:
    if task.Importance == importance_level and task.Status != 2:
        high_priority_tasks.append(task)

# Create a tkinter window
window = tk.Tk()
window.title("Important Tasks")

# Create a label to show the number of tasks
label = tk.Label(window, text=f"You have {len(high_priority_tasks)} important tasks:")
label.pack()

# Create a listbox to show the tasks
listbox = tk.Listbox(window)
list_font = font.Font(family='Helvetica',size=10, weight='normal')
listbox.config(width=200, height=50, font=list_font)
listbox.pack()

# Loop through the tasks and insert them into the listbox
for task in high_priority_tasks:
    # Format the task information
    task_info = f"{task.Subject}"
    # Insert the task into the listbox
    listbox.insert(tk.END, task_info)

# Define a function to move an item up in the listbox
def move_up(event=None):
    # Get the index of the selected item
    index = listbox.curselection()
    # If there is an item selected and it is not the first one
    if index and index[0] > 0:
        # Get the value of the selected item
        value = listbox.get(index)
        # Delete the selected item from the listbox
        listbox.delete(index)
        # Insert the item one position above in the listbox
        listbox.insert(index[0]-1, value)
        # Select the item again
        listbox.selection_set(index[0]-1)
        # Swap the corresponding task objects in the high priority tasks list 
        high_priority_tasks[index[0]], high_priority_tasks[index[0]-1] = high_priority_tasks[index[0]-1], high_priority_tasks[index[0]]

# Define a function to move an item down in the listbox
def move_down(event=None):
    # Get the index of the selected item
    index = listbox.curselection()
    # If there is an item selected and it is not the last one
    if index and index[0] < listbox.size()-1:
        # Get the value of the selected item
        value = listbox.get(index)
        # Delete the selected item from the listbox
        listbox.delete(index)
        # Insert the item one position below in the listbox
        listbox.insert(index[0]+1, value)
        # Select the item again
        listbox.selection_set(index[0]+1)
        # Swap the corresponding task objects in the high priority tasks list 
        high_priority_tasks[index[0]], high_priority_tasks[index[0]+1] = high_priority_tasks[index[0]+1], high_priority_tasks[index[0]]

# Define a function to open a new window with a timer and a task subject when double clicking a task in the listbox 
def open_window(event=None):
    # Get the index of the double clicked item 
    index = listbox.curselection()
    # If there is an item double clicked 
    if index:
        # Get the corresponding task object from the high priority tasks list 
        task = high_priority_tasks[index[0]]
        # Create a new window 
        new_window = tk.Toplevel(window)
        new_window.title(task.Subject)
        
        # Make the new window full screen 
        new_window.attributes('-fullscreen', True)

        # Change the background color of the new window to white 
        new_window.config(bg="white")

        # Create a label to show the task subject in big letters 
        subject_label = tk.Label(new_window, text=task.Subject, font=("Helvetica", 48), bg="white")
        subject_label.pack()

        # Create a label to show the body of the task if it has one 
        if task.Body:
            body_label = tk.Label(new_window, text=task.Body, font=list_font, wraplength=200, bg="white")
            body_label.pack()

        # Create a label to show a timer that counts down from 20 minutes to zero 
        timer_label = tk.Label(new_window, text="20:00", font=("Helvetica", 24), bg="white")
        timer_label.pack()

        # Define a function to update the timer label every second 
        def update_timer():
            # Get the current text of the timer label 
            current_time = timer_label.cget("text")
            # Split the text into minutes and seconds 
            minutes, seconds = map(int, current_time.split(":"))
            # If the timer is not zero 
            if minutes or seconds:
                # Decrease the seconds by one 
                seconds -= 1
                # If the seconds are negative 
                if seconds < 0:
                    # Decrease the minutes by one 
                    minutes -= 1
                    # Reset the seconds to 59 
                    seconds = 59
                # Format the new time as mm:ss 
                new_time = f"{minutes:02}:{seconds:02}"
                # Update the timer label with the new time 
                timer_label.config(text=new_time)
                # Schedule the function to run again after one second 
                new_window.after(1000, update_timer)
            else:
                # The timer is zero, so show a message 
                timer_label.config(text="Time's up!")

        # Start the timer function 
        update_timer()
        new_window.focus_force()

        # Define a function to close the new window when pressing escape 
        def close_window(event=None):
            # Destroy the new window 
            new_window.destroy()

        # Bind escape key to close the new window 
        new_window.bind("<Escape>", close_window)

def mark_done(event=None):
    # Get the index of the selected item
    index = listbox.curselection()
    # If there is an item selected
    if index:
        # Get the corresponding task object from the high priority tasks list
        task = high_priority_tasks[index[0]]
        # Try to mark the task as done
        try:
            task.Status = 2
            # Save the changes to the task
            task.Save()
            # Delete the item from the listbox
            listbox.delete(index)
        except Exception as e:
            # Print the error message
            print(e)

def open_task(event=None):
    index = listbox.curselection()
    # If there is an item selected
    if index:
        # Get the corresponding task object from the high priority tasks list
        task = high_priority_tasks[index[0]]
        task.Display()
    

# Create a button to move an item up in the listbox
button_up = tk.Button(window, text="↑", command=move_up)
button_up.pack(side=tk.LEFT)

# Create a button to move an item down in the listbox
button_down = tk.Button(window, text="↓", command=move_down)
button_down.pack(side=tk.RIGHT)

# Bind CTRL+Shift+Up to move an item up in the listbox
listbox.bind("<Control-Shift-Up>", move_up)

# Bind CTRL+Shift+Down to move an item down in the listbox
listbox.bind("<Control-Shift-Down>", move_down)

# Bind Double Left Mouse Button Click to open a new window with a timer and a task subject 
listbox.bind("<Double-Button-1>", open_window)
popup_menu = tk.Menu(listbox, tearoff=0)

# Add a command to mark an item as done
popup_menu.add_command(label="Mark as done", command=mark_done)
popup_menu.add_command(label="Open", command=open_task)


# Define a function to display the popup menu
def show_popup(event):
    # Get the index of the item under the cursor
    index = listbox.nearest(event.y)
    # Select the item
    listbox.selection_clear(0, 'end')
    listbox.selection_set(index)
    # Display the popup menu
    popup_menu.post(event.x_root, event.y_root)

# Bind CTRL+D to mark an item as done
listbox.bind("<Control-d>", mark_done)

# Bind Button-3 (right-click) to show the popup menu
listbox.bind("<Button-3>", show_popup) # Button-2 on Mac

# Start the main loop of the window
window.mainloop()

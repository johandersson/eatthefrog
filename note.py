# # Import the modules
# import tkinter as tk
# import win32com.client as win32
#
# # Create a tkinter window
# window = tk.Tk()
# window.title("Save Note in Outlook")
# window.geometry("500x500")
#
# # Create a label and a text widget for the note content
# label = tk.Label(window, text="Enter your note:")
# label.pack()
# text = tk.Text(window)
# text.pack()
#
# # Create a label and an option menu for the note category
# label2 = tk.Label(window, text="Choose a category:")
# label2.pack()
# # Define a list of color categories
# categories = ["Blue", "Green", "Orange", "Purple", "Red"]
# # Create a variable to store the selected category
# selected_category = tk.StringVar()
# selected_category.set(categories[0])  # Set the default value to the first category
# # Create an option menu with the categories list and the selected category variable
# option_menu = tk.OptionMenu(window, selected_category, *categories)
# option_menu.pack()
#
#
# # Define a function to save the note in Outlook
# def save_note():
#     # Get the note content from the text widget
#     note_content = text.get("1.0",
#                             "end-1c")  # Get all the text from the first character to the last character excluding the newline
#     # Get the selected category from the option menu
#     note_category = selected_category.get()
#     # Create an Outlook application object
#     outlook = win32.Dispatch("Outlook.Application")
#     # Create a new note item
#     note = outlook.CreateItem(5)  # 5 is the constant for olNoteItem
#     # Set the body of the note to the note content
#     note.Body = note_content
#     # Set the color of the note to match the selected category
#     # The color constants are from https://docs.microsoft.com/en-us/office/vba/api/outlook.olnotecolor
#     if note_category == "Blue":
#         note.Color = 0  # olBlue
#     elif note_category == "Green":
#         note.Color = 1  # olGreen
#     elif note_category == "Orange":
#         note.Color = 2  # olOrange
#     elif note_category == "Purple":
#         note.Color = 3  # olPurple
#     elif note_category == "Red":
#         note.Color = 4  # olRed
#     # Save the note in Outlook
#     note.Save()
#     # Display a message in the statusbar to confirm the saving
#     statusbar.config(text="Your note has been saved in Outlook.")
#     # Clear the text widget after saving the note
#     text.delete("1.0", "end")
#
#
# # Define a function to bind the CTRL+S command to the save_note function
# def bind_ctrl_s(event):
#     save_note()
#
#
# # Bind the CTRL+S command to the window and call the bind_ctrl_s function
# window.bind("<Control-s>", bind_ctrl_s)
#
# # Create a button to trigger the save_note function
# button = tk.Button(window, text="Save Note", command=save_note)
# button.pack()
#
# # Create a statusbar to show messages
# statusbar = tk.Label(window, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
# statusbar.pack(side=tk.BOTTOM, fill=tk.X)
#
# # Start the main loop of the window
# window.mainloop()
# Refactor to Note class
# Path: note.py
# Import the modules
import tkinter as tk
import win32com.client as win32


# Create a Note class
class Note:
    def __init__(self, window=None):
        if window is None:
            self.window = tk.Tk()

        self.window = window
        self.window.title("Save Note in Outlook")
        self.window.geometry("500x500")

        # Create a label and a text widget for the note content
        self.label = tk.Label(window, text="Enter your note:")
        self.label.pack()
        self.text = tk.Text(window)
        self.text.pack()

        # Create a label and an option menu for the note category
        self.label2 = tk.Label(window, text="Choose a category:")
        self.label2.pack()
        # Define a list of color categories
        self.categories = ["Blue", "Green", "Orange", "Purple", "Red"]
        # Create a variable to store the selected category
        self.selected_category = tk.StringVar()
        self.selected_category.set(self.categories[0])  # Set the default value to the first category
        # Create an option menu with the categories list and the selected category variable
        self.option_menu = tk.OptionMenu(window, self.selected_category, *self.categories)
        self.option_menu.pack()

        # Create a button to trigger the save_note function
        self.button = tk.Button(window, text="Save Note", command=self.save_note)
        self.button.pack()

        # Create a statusbar to show messages
        self.statusbar = tk.Label(window, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Bind the CTRL+S command to the window and call the bind_ctrl_s function
        window.bind("<Control-s>", self.bind_ctrl_s)

    # Define a function to save the note in Outlook
    def save_note(self):
        # Get the note content from the text widget
        note_content = self.text.get("1.0")
        # Get the selected category from the option menu
        note_category = self.selected_category.get()
        # Create an Outlook application object
        outlook = win32.Dispatch("Outlook.Application")
        # Create a new note item
        note = outlook.CreateItem(5)
        # Set the body of the note to the note content
        note.Body = note_content
        # Set the color of the note to match the selected category
        # The color constants are from https://docs.microsoft.com/en-us/office/vba/api/outlook.olnotecolor
        if note_category == "Blue":
            note.Color = 0
        elif note_category == "Green":
            note.Color = 1
        elif note_category == "Orange":
            note.Color = 2
        elif note_category == "Purple":
            note.Color = 3
        elif note_category == "Red":
            note.Color = 4
        # Save the note in Outlook
        note.Save()
        # Display a message in the statusbar to confirm the saving
        self.statusbar.config(text="Your note has been saved in Outlook.")
        # Clear the text widget after saving the note
        self.text.delete("1.0", "end")

    # Define a function to bind the CTRL+S command to the save_note function
    def bind_ctrl_s(self, event):
        self.save_note()

    def run(self):
        # Start the main loop of the window
        self.window.mainloop()
        # on load focus on note content
        self.focus_note_content()

    def focus_note_content(self):
        self.text.focus_set()

    #on load focus on note content
    def __enter__(self):
        self.focus_note_content()
        return self

#if main
if __name__ == "__main__":
    #if this is run standalone, create a Note object and run it
    window = tk.Tk()
    note = Note(window)
    note.run()

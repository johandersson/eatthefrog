import tkinter as tk
# Import the win32com.client to create an Outlook application object
import win32com.client as win32


# Create a Note class
class Note(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.title("Save Note in Outlook")
        self.geometry("500x500")

        # Create a label and a text widget for the note content
        self.label = tk.Label(self, text="Enter your note:")
        self.label.pack()
        self.text = tk.Text(self)
        self.text.pack()

        # Create a label and an option menu for the note category
        self.label2 = tk.Label(self, text="Choose a category:")
        self.label2.pack()
        # Define a list of color categories
        self.categories = ["Blue", "Green", "Pink", "Yellow", "White"]
        # Create a variable to store the selected category
        self.selected_category = tk.StringVar()
        self.selected_category.set(self.categories[0])  # Set the default value to the first category
        # Create an option menu with the categories list and the selected category variable
        self.option_menu = tk.OptionMenu(self, self.selected_category, *self.categories)
        self.option_menu.pack()

        # Create a button to trigger the save_note function
        self.button = tk.Button(self, text="Save Note", command=self.save_note)
        self.button.pack()

        # Create a statusbar to show messages
        self.statusbar = tk.Label(self, text="", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Bind the CTRL+S command to the window and call the bind_ctrl_s function
        self.bind("<Control-s>", self.bind_ctrl_s)
        #focus on the text widget
        self.text.focus_set()

    # Define a function to save the note in Outlook
    def save_note(self):
        # Get the note content from all the text from the first character to the last character excluding the newline
        note_content = self.text.get("1.0", "end-1c")
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
        elif note_category == "Pink":
            note.Color = 2
        elif note_category == "Yellow":
            note.Color = 3
        elif note_category == "White":
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

    def focus_note_content(self):
        self.text.focus_set()

if __name__ == "__main__":
    app = Note()
    app.mainloop()

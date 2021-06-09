from tkinter import *
from tkinter import filedialog


# Function for opening the
# file explorer window
def browseFiles():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("CSV Files", "*.csv*"),
                                                     ("Excel Files", "*.xlsx*")))
    if filename != "":
        label_file_explorer.configure(text="File Opened: " + filename)


window = Tk()

window.title('Excel Operations using Python')

window.geometry("500x500")

window.config(background="white")

label_file_explorer = Label(window,
                            text="Excel Operations using Python",
                            width=70, height=4,
                            fg="blue")

button_explore = Button(window,
                        text="Browse Files",
                        command=browseFiles)

v = DoubleVar()
scale = Scale(window, variable=v, from_=0.00, to=10.00, orient=HORIZONTAL)


button_ = Button(window,
                        text="Input File",
                        command=browseFiles)

button_process = Button(window,
                     text="Process Data",
                     command=exit)

label_file_explorer.grid(column=1, row=1)

button_explore.grid(column=1, row=2)

scale.grid(column=1, row= 3)

button_process.grid(column=1, row=4)

window.mainloop()

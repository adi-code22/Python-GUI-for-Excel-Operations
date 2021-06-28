from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import pathlib
import os

pd.options.mode.chained_assignment = None





def browsefiles():
    try:
        temp1 = filedialog.askopenfilename(initialdir="/",
                                           title="Select a File",
                                           filetypes=(
                                               ("Excel Files", "*.xlsx*"), ("CSV Files", "*.csv*")))
        filename.set(temp1)
        if pathlib.Path(temp1).suffix == ".xlsx":
            pass
        elif pathlib.Path(temp1).suffix == ".csv":
            pass
        if temp1 != "":
            label_file_explorer.configure(text="File Opened: " + filename.get())
            fileselect.configure(text="File Selected and Ready to Process", fg="green")

    except Exception as e:
        print(e)


def process():
    try:
        temp = filename.get()
        temp2 = pd.read_excel(temp)
        df.set(temp2)
        temp2['Difference'] = temp2['Amount 1'] - temp2['Amount 2']
        temp2['Assessment'] = ""
        slider = v.get()
        for name, value in temp2.iteritems():
            if name == "Difference":
                for i in range(value.size):
                    if value[i] >= slider:
                        for name2, value2 in temp2.iteritems():
                            if name2 == "Assessment":
                                for i2 in range(value.size):
                                    value2[i] = 'inappropriate'

                    else:
                        for name2, value2 in temp2.iteritems():
                            if name2 == "Assessment":
                                for i2 in range(value.size):
                                    value2[i] = 'appropriate'
        processdata.configure(text="Data Successfully Processed, Ready to LAUNCH", fg="green")
        print(temp2.head())
        writer = pd.ExcelWriter(temp)
        temp2.to_excel(writer, index=False)
        writer.save()
    except Exception as e:
        print(e)


def launch():
    try:
        launchl.configure(text="Excel Successfully Launched", fg="green")
        temp = filename.get()
        os.system(f"start EXCEL.EXE {temp}")
    except Exception as e:
        print(e)


start = Tk()
start.title("Excel Operation with Python")
start.geometry("500x400")

filename = StringVar()
df = Variable()
v = DoubleVar()
ex = BooleanVar()
cs = BooleanVar()

# Create Main Frame
main_frame = Frame(start)
main_frame.pack(fill=BOTH, expand=1)
# Create Canvas
my_canvas = Canvas(main_frame)
my_canvas.pack(side=LEFT, fill=BOTH, expand=1)
# Add Scrollbar to Canvas
my_scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
my_scrollbar.pack(side=RIGHT, fill=Y)

# Configure the Canvas
my_canvas.configure(yscrollcommand=my_scrollbar.set)
my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
# Create new frame inside Canvas
window = Frame(my_canvas)
# Add new frame to window
my_canvas.create_window((0, 0), window=window, ancho='nw')

# window.config(background="white")

t = Text(window, height=1, width=20)


def excelf():
    cs.set(False)


excel = Radiobutton(window, text="Save as .xlsv", variable=ex, command=excelf)


def csvf():
    ex.set(True)


csv = Radiobutton(window, text="Save as .csv", variable=cs, command=csvf)

fileselect = Label(window,
                   text="-------------------------------------",
                   width=70, height=1,
                   fg="red")

processdata = Label(window,
                    text="-------------------------------------",
                    width=70, height=1,
                    fg="red")

launchl = Label(window,
                text="-------------------------------------",
                width=70, height=1,
                fg="red")

label_file_explorer = Label(window,
                            text="Add Date",
                            width=70, height=4,
                            fg="blue")

button_explore = Button(window,
                        text="Browse Files",
                        command=browsefiles, width=25, height=2)

scale = Scale(window, variable=v, from_=0.1, to=10.0, orient=HORIZONTAL, resolution=0.1, sliderlength=20,
              activebackground="blue", width=25, length=300, label="Threshold")

button_launch = Button(window,
                       text="Launch Output",
                       command=launch, width=25, height=2)

button_process = Button(window,
                        text="Process Data",
                        command=process, width=25, height=2)

label_file_explorer.pack(side=TOP, padx=(0, 0))

t.pack(side=TOP)

button_explore.pack(side=TOP, pady=(25, 25))

fileselect.pack(side=TOP, padx=(0, 0))

scale.pack(side=TOP)

button_process.pack(side=TOP, pady=(25, 25))

processdata.pack(side=TOP, padx=(0, 0))

excel.pack(side=TOP)
csv.pack(side=TOP)
button_launch.pack(side=TOP, pady=(25, 25))

launchl.pack(side=TOP, padx=(0, 0))

window.mainloop()

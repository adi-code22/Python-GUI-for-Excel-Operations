from tkinter import *
from tkinter import filedialog
import pandas as pd
import pathlib
import os

pd.options.mode.chained_assignment = None


# Function for opening the
# file explorer window


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
    except EXCEPTION as e:
        print(e)


def process():
    try:
        temp = filename.get()
        temp2 = pd.read_excel(temp)
        df.set(temp2)
        temp2['Difference'] = temp2['Column1'] - temp2['Column2']
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
    except EXCEPTION as e:
        print(e)


    # try:
    #     temp = filename.get()
    #     temp2 = pd.read_csv(temp)
    #     print(temp2.head())
    #     df.set(temp2)
    #     temp2['Difference'] = ""
    #     temp2.to_csv(temp, index=False)
    #     for name, values in temp2.iteritems():
    #         if name == "column1":
    #             print(values[2])
    #     temp2['Assessment'] = ""
    #
    #     slider = v.get()
    #     for name, value in temp2.iteritems():
    #         if name == "Difference":
    #             for i in range(value.size):
    #                 print(value[i])
    #                 if int(value[i]) >= int(slider):
    #                     for name2, value2 in temp2.iteritems():
    #                         if name2 == "Assessment":
    #                             for i2 in range(value.size):
    #                                 value2[i] = 'inappropriate'
    #
    #                 else:
    #                     for name2, value2 in temp2.iteritems():
    #                         if name2 == "Assessment":
    #                             for i2 in range(value.size):
    #                                 value2[i] = 'appropriate'
    #     processdata.configure(text="Data Successfully Processed, Ready to LAUNCH", fg="green")
    #     print(temp2.head())
    #     writer = pd.ExcelWriter(temp)
    #     temp2.to_csv(writer, index=False)
    #     writer.save()
    # except Exception as e:
    #     print(e)





def launch():
    try:
        launchl.configure(text="Excel Successfully Launched", fg="green")
        temp = filename.get()
        os.system(f"start EXCEL.EXE {temp}")
    except EXCEPTION as e:
        print(e)


window = Tk()
filename = StringVar()
df = Variable()
v = DoubleVar()

window.title('Excel Operations using Python')

window.geometry("500x500")

window.config(background="white")

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
                            text="Excel Operations using Python",
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

label_file_explorer.grid(column=1, row=1, padx=(0, 0))

button_explore.grid(column=1, row=2, padx=(0, 0), pady=(25, 25))

fileselect.grid(column=1, row=3, padx=(0, 0))

scale.grid(column=1, row=4)

button_process.grid(column=1, row=5, pady=(25, 25))

processdata.grid(column=1, row=6, padx=(0, 0))

button_launch.grid(column=1, row=7, pady=(25, 25))

launchl.grid(column=1, row=8, padx=(0, 0))

window.mainloop()

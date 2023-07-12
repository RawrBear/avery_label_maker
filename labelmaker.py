import tkinter as tk
import ttkbootstrap as ttk
from utils import runProcess
import global_
from threading import *
from tkinter import filedialog as fd


#### GUI ####

# Set up processing function on a second thread


def process_threading():
    # Call work function
    t1 = Thread(target=runProcess)
    t1.start()


def message_to_console():
    # console_output_string.set(global_.message)
    console_output_label.config(text=global_.message)
    console_output_label.after(1000, message_to_console)


def open_data_file():
    # Open filechooser
    # listPath = filechooser.open_file(
    #     title="Choose your list..", filters=[("Excel", "*.xlsx")]
    listPath = fd.askopenfilenames(
        title="Choose your list..",
        initialdir="/",
        filetypes=[("Excel", "*.xlsx")],
    )

    # Update the global variable
    global_.list_location = listPath[0]

    # Update the textbox variable
    list_file_path.set(listPath[0])

    # Update the console output
    global_.message = "Excel File Selected"
    message_to_console()


def open_save_location():
    # Open filechooser
    # outputPath = filechooser.choose_dir(
    #     title="Where to save the output?", filters=[("All Files", "*.*")]
    # )
    outputPath = fd.askdirectory(
        title="Where to save the output?",
        initialdir="/"
        # filetypes=[("All Files", "*.*")]
        # initialdir="/"
    )
    print(outputPath)

    # Update the global variable
    global_.save_location = outputPath

    # Update the textbox variable
    save_location_path.set(outputPath)

    # Update the console output
    global_.message = "Output Folder Selected"


# Window
window = ttk.Window(themename="darkly")
window.title("Awesome LabelMaker")
window.geometry("600x400")
window.iconbitmap("./labelmaker.ico")

# Title
title_label = ttk.Label(window, text="Awesome LabelMaker", font=("Calibri", 24))
title_label.pack()

# Open Date File
open_data_frame = ttk.Frame(window)
list_file_path = ttk.StringVar()
list_file_entry = ttk.Entry(open_data_frame, width=30, textvariable=list_file_path)
data_file_button = ttk.Button(
    open_data_frame, text="Open List...", command=open_data_file
)
list_file_entry.pack(side="right", padx=10)
data_file_button.pack(side="left")
open_data_frame.pack(pady=10)


# Open Output Folder
open_save_frame = ttk.Frame(window)
save_location_path = ttk.StringVar()
save_location_entry = ttk.Entry(
    open_save_frame, width=30, textvariable=save_location_path
)
save_location_button = ttk.Button(
    open_save_frame, text="Save Location...", command=open_save_location
)
open_save_frame.pack(pady=10)
save_location_entry.pack(side="right", padx=10)
save_location_button.pack(side="left")


# Process Button
process_button = ttk.Button(window, text="Process", command=process_threading)

# Pack
process_button.pack(pady=10)

# Console Output
console_output_label = ttk.Label(window, text="Output", font=("Calibri", 20))
console_output_label.pack(
    pady=20,
)

# RUN
window.mainloop()

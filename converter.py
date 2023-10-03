#!/usr/bin/env python3

import pyexcel as pe
from ethiopian_date import EthiopianDateConverter
import datetime
import pyperclip
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

Tk().withdraw()
file_path = filedialog.askopenfilename()
if not file_path:
    exit()
sheet = ""
frame = Tk()

def get_sheet():
    sheet_name = inputtxt.get()
    conv = EthiopianDateConverter.to_gregorian
    book = pe.get_book(file_name=file_path)
    try:
        sheet = book[sheet_name]
    except Exception:
        messagebox.showerror("Error","Can't find Sheet")
        exit()
    try:
        ethopian_dates = sheet.column[0]
        result = []
        ethopian_dates.pop()
        length = len(ethopian_dates)
        for i in range(length):
            if i == 0:
                continue
            if (type(ethopian_dates[i]) != datetime.date):
                result.append("")
                continue
            result.append(
                conv(ethopian_dates[i].year, ethopian_dates[i].month, ethopian_dates[i].day).strftime("%d/%m/%Y"))
        sheet.column[1] = result
    except Exception as err:
        messagebox.showerror("Unkown Error", "Error: {}".format(err))
    pyperclip.copy('\n'.join(result))
    messagebox.showinfo("Success","Converted dates copied to clipboard")
    frame.destroy()
    exit()

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        frame.destroy()
        exit()

frame.title("Name of Sheet")
frame.geometry('200x150')
frame.eval('tk::PlaceWindow . center')
frame.protocol("WM_DELETE_WINDOW", on_closing)

infolbl = Label(frame,text="INFO: Dates must have the form \n dd/mm/yyyy")
infolbl.place(x=0,y=70)

inputtxt = Entry(frame)
inputtxt.pack()

selectbttn = Button(frame,text = "Select", command = get_sheet)
selectbttn.pack()

frame.mainloop()

# LABORATORY #4 BSEE 3RD YEAR BLOCK IDENTIFICATION CODE
# This code is Submitted By: Matthew David M. Loquinerio for the subject Computer Programming

# Directions: Create a python code that can identify the BLOCK ( A, B, C, D) of a student.
# The Input of the code is either the First Name or the Student Number
# The code must ask  the Last name if ever the first name is identical with other students.
# The output is the Block of the student.

#Import all dependencies
import numbers
from threading import local
from tkinter import *
from tkinter import ttk
import tkinter
from tkinter import messagebox
import pandas as pd

#Admin GUI
root = tkinter.Tk()
root.geometry("1300x800")
root.pack_propagate(False)
root.resizable(0, 0)
root.title('BSEE BLOCK LISTING')

#Main Frame to Display Block List
dataFrame = LabelFrame(root, text="Excel Data")
dataFrame.place(height=656, width=1300)

#Search Frame to input data
Frm = LabelFrame(root, text="Search Here")
Frm.place(height=128, width=720, rely=0.82, relx=0.25)

labelName = Label(Frm, text='FILE NAME:')
labelName.place(rely = 0.02, relx = 0.007)

labelFile = Label(Frm, text="No File Selected")
labelFile.place(rely=0.02, relx=0.11)

labelBlock = Label(Frm, text='STUDENT BLOCK:')
labelBlock.place(rely = 0.02, relx = 0.425)

#Resulting Block from file origin
labelSheet = Label(Frm, text="No Block Indicated")
labelSheet.place(rely=0.02, relx=0.575)

#Student Number input
Label(Frm, text='STUDENT NUMBER:').pack(side=LEFT, padx=2)
entryName = Entry(Frm, width=30)
entryName.pack(side=LEFT, padx=2)

entryName.focus_set()

#Student Number input
Label(Frm, text='STUDENT NAME:').pack(side=LEFT, padx=2)
entryNumber = Entry(Frm, width=50)
entryNumber.pack(side=LEFT, padx=2)

#Initialize to run search operation
button2 = Button(Frm, text="Submit Response", command=lambda: Load_excel_data())
button2.place(rely=0.68, relx=0.42)

#Display Block Data in Tables
tv1 = ttk.Treeview(dataFrame)
tv1.place(relheight=1, relwidth=1)

treescrolly = Scrollbar(dataFrame, orient="vertical", command=tv1.yview)
treescrollx = Scrollbar(dataFrame, orient="horizontal", command=tv1.xview) 
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")

#Origin File
file = "blocklists.xlsx"
labelFile["text"] = file

#Main Function
def Load_excel_data():
    inputName = entryName.get()
    inputNumber = entryNumber.get()

    name = inputName.upper()
    number = inputNumber.upper()

    sheet1 = 'BSEE3A'
    sheet2 = 'BSEE3B'
    sheet3 = 'BSEE3C'
    sheet4 = 'BSEE3D'

    all = pd.read_excel(file)
    df = pd.read_excel(file, sheet_name=sheet1)
    df2 = pd.read_excel(file, sheet_name=sheet2)
    df3 = pd.read_excel(file, sheet_name=sheet3)
    df4 = pd.read_excel(file, sheet_name=sheet4)

    name1 = "JOSHUA"
    name2 = "MARK ANTHONY"

    if name == name1:
        messagebox.showwarning("Warning","Warning message for user")
        return None
        
    elif name == name2:
        messagebox.showwarning("Warning","Warning message for user")
        return None
    
    def studentName():
        if name in df.values:
            try:
                print(sheet1)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet1

        elif name in df2.values:
            try:
                print(sheet2)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df2.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df2.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet2

        elif name in df3.values:
            try:
                print(sheet3)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df3.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df3.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet3

        elif name in df4.values:
            try:
                print(sheet4)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df4.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df4.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet4

    def studentNumber():
        if number in df.values:
            try:
                print(sheet1)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet1

        elif number in df2.values:
            try:
                print(sheet2)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df2.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df2.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet2

        elif number in df3.values:
            try:
                print(sheet3)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df3.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df3.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet3

        elif number in df4.values:
            try:
                print(sheet4)
            except ValueError:
                print('Value does not exist')

            clear_data()
            tv1["column"] = list(df4.columns)
            tv1["show"] = "headings"
            for column in tv1["columns"]:
                tv1.heading(column, text=column) 

            df_rows = df4.to_numpy().tolist()
            for row in df_rows:
                tv1.insert("", "end", values=row)

            labelSheet["text"] = sheet4

    studentName()
    studentNumber()


def clear_data():
    tv1.delete(*tv1.get_children())
    return None

root.mainloop()
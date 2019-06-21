import tkinter as tk
from tkinter import *
from tkinter import ttk
from docx import Document
import datetime
import docx
import os
import sqlite3
app = tk.Tk()

def Import():

    db_name = 'Recruit_Database.db'

    if "\\" in File_path.get():
        New_Path = str(File_path.get()).replace("\\","/")
        file_Path = os.path.abspath(New_Path)
        doc = docx.Document(file_Path)
        print(doc.paragraphs[0].text)
        Role = doc.paragraphs[0].text
        print(doc.paragraphs[2].runs[5].text)
        Type = doc.paragraphs[2].runs[5].text
        Para3 = (len(doc.paragraphs[3].runs)-1)
        Location = doc.paragraphs[3].runs[2].text + doc.paragraphs[3].runs[Para3].text
        Email = Assign.get()
        Status = "New"
        Record_list = [str(Date),
                       Role,
                       Type,
                       Location,
                       Email,
                       str(Due_Date),
                       Status]
        print(Record_list)
        con = sqlite3.connect('Recruit_Database.db')
        cur = con.cursor()
        cur.execute(
            ' INSERT INTO Progression_Record(Date, Role, Type, Location, Assign_To, Due_Date, Status) VALUES(?,?,?,?,?,?,?)',
            Record_list)
        con.commit()
        show_record()
    else:
        file_Path = os.path.abspath(File_path.get())
        doc = docx.Document(file_Path)
        print(doc.paragraphs[0].text)
        Role = doc.paragraphs[0].text
        print(doc.paragraphs[2].runs[5].text)
        Type = doc.paragraphs[2].runs[5].text
        Para3 = (len(doc.paragraphs[3].runs) - 1)
        Location = doc.paragraphs[3].runs[2].text + doc.paragraphs[3].runs[Para3].text
        Email = Assign.get()
        Status = "New"
        Record_list = [str(Date),
                       Role,
                       Type,
                       Location,
                       Email,
                       str(Due_Date),
                       Status]
        print(Record_list)
        con = sqlite3.connect('Recruit_Database.db')
        cur = con.cursor()
        cur.execute(
            ' INSERT INTO Progression_Record(Date, Role, Type, Location, Assign_To, Due_Date, Status) VALUES(?,?,?,?,?,?,?)',
            Record_list)
        con.commit()
        show_record()

def show_record():
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect('Recruit_Database.db')
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record ORDER BY ID DESC')
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


Date = datetime.datetime.now().date()
Due_Date = datetime.date.today() + datetime.timedelta(days=1)
Label(app, text = "กรุณาใส่ที่อยู่ไฟล์").grid(row = 0)
File_path = tk.Entry(app, width = 40)
File_path.grid(row = 0, column = 1)

Label(app, text = "Assign To").grid(row = 0 , column = 2)
Assign = tk.Entry(app, width = 40)
Assign.grid(row = 0, column = 3)

Import_but = tk.Button(app, text = "Import", command = Import)
Import_but.grid(row = 0,column = 4)

Recoard_frame = LabelFrame(app, text ="")
Recoard_frame.place(x = 10, y = 50)

tree = ttk.Treeview(Recoard_frame, height=15, column=("1","2","3","4","5","6"))
tree.grid(row=1, column=0)
tree.heading('#0', text='Date', anchor=W)
tree.heading(1, text='Role', anchor=W)
tree.heading(2, text='Type', anchor=W)
tree.heading(3, text='Location', anchor=W)
tree.heading(4, text='Assign To', anchor=W)
tree.heading(5, text='Due Date', anchor=W)
tree.heading(6, text='Status', anchor=W)
tree.column('#0', width = 100)
tree.column(1, width = 200)
tree.column(2, width = 100)
tree.column(3, width = 100)
tree.column(4, width = 200)
tree.column(5, width = 100)
tree.column(6, width = 100)
show_record()





app.geometry("1000x500")
app.mainloop()



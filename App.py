import tkinter as tk
from tkinter import *
from tkinter import ttk
from docx import Document
import datetime
import docx
import os
import sqlite3
import xlsxwriter
from tkinter import filedialog

app = tk.Tk()
db_name = 'Recruit_Database.db'


def Import():
    db_name = 'Recruit_Database.db'

    file_Path = os.path.abspath(File_path.get())
    doc = docx.Document(file_Path)
    Role = doc.paragraphs[0].text
    Para2 = doc.paragraphs[2].text
    new_para2 = re.sub('\s+', '', Para2)
    Type = new_para2.split(":", 1)[1]
    Para3 = doc.paragraphs[3].text
    new_para3 = re.sub('\s+', '', Para3)
    Location = new_para3.split(":", 1)[1]
    if "BTS" in Location:
        BTS = Location[0:3]
        Station = Location[3:]
        Location = BTS + " " + Station

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
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute(
        ' INSERT INTO Progression_Record(Date, Role, Type, Location, Assign_To, Due_Date, Status) VALUES(?,?,?,?,?,?,?)',
        Record_list)
    con.commit()
    show_record()
    File_path.delete(0, 'end')
    Assign.delete(0, 'end')


def get_path():
    word_file = filedialog.askopenfilename()
    File_path.insert(0, word_file)


def export_excel():
    excel_path = filedialog.askdirectory()
    list = []
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT Date,Role,Type,Location,Assign_To,Due_Date,Status FROM Progression_Record ORDER BY ID DESC')
    for row in cur.fetchall():
        list.insert(0, row)

    workbook = xlsxwriter.Workbook(excel_path + '/Test.xlsx')
    worksheet = workbook.add_worksheet("My sheet")
    top_row = 0
    top_col = 0
    row = 1
    col = 0
    Hcell_format = workbook.add_format()
    Hcell_format.set_pattern()
    Hcell_format.set_bg_color('yellow')
    Hcell_format.set_bold()

    Stcell_format = workbook.add_format()
    Stcell_format.set_pattern()
    Stcell_format.set_bg_color('red')

    for width in range(0, 7):
        worksheet.set_column(width, width, 25)

    head_col = ["Date", "Role", "Type", "Location", "Assign To", "Due Date", "Status"]

    for head in head_col:
        worksheet.write(top_row, top_col, head, Hcell_format)
        top_col += 1

    for date, role, type, location, email, due_date, status in (list):
        worksheet.write(row, col, date)
        worksheet.write(row, col + 1, role)
        worksheet.write(row, col + 2, type)
        worksheet.write(row, col + 3, location)
        worksheet.write(row, col + 4, email)
        worksheet.write(row, col + 5, due_date)
        if status == 'Delay':
            worksheet.write(row, col + 6, status, Stcell_format)
        else:
            worksheet.write(row, col + 6, status)
        row += 1

    workbook.close()


def show_record():
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record ORDER BY ID DESC')
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


def Date_Search(event):
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record WHERE Date like ?', ('%' + Date_search.get() + '%',))
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


def Role_Search(event):
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record WHERE Role like ?', ('%' + Role_search.get() + '%',))
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


def Type_Search(event):
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record WHERE Type like ?', ('%' + Type_search.get() + '%',))
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


def Location_Search(event):
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record WHERE Location like ?', ('%' + Location_search.get() + '%',))
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


def Email_Search(event):
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record WHERE Assign_To like ?', ('%' + Email_search.get() + '%',))
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


def Due_Search(event):
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record WHERE Due_Date like ?', ('%' + Due_search.get() + '%',))
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


def Status_Search(event):
    records = tree.get_children()
    for element in records:
        tree.delete(element)
    con = sqlite3.connect(db_name)
    cur = con.cursor()
    cur.execute('SELECT * FROM Progression_Record WHERE Status like ?', ('%' + Status_search.get() + '%',))
    for row in cur.fetchall():
        tree.insert('', 0, text=row[1], values=(row[2], row[3], row[4], row[5], row[6], row[7]))


Date = datetime.datetime.now().date()
Due_Date = datetime.date.today() + datetime.timedelta(days=1)
Label(app, text="File Path: ").grid(row=0, column=0)
File_path = tk.Entry(app, width=40)
File_path.grid(row=0, column=1)

Label(app, text="Assign To: ").grid(row=0, column=3)
Assign = tk.Entry(app, width=40)
Assign.grid(row=0, column=4)

Import_but = tk.Button(app, text="Import", command=get_path)
Import_but.grid(row=0, column=2)

Export_but = tk.Button(app, text="Export as Excel", command=export_excel)
Export_but.place(x=820, y=20)

Save_but = tk.Button(app, text="Save", command=Import)
Save_but.grid(row=0, column=5)
Recoard_frame = LabelFrame(app, text="")
Recoard_frame.place(x=10, y=50)

tree = ttk.Treeview(Recoard_frame, height=15, column=("1", "2", "3", "4", "5", "6"))
tree.grid(row=1, column=0)
tree.heading('#0', text='Date', anchor=W)
tree.heading(1, text='Role', anchor=W)
tree.heading(2, text='Type', anchor=W)
tree.heading(3, text='Location', anchor=W)
tree.heading(4, text='Assign To', anchor=W)
tree.heading(5, text='Due Date', anchor=W)
tree.heading(6, text='Status', anchor=W)
tree.column('#0', width=100)
tree.column(1, width=200)
tree.column(2, width=100)
tree.column(3, width=100)
tree.column(4, width=200)
tree.column(5, width=100)
tree.column(6, width=100)

Search_frame = LabelFrame(app, text="Searching")
Search_frame.place(x=10, y=390)

Label(Search_frame, text="Date:").grid(row=0, column=0)
Date_search = tk.Entry(Search_frame)
Date_search.bind('<KeyRelease>', Date_Search)
Date_search.grid(row=0, column=1)

Label(Search_frame, text="Role:").grid(row=0, column=2)
Role_search = tk.Entry(Search_frame)
Role_search.bind('<KeyRelease>', Role_Search)
Role_search.grid(row=0, column=3)

Label(Search_frame, text="Type:").grid(row=0, column=4)
Type_search = tk.Entry(Search_frame)
Type_search.bind('<KeyRelease>', Type_Search)
Type_search.grid(row=0, column=5)

Label(Search_frame, text="Location:").grid(row=0, column=6)
Location_search = tk.Entry(Search_frame)
Location_search.bind('<KeyRelease>', Location_Search)
Location_search.grid(row=0, column=7)

Label(Search_frame, text="Email:").grid(row=1, column=0)
Email_search = tk.Entry(Search_frame)
Email_search.bind('<KeyRelease>', Email_Search)
Email_search.grid(row=1, column=1)

Label(Search_frame, text="Due_Date:").grid(row=1, column=2)
Due_search = tk.Entry(Search_frame)
Due_search.bind('<KeyRelease>', Due_Search)
Due_search.grid(row=1, column=3)

Label(Search_frame, text="Status:").grid(row=1, column=4)
Status_search = tk.Entry(Search_frame)
Status_search.bind('<KeyRelease>', Status_Search)
Status_search.grid(row=1, column=5)

show_record()

app.geometry("1000x500")
app.mainloop()



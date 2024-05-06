from tkinter import filedialog, ttk
from tkinter import *
import pandas as pd
import sqlite3
from datetime import datetime
import os

path = ''
table_name = ''
file_name = ''
file_dir = ''


def openfile():
    file_entry.delete(0, END)
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select A File",
                                               filetypes=(("SQLite files", "*.db"), ("all files", "*.*")))
    global path
    path = root.filename
    file_entry.insert(0, path)

    read_table_names(path)
    state_label.config(text="")


def read_table_names(_db_path):
    connection = sqlite3.connect(f"{_db_path}")
    cursor = connection.cursor()
    cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    connection.close()
    tables_lst = []
    for i in range(len(tables)):
        if tables[i][0] != 'sqlite_sequence':
            tables_lst.append(tables[i][0])
    table_name_drop.config(values=tables_lst)
    table_name_drop.current(0)
    table_name_drop['state'] = NORMAL
    table_name_drop.config(state='readonly')
    run_button['state'] = NORMAL


def run_convert():
    state_label.config(text="converting...")
    current_datetime = datetime.now().strftime("%Y%m%d-%H%M%S")
    connection = sqlite3.connect(f"{path}")
    df = pd.read_sql_query(f"SELECT * FROM {table_name_drop.get()}", connection)
    connection.close()
    global file_name
    global file_dir
    file_name = f"{table_name_drop.get()}-{current_datetime}.xlsx"
    file_dir = path[0:path.rfind("/")]
    df.to_excel(f"{file_dir}/{file_name}", index=False)
    state_label.config(text=f"File saved in {file_dir}/{file_name}")
    open_xlsx_button['state'] = NORMAL


def open_xlsx():
    command = f'start excel.exe "{file_dir}/{file_name}"'
    os.system(command)


# --------------------------------------------------------------------------------------------- GUI
# create GUI
root = Tk()
root.title('SQLite to excel')
root.geometry("400x240")
root.resizable(False, False)
root.grid_columnconfigure(2, weight=1)

# ------------------------------------------------------------------------------------------ widgets
open_label = Label(text="Open database", justify="left")
open_label.grid(row=0, column=0, sticky="w", padx=10, pady=(6, 0))

file_entry = Entry()
file_entry.grid(row=1, column=0, columnspan=2, pady=(6, 0), padx=10, ipadx=50, sticky="EW")

open_button = Button(text="Open SQLite file", command=openfile, width=2)
open_button.grid(row=1, column=2, pady=(6, 0), padx=10, ipadx=50)

table_name_label = Label(text="Select table")
table_name_label.grid(row=2, column=0, sticky="w", padx=10, pady=(6, 0))

table_name_drop = ttk.Combobox(width=14)
table_name_drop.grid(row=3, column=0, columnspan=2, pady=(6, 0), padx=10, ipadx=50, sticky="NEWS")
table_name_drop['state'] = DISABLED

run_button = Button(text="Run", command=run_convert, width=2)
run_button.grid(row=4, column=0, columnspan=3, pady=(26, 0), padx=10, ipadx=80)
run_button['state'] = DISABLED

state_label = Label(text="")
state_label.grid(row=5, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 0))

open_xlsx_button = Button(text="Open .xlsx", command=open_xlsx, width=2)
open_xlsx_button.grid(row=6, column=2, pady=(6, 0), padx=10, ipadx=30, sticky="EW")
open_xlsx_button['state'] = DISABLED

# create the loop
root.mainloop()

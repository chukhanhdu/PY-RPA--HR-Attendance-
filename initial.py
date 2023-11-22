import tkinter as tk
from tkinter import ttk,messagebox
import sqlite3

class Warning:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
    def show_warning(self,warning_content):
        self.root.withdraw()
        messagebox.showerror("HR SYS", warning_content)
        self.root.destroy()
        return
    def show_message(self,message):
        self.root.withdraw()
        messagebox.showinfo("HR SYS", message)
        self.root.destroy()
        return
    
email_pattern = r'^\w+@\w+\.\w+(?:;\w+@\w+\.\w+)+$'

window_opened = False
root = None

def on_child_window_close():
    global window_opened
    window_opened = False
    root.destroy()

def initialize():
    global window_opened, root
    if not window_opened:
        # root window
        root = tk.Tk()
        root.geometry("360x300")

        root.title('Initial')
        root.resizable(0, 0)
        
        # configure the grid
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=3)

        # Declare username_entry và password_entry là biến toàn cục để truy cập từ các hàm khác
        global username_entry, password_entry, manager_mail_entry, BOD_mail_entry,hr_mail_entry

        # e-mail
        username_label = ttk.Label(root, text="Email:")
        username_label.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

        username_entry = ttk.Entry(root)
        username_entry.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

        # password
        password_label = ttk.Label(root, text="Password:")
        password_label.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)

        password_entry = ttk.Entry(root, show="*")
        password_entry.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)

        # To Manager mail
        manager_label = ttk.Label(root, text="Manager")
        manager_label.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)

        manager_mail_entry = ttk.Entry(root)
        manager_mail_entry.grid(column=1, row=2, sticky=tk.E, padx=5, pady=5)

        # cc mail
        BOD_label = ttk.Label(root, text="BOD")
        BOD_label.grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)

        BOD_mail_entry = ttk.Entry(root)
        BOD_mail_entry.grid(column=1, row=3, sticky=tk.E, padx=5, pady=5)
    
        # hr mail 
        hr_label = ttk.Label(root, text="hr")
        hr_label.grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)

        hr_mail_entry = ttk.Entry(root)
        hr_mail_entry.grid(column=1, row=4, sticky=tk.E, padx=5, pady=5)

        # Save button
        Save_button = ttk.Button(root, text="Save", command=save_login_info)
        Save_button.grid(column=1, row=5, sticky=tk.E, padx=5, pady=5)

        #Check box to save to login 
        
        check_box = ttk.Checkbutton(root,text="Load saved information",command=load_saved_infor)
        check_box.grid(column=0, row=6, sticky=tk.E, padx=5, pady=5)
        
        window_opened = True
        root.protocol("WM_DELETE_WINDOW", on_child_window_close)
    root.mainloop()    
    
def create_table():
    conn = sqlite3.connect('login_database.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL,
            password TEXT NOT NULL,
            manager TEXT NOT NULL,
            BOD TEXT NOT NULL,
            hr TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

def save_login_info():
    email = username_entry.get()
    password = password_entry.get()
    manager = manager_mail_entry.get()
    BOD = BOD_mail_entry.get()
    hr = hr_mail_entry.get()
    if email == '' or password =='' or manager == '' or BOD == '' or hr =='':
        Warning().show_warning(warning_content='There is empty login information')
#     elif not re.match(BOD_mail_entry.get(), email_pattern) :
#         Warning().show_warning(warning_content=
#                                'Please check BOD email format\n\
# Emails are separated by semicolons'
#                                )
    else:    
        conn = sqlite3.connect('login_database.db')
        cursor = conn.cursor()
        cursor.execute('INSERT INTO users (email, password,manager,BOD,hr) VALUES (?,?,?,?,?)', (email, password, manager, BOD,hr))
        conn.commit()
        conn.close()
        Warning.show_message(message='Saved')

def load_saved_infor():
    conn = sqlite3.connect('login_database.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM users ORDER BY id DESC LIMIT 1')  # Lấy bản ghi mới nhất
    row = cursor.fetchone()
    if row:
        email, password, manager, BOD, hr = row[1], row[2], row[3], row[4],row[5]
        username_entry.delete(0, tk.END)  # Xóa nội dung hiện tại trong ô nhập liệu
        username_entry.insert(0, email)
        password_entry.delete(0, tk.END)
        password_entry.insert(0, password)
        manager_mail_entry.delete(0, tk.END)
        manager_mail_entry.insert(0, manager)
        BOD_mail_entry.delete(0, tk.END)
        BOD_mail_entry.insert(0, BOD)
        hr_mail_entry.delete(0, tk.END)
        hr_mail_entry.insert(0, hr)

    conn.close()
    
def call_login_db_newest():
    # connet to SQLite
    conn = sqlite3.connect('login_database.db')
    cursor = conn.cursor()
    # Truy vấn dữ liệu từ bảng (ví dụ: bảng 'users')
    cursor.execute('SELECT * FROM users ORDER BY id DESC LIMIT 1')# Lấy bản ghi mới nhất

    # Get data from excute
    data = cursor.fetchall()[0]
    return data


import tkinter as tk
from tkcalendar import DateEntry
from tkinter import PhotoImage,ttk,messagebox,filedialog
import openpyxl,datetime,re,threading,time
from GUI.initial import *
from GUI.collector import collector
from smtplib import SMTPAuthenticationError
import csv

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import policy
import smtplib,datetime,random
from pretty_html_table import build_table
import pandas as pd
 

#'Employee name','Position','Status','Start date','End date','Days','AM/PM','Start time','End time','Detail'
# Show warning
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

#Test time partern
class Check_time_partern():
    def __init__(self):
        pass
    def validate_time_input(user_input1,user_input2):
        time_pattern = r'^[0-2]?[0-9]:[0-5][0-9]$'  # Biểu thức chính quy cho "hh:mm"      
        if re.match(time_pattern, user_input1) and re.match(time_pattern, user_input2):
            Employee_name = Employee_name_entry.get()
            Position = Position_combobox.get()
            Status = Status_combobox.get()
            St_date = Start_date.get()
            En_date = End_date.get()
            Days = Days_entry.get()
            AMtoPM = AM_PM_combobox.get()
            St_time = Start_time_spinbox.get()
            En_time = End_time_spinbox.get()
            Details = Details_entry.get()

            row_values = [Employee_name,Position,Status,St_date,En_date,Days,AMtoPM,St_time,En_time,Details]
            #insert into treeview
            treeview.insert('',tk.END,values=row_values)       
        else:
              Warning().show_warning(warning_content="Time must be in 'hh:mm' format")  

def center_window(window):
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    window_width = window.winfo_width()
    window_height = window.winfo_height()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    window.geometry(f"{window_width}x{window_height}+{x}+{y}")

def random_request_number():
    random_number = random.randint(10**7,(10**8)-1)
    return random_number

def send_mail_with_progress():
    
    def update_progress():
        for i in range(101):
            progress_var.set(i)
            time.sleep(0.05)  
        progress_window.destroy()
        
    def send_mail_and_close_progress():
        try:
            sendmail() 
            Warning().show_message(message='Done')     

        except SMTPAuthenticationError as e:
            Warning().show_warning(warning_content='SMTP Authentication Error:\
Please check your email and password')

        except Exception as e:
            Warning().show_warning(warning_content=f"An error occurred: {str(e)}")

        progress_window.destroy()
    
    progress_window = tk.Toplevel()
    progress_window.title('Sending Mail Progress')
    progress_window.geometry('200x100')
    progress_window.resizable(0, 0)
    
    center_window(progress_window)

    progress_label = tk.Label(progress_window, text='Sending Mail...', font=("Arial", 12))
    progress_label.pack(pady=10)
    
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(progress_window, mode="indeterminate", variable=progress_var, maximum=100)
    progress_bar.pack(pady=5)
    
    progress_thread = threading.Thread(target=update_progress)
    progress_thread.start()
    
    # Call your sendmail function in a separate thread
    send_thread = threading.Thread(target=send_mail_and_close_progress)
    send_thread.start()
            
def sendmail():
    
    ut = datetime.date.today().strftime('%d/%m/%Y') #date time dd/mm/yyyy
    
    data = call_login_db_newest()

    # Send Mail

    application_info = []
    for item in treeview.get_children():
        row_data = treeview.item(item, "values")
        application_info.append(row_data)
    if application_info == []:
        Warning().show_warning(warning_content='Please check application form')
        return
    else:
        df = pd.DataFrame(application_info,columns=cols)
        html_table_blue_light =build_table(df,'green_light',width_dict=['150px','150px','150px','150px','150px',\
                                                                        '150px','150px','150px','150px','350px'],font_size=12)
   
    if application_info[0][1]=='Manager':
        sender = data[1] # Sender is manager
        recipient = data[4] # recipient is BOD
        cc=data[5] # cc is hr mail 
    
    else:
        sender = data[1] # Sender is staff
        recipient = data[3] # recipient is Manager
        cc=data[4]+';'+data[5] # cc are hr mail,BOD 
    # Create an email message
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = recipient
    msg['cc'] = cc
    msg['Subject'] = str(random_request_number())+'_'+application_info[0][0]+'_apply date: '+ut

    # Add the email body text with a hyperlink
    message = f'Dear Sir,<br>\
Please <a href="mailto:{sender}?cc={cc}&subject={msg["Subject"]}&body=Approved">click here</a> to approve this request.<br>\
or <a href="mailto:{sender}?cc={cc}&subject={msg["Subject"]}&body=Rejected">Reject</a><br>\
※For sender: If you want to cancel this request, Please <a href="mailto:{recipient}?cc={cc}&subject={msg["Subject"]}&body=Cancel">click here\
{html_table_blue_light}'
   
    # Encode the message as HTML
    msg.attach(MIMEText(message, "html", policy=policy.default))
   
    # # Attach the HTML content
    # msg.attach(MIMEText(html_table_blue_light, 'html'))
    
    server = smtplib.SMTP('smtp-mail.outlook.com', port=587)
    server.starttls()
    server.login(sender, data[2])
    server.send_message(msg)
    server.quit()

def contact_to_admin():
    Warning().show_warning(warning_content='User accounts do not have this feature.\n\
Please check with admin')

# Test input data is dd/MM/YYYY
def is_valid_date(date_str):
    try:
        datetime.datetime.strptime(date_str, '%d/%m/%Y')
        return True
    except ValueError:
        return False
def valid_start_date_input():
    user_input = Start_date.get()
    if is_valid_date(user_input):
        pass
    else:
        Warning().show_warning(warning_content="Dates must be in the format 'dd/MM/YYYY'")
def valid_end_date_input():
    user_input = End_date.get()
    if is_valid_date(user_input):
        pass
    else:
        Warning().show_warning(warning_content="Dates must be in the format 'dd/MM/YYYY'")
        
# delete button F8
def on_key_del(event):
    if event.keysym == 'F8':
        delete_selected_row()

# insert button F4
def on_key_insert(event):
    if event.keysym == 'F4':
       insert()

# select all row
def select_all_rows(event):
    all_items = treeview.get_children()
    treeview.selection_set(all_items)

def clear_selection():
    treeview.selection_remove(treeview.selection())
def on_key_esc(event):
    if event.keysym == 'Escape':
        clear_selection()
# Click mouse 1
def mouse_1_click(event):
    if event.keysym == '<ButtonRelease-1>':
        copy_selected_row()

def open_excel_load_data():
    #open dialog 
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx'),('Excel Files', '*.xls'),('Excel Files', '*.csv')])
    if file_path:
        if file_path.endswith('.csv'):
            # Handle CSV file
            with open(file_path, 'r', newline='') as csv_file:
                csv_reader = csv.reader(csv_file)
                list_values = list(csv_reader)
        else:
            workbook = openpyxl.load_workbook(file_path) # Open work book
            sheet = workbook.active 
            list_values = list(sheet.values) # Get active workbook data
            new_list_values = []
        
            for row in list_values:
                # Check if all elements in the row are None
                all_none = all(item is None for item in row)
                if not all_none:
                    new_list_values.append(row)
        
                # Replace list_values with the filtered list
                list_values = new_list_values
        if not list_values:
            Warning().show_warning(warning_content="The selected file is empty.")
            return
        
        if len(list_values[0]) != len(cols): 
           Warning().show_warning(warning_content="Please check import file !!!") 
        
        elif len(list_values[0]) == len(cols):
            for i in range(len(cols)):
                if cols[i] != list_values[0][i]:
                    Warning().show_warning(warning_content="Please check import file !!!")
                    break 
            for i1 in range(len(list_values)):
                all_none = True  # Ban đầu giả sử tất cả đều là None cho mỗi hàng
                for item in list_values[i1]:
                    if item is not None:
                        all_none = False  # Nếu có một phần tử không phải None, đặt biến all_none thành False
                        break  # Thoát khỏi vòng lặp nếu có phần tử khác None
                if all_none:
                    Warning().show_warning(warning_content="Please check import file !!!\n\
The file contains the line blank")
                    break
            else:
                for col_name in list_values[0]:
                    treeview.heading(col_name,text=col_name)
                for value_tuple in list_values[1:]:
                    change_list = list(value_tuple)
                    for i in range(len(change_list)):
                        if change_list[i] == None:
                            change_list[i] = ''
                    change_list[3] = change_list[3].strftime('%d/%m/%Y')
                    change_list[4] = change_list[4].strftime('%d/%m/%Y')
                    if isinstance(change_list[7], str):
                        pass
                    else:
                        change_list[7] = change_list[7].strftime('%H:%M')
                    if isinstance(change_list[8], str):
                        pass
                    else:
                        change_list[8] = change_list[8].strftime('%H:%M')

                    value_tuple = tuple(change_list)
                    treeview.insert('', tk.END, values=value_tuple)
        
        else:
            pass            
    return        
    # Load data          
    
def insert():
    if Employee_name_entry.get() == '' or Position_combobox.get() == '' \
        or Status_combobox.get() =='':
        Warning().show_warning(warning_content="Can't insert blank data\n\
Please check 'Employee name','Position','Status'")
        
    elif Position_combobox.get() not in options \
        or Status_combobox.get() not in status:
        
        Warning().show_warning(warning_content="The position must be either 'Manager' or 'Staff'\n\
The Status must be either 'Paid leave','Unpaid leave','Over Time','Late/Leave early'\n\
Please check 'Position','Status'")
        
    elif Status_combobox.get()=='Over Time' and (Start_time_spinbox.get() ==''or End_time_spinbox.get() ==''):
        Warning().show_warning(warning_content="Can't insert blank data\nPlease check 'Start_time','End_time'")

    elif Status_combobox.get()=='Over Time' and (Start_time_spinbox.get() !='' or End_time_spinbox.get() !=''):
        Check_time_partern.validate_time_input(Start_time_spinbox.get(),End_time_spinbox.get())

    elif Days_entry.get():
        try: 
            float_value = float(Days_entry.get())
            
            Employee_name = Employee_name_entry.get()
            Position = Position_combobox.get()
            Status = Status_combobox.get()
            St_date = Start_date.get()
            En_date = End_date.get()
            Days = Days_entry.get()
            AMtoPM = AM_PM_combobox.get()
            St_time = Start_time_spinbox.get()
            En_time = End_time_spinbox.get()
            Details = Details_entry.get()

            row_values = [Employee_name,Position,Status,St_date,En_date,Days,AMtoPM,St_time,En_time,Details]
            # insert into treeview
            treeview.insert('',tk.END,values=row_values)
                
        except ValueError:
            Warning().show_warning(warning_content="Can't insert 'Days'\nDays must be float")
        pass
    else:
        Employee_name = Employee_name_entry.get()
        Position = Position_combobox.get()
        Status = Status_combobox.get()
        St_date = Start_date.get()
        En_date = End_date.get()
        Days = Days_entry.get()
        AMtoPM = AM_PM_combobox.get()
        St_time = Start_time_spinbox.get()
        En_time = End_time_spinbox.get()
        Details = Details_entry.get()

        row_values = [Employee_name,Position,Status,St_date,En_date,Days,AMtoPM,St_time,En_time,Details]
        #insert into treeview
        treeview.insert('',tk.END,values=row_values)
        return

def copy_selected_row(event):
    selected_items = treeview.selection()
    for selected_item in selected_items:
        row_data = treeview.item(selected_item, 'values')

        clear_record()

        Employee_name_entry.insert(0,row_data[0])
        Position_combobox.insert(0,row_data[1])
        Status_combobox.insert(0,row_data[2])
        Start_date.insert(0,row_data[3])
        End_date.insert(0,row_data[4])
        Days_entry.insert(0,row_data[5])
        AM_PM_combobox.insert(0,row_data[6])
        Start_time_spinbox.insert(0,row_data[7])
        End_time_spinbox.insert(0,row_data[8])
        Details_entry.insert(0,row_data[9])

    return 
def delete_selected_row():
    selected_item = treeview.selection()  # Get the selected item(s)
    for item in selected_item:
        treeview.delete(item) # Delete the selected item(s)

def clear_record():
    Employee_name_entry.delete(0, "end")
    Position_combobox.delete(0, "end")
    Status_combobox.delete(0, "end")
    Start_date.delete(0, "end")
    End_date.delete(0, "end")
    Days_entry.delete(0, "end")
    AM_PM_combobox.delete(0, "end")
    Start_time_spinbox.delete(0, "end")
    End_time_spinbox.delete(0, "end")
    Details_entry.delete(0, "end")
    return

# Frame 1

def create_employee_table():
    conn = sqlite3.connect('employee_database.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS employee_list (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Name TEXT NOT NULL,
            Position TEXT NOT NULL
        )        
                   ''')
    conn.commit()
    conn.close()

def add_employee_name():
    epl_name = text_entry.get()
    position = combobox.get()

    if epl_name =='' or combobox =='':
        Warning().show_warning(warning_content="Can't input blank data")
    else:
        conn = sqlite3.connect('employee_database.db')
        cursor = conn.cursor()
        cursor.execute('INSERT INTO employee_list(Name,Position) VALUES(?,?)'
                       ,(epl_name,position))
        conn.commit()
        conn.close()
        Warning().show_message(message="Sign Up Success")
# frame 4 
def call_db():
    conn = sqlite3.connect('employee_database.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM employee_list')

    employee_list = []
    items = cursor.fetchall()
    conn.close()
    
    for item in items:
        employee_list.append(item)
    
    return employee_list

def match_string():
    employee_list = call_db()
    hits = []
    got = auto.get()
    for i in range(len(employee_list)):
        if employee_list[i][1].strip().startswith(got):
            hits.append(employee_list[i][1].strip())
    return hits    

def get_typed(event):
    if len(event.keysym) == 1:
        hits = match_string()
        show_hit(hits)

def show_hit(lst):
    if len(lst) == 1:
        auto.set(lst[0])
        detect_pressed.filled = True

def detect_pressed(event):    
    key = event.keysym
    if len(key) == 1 and detect_pressed.filled is True:
        pos = Employee_name_entry.index(tk.INSERT)
        Employee_name_entry.delete(pos, tk.END)

detect_pressed.filled = False

def dash_board():

    global Employee_name_entry, Position_combobox, Status_combobox, Start_date, End_date,\
            Days_entry, AM_PM_combobox, Start_time_spinbox, End_time_spinbox, treeview,\
            cols, Details_entry, options, status, text_entry, combobox, auto

    app = tk.Tk()
    app.title("HR SYSTEM AITC")
    app.resizable(0,0)

    # #image object
    # AITC_photo = PhotoImage(file='images/aitc.png').subsample(10)

    frame1 = tk.Frame(app,height=50,width=1400,pady=5, padx=5,borderwidth=1, relief="groove")
    frame1.propagate(False)

    # image_label = tk.Label(frame1, image=AITC_photo)
    # image_label.pack(side='right')

    ## Employee name
    label1 = tk.Label(
        frame1,
        text ='Employee name',
        font = ('Times New Roman',12),width= 15, height = 1)
    label1.pack(side='left')

    text_entry = tk.Entry(frame1,font = ('Times New Roman',12))
    text_entry.pack(side='left',padx=10)

    ## Position
    label2 = tk.Label(
        frame1,
        text ='Position',
        font = ('Times New Roman',12),width= 15, height = 1)
    label2.pack(side='left',padx=10)

    options = ["Manager", "Staff"]
    combobox = ttk.Combobox(frame1, values=options,font = ('Times New Roman',12))
    combobox.pack(side='left',padx=10)

    ## Add
    add_button=tk.Button(frame1,text='Add',font=('Times New Roman',12), bg='blue',fg='white',command=add_employee_name)
    add_button.pack(side='left',padx=10)

    ## Show list Nhan vien 
    show_button=tk.Button(frame1,text='Show list Employee',font=('Times New Roman',12), bg='blue',fg='white',command=contact_to_admin)
    show_button.pack(side='left',padx=10)

    ## Login Status
    login_status=tk.Label(frame1,text='Please input email, password to login',font=('Times New Roman',12))
    login_status.pack(side='right',padx=10)

    frame1.grid(row=0,column=0,padx=2,pady=2)

    # frame 2
    frame2 = tk.Frame(app,height=50,width=1400, pady=5,padx=5,borderwidth=1, relief="groove")
    frame2.propagate(False)

    label3 = tk.Label(
        frame2,
        text ='F8 - Delete row, F4 - Insert',
        font = ('Times New Roman',12),)
    label3.pack(side='right',padx=10)

    frame2.grid(row=1,column=0,padx=2,pady=2)


    #frame 3
    frame3 = tk.LabelFrame(app,height=400,width=1400, pady=10,padx=10,borderwidth=1, relief="groove",text='Application Form')
    frame3.propagate(False)
    frame3.grid(row=2,column=0,padx=2,pady=2)


    # view, input data
    treeFrame = ttk.Frame(frame3)
    treeFrame.grid(row=0,column=1,pady=10)
    treeScroll = ttk.Scrollbar(treeFrame)
    treeScroll.pack(side='right',fill='y')

    cols = ['Employee name','Position','Status','Start date','End date','Days','Balance','AM/PM','Start time','End time','Detail']
    treeview=ttk.Treeview(treeFrame,show="headings",
                        yscrollcommand=treeScroll.set,columns=cols,height=13)

    for col_name in cols:
        treeview.heading(col_name,text=col_name)
        treeview.column(col_name,width=140)

    treeview.pack()
    treeScroll.config(command=treeview.yview)

    treeview.bind("<F8>", on_key_del)
    treeview.bind("<F4>", on_key_insert)
    treeview.bind("<Control-a>", select_all_rows)
    treeview.bind("<Escape>", on_key_esc)
    treeview.bind("<ButtonRelease-1>", copy_selected_row)

    #frame 4
    frame4 = tk.LabelFrame(app,height=100,width=1400, pady=10,padx=10,borderwidth=1, relief="groove",text='Record')
    frame4.propagate(False)
    frame4.grid(row=3,column=0,padx=2,pady=2)

    ## Input Name
    auto = tk.StringVar()
    Employee_name_entry = tk.Entry(frame4,width= 15,font=('Times New Roman', 12),textvariable=auto)
    Employee_name_entry.insert(0,'Tokugawa Ieyasu')
    # Employee_name_entry.bind("<FocusIn>",lambda e: Employee_name_entry.delete('0','end'))
    Employee_name_entry.grid(row=0,column=0,sticky='w',padx=1,pady=1)


    Employee_name_entry.focus_set()
    Employee_name_entry.bind('<KeyRelease>', get_typed)
    Employee_name_entry.bind('<Key>', detect_pressed)
    Employee_name_entry.bind('<ButtonRelease-1>', )

    ## Position
    Position_combobox = ttk.Combobox(frame4,values=options,width= 15)
    Position_combobox.grid(row=0,column=1,sticky='w',padx=1,pady=1)

    ## Status
    status = ['Paid leave','Unpaid leave','Over Time','Late/Leave early','Forgot card']
    Status_combobox = ttk.Combobox(frame4,values=status,width= 15)
    Status_combobox.grid(row=0,column=2,sticky='w',padx=1,pady=1)

    ##Start date 
    Start_date = DateEntry(frame4, width=15, borderwidth=2,locale='en_US',date_pattern='dd/MM/yyyy')
    Start_date.grid(row=0,column=3,sticky='w',padx=1,pady=1)

    ##End date
    End_date = DateEntry(frame4, width=15, borderwidth=2,locale='en_US',date_pattern='dd/MM/yyyy')
    End_date.grid(row=0,column=4,sticky='w',padx=1,pady=1)

    ##Days
    Days_entry = ttk.Entry(frame4,width= 15)
    Days_entry.insert(0,0.5)
    Days_entry.bind("<FocusIn>",lambda e: Days_entry.delete('0','end'))
    Days_entry.grid(row=0,column=5,sticky='w',padx=1,pady=1)

    ##AM/PM
    AMtoPM = ['AM','PM','All day']
    AM_PM_combobox = ttk.Combobox(frame4,values=AMtoPM,width= 15)
    AM_PM_combobox.grid(row=0,column=6,sticky='w',padx=1,pady=1)

    ##Start time 
    time_values = [f"{hour:02d}:{minute:02d}" for hour in range(24) for minute in range(0, 60, 15)]
    Start_time_spinbox = ttk.Spinbox(frame4, values= time_values,width= 14)
    Start_time_spinbox.grid(row=0,column=7,sticky='w',padx=1,pady=1)

    ##End time
    End_time_spinbox = ttk.Spinbox(frame4, values= time_values,width= 14)
    End_time_spinbox.grid(row=0,column=8,sticky='w',padx=1,pady=1)
    ##Detail
    Details_entry = ttk.Entry(frame4,width= 15)
    Details_entry.insert(0,'Personal reason')
    # Details_entry.bind("<FocusIn>",lambda e: Details_entry.delete('0','end'))
    Details_entry.grid(row=0,column=9,sticky='w',padx=1,pady=1)

    ## insert button 
    insert_button = ttk.Button(frame4,text='Insert',command=insert)
    insert_button.grid(row=2,column=0,sticky='nsew',padx=1,pady=1)
    insert_button.bind("<F4>", on_key_insert)

    ## clear button 
    clear_button = ttk.Button(frame4,text='Clear',command=clear_record)
    clear_button.grid(row=2,column=1,sticky='nsew',padx=1,pady=1)


    ## delete button
    delete_button = ttk.Button(frame4,text='Delete',command=delete_selected_row)
    delete_button.grid(row=2,column=2,sticky='nsew',padx=1,pady=1)

    ## Import button
    import_button = ttk.Button(frame4,text='Import',command=open_excel_load_data)
    import_button.grid(row=2,column=3,sticky='nsew',padx=1,pady=1)

    #frame 5
    frame5 = tk.Frame(app,height=50,width=1400, pady=10,padx=10,borderwidth=1, relief="groove")
    frame5.propagate(False)

    Initialize_bt=tk.Button(frame5,text='Inititalize',font=('Times New Roman',12), bg='blue',fg='white' ,width= 15, height = 1,command=initialize)
    Initialize_bt.pack(padx=10,side='left')

    Send_Application_bt=tk.Button(frame5,text='Send Request',font=('Times New Roman',12), bg='blue',fg='white' ,width= 15, height = 1,command=send_mail_with_progress)
    Send_Application_bt.pack(padx=10,side='left')

    Collect_bt=tk.Button(frame5,text='Collect',font=('Times New Roman',12), bg='blue',fg='white' ,width= 15, height = 1,command=collector)
    Collect_bt.pack(padx=10,side='left')

    frame5.grid(row=4,column=0,padx=2,pady=2)

    create_table()
    create_employee_table()
    app.mainloop()

    return

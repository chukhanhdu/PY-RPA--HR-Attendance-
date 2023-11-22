import win32com.client
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

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

def collector():
    try:  
        outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(6) # 6 inbox


        folders = inbox.Folders

        messages = folders('HR AITC VN').Items

        header = ['Employee name','Position','Status','Start date','End date','Days','AM/PM',
                'Start time','End time','Detail','Receivedtime','Subject','Sender','Approved stastus']

        df = pd.DataFrame(columns = header)

        i = 0
        for message in messages: 
            body = message.HTMLBody
            soup = BeautifulSoup(body,"html.parser")
            table = soup.find('table',{'class':'dataframe'})
            
            # Get application form line
            if table :
                xs = table.find('tbody').find_all('td')
                data = []
                for x in xs:
                    data.append(x.text)
                for y in range(int(len(data)/10)):
                    list_0 = data[y*10:y*10+10]
            else:
                list_0 = ['' for _ in range(10)]
            
            # Get received time
            ReceivedTime = datetime.strptime(str(message.ReceivedTime)[:-6],
                                                "%Y-%m-%d %H:%M:%S.%f").strftime(
                                                "%d/%m/%Y %H:%M")
            
            # Get mail subject
            mail_subject = message.Subject

            # Get mail Sender
            mail_sender = str(message.Sender)
            list_0.extend([ReceivedTime,mail_subject,mail_sender])
            # Get Approved status
            for soup in soup:
                data = []
                data.append(soup.text.strip())
                if data[0] == 'Approved' or data[0] == 'Rejected'\
                    or data[0] == 'Cancel':
                    list_0.append(data[0])
                else:
                    list_0.append('')
                    break
            df.loc[len(df)] = list_0

            i +=1
        df.to_csv('Collected file.csv')
        Warning().show_message(message='Done')
    except Exception as e:
        Warning().show_warning(warning_content=str(e))
    return
collector()
print("chuong trinh da thuc hien xong")
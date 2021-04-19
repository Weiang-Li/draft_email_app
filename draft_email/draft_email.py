import sqlite3
import os.path
from tkinter import *
from PIL import ImageTk
from tkinter import messagebox
from tkinter import ttk
from email.mime.text import MIMEText
import datetime
from datetime import date
import imaplib
import time
from email.mime.multipart import MIMEMultipart
from tkcalendar import *


class EmailDatabase():
    def __init__(self):
        self.conn = sqlite3.connect('Email.db')
        self.c = self.conn.cursor()
        print('database connceted')

    def query(self):
        return self.c

    def create_table(self):
        try:
            self.c.execute("""CREATE TABLE Email (
            email_title text,
            email_to text,
            email_cc text,
            email_message text,
            email_username text,
            email_start_on text,
            Frequency text
        )
        """)
            self.conn.commit()

        except:
            print('Table already exists')

    def insert(self, email_title=None, email_to=None, email_cc=None, email_message=None, email_username=None,
               email_start=None,
               Frequency=None):
        with self.conn:
            self.c.execute(
                "INSERT INTO Email VALUES (:email_title, :email_to, :email_cc, :email_message, :email_username, :email_start, :Frequency)",
                {'email_title': email_title, 'email_to': email_to, 'email_cc': email_cc, 'email_message': email_message,
                 'email_username': email_username, 'email_start': email_start, 'Frequency': Frequency})

    def get_report_by_reportname(self, email_title):
        with self.conn:
            res = self.c.execute("SELECT * FROM Email WHERE email_title = :email_title",
                                 {'email_title': email_title}).fetchall()
            return res

    def update_report(self, email_title, email_to=None, email_cc=None, email_message=None, email_username=None):
        with self.conn:
            self.c.execute("""UPDATE email SET email_title = :email_title,
                    email_to = :email_to,
                    email_cc = :email_cc,
                    email_username = :email_username,
                    email_email_attachment = :email_attachment

                    WHERE email_title = :email_title
            """, {'email_title': email_title, 'email_to': email_to, 'email_cc': email_cc,
                  'email_message': email_message, 'email_username': email_username})

    def remove_report(self, email_title):
        with self.conn:
            self.c.execute("DELETE from Email WHERE email_title=:email_title", {'email_title': email_title})

    def delete_all_databsae(self, email_username):
        with self.conn:
            delete_all_info = self.c.execute("DELETE FROM Email WHERE email_username=:email_username",
                                             {'email_username': email_username})

    def delete_data(self):
        with self.conn:
            delete = self.c.execute("Del FROM Email").fetchall()
            return delete

    def delete_one(self, email_title):
        with self.conn:
            delete_one = self.c.execute("DELETE FROM Email WHERE email_title=:email_title",
                                        {'email_title': email_title})

    def search(self):
        with self.conn:
            search = "SELECT * FROM Email WHERE email_title LIKE '%'+'%"

    def get_frequency(self, email_username):
        with self.conn:
            frequency = self.c.execute("SELECT Frequency FROM Email WHERE email_username=:email_username",
                                       {'email_username': email_username}).fetchall()
            self.conn.commit()
            return frequency

    def get_email_start_on(self, email_username):
        with self.conn:
            start = self.c.execute("SELECT email_start_on FROM Email WHERE email_username=:email_username",
                                   {'email_username': email_username}).fetchall()
            self.conn.commit()
            return start

    def show_all_database(self, email_username):
        with self.conn:
            all_info = self.c.execute("SELECT * FROM Email WHERE email_username=:email_username",
                                      {'email_username': email_username}).fetchall()
            self.conn.commit()
            return all_info


class EmailDrafter:
    def __init__(self, root):
        self.root = root
        self.root.title('Gmail Login System')
        self.root.geometry('400x500')

        # ----------All Images --------

        self.bg_icon = ImageTk.PhotoImage(file="C:/your path/your image.png")

        # ---------Variables----------
        self.username_input = StringVar()
        self.password_input = StringVar()

        bg_lbl = Label(self.root, image=self.bg_icon).pack()
        # title = Label(self.root, text = 'Login System', font=('time new roman',20)).place(x=0,y=0,relwidth=1)

        self.Login_Frame = Frame(self.root, bg='white').place(y=0, x=0)

        lbluser = Label(self.Login_Frame, text='Username').place(x=100, y=150)
        txtusername = Entry(self.Login_Frame, textvariable=self.username_input, bd=5, relief=GROOVE).place(x=180, y=150)

        lblpassword = Label(self.Login_Frame, text='Password').place(x=100, y=200)
        txtpassword = Entry(self.Login_Frame, bd=5, textvariable=self.password_input, relief=GROOVE)
        txtpassword.place(x=180, y=200)
        txtpassword.config(show="*")

        btn_login = Button(self.Login_Frame, text='Login', command=self.login, width=18).place(x=180, y=260)

        # -------------------connect to saved email database ---------------------
        self.conn = EmailDatabase()
        self.conn.create_table()

        # -------------------connect to gmail-------------------------------------
        self.gmail_conn = imaplib.IMAP4_SSL('imap.gmail.com', port=993)

    def login(self):
        try:
            print(self.username_input.get(), self.password_input.get())
            self.gmail_conn.login(self.username_input.get(), self.password_input.get())
            self.gmail_conn.select('[Gmail]/Drafts')
            print('login successful')
            self.root.destroy()
            self.form()

        except:
            if self.username_input.get() == '' or self.password_input.get() == '':
                messagebox.showerror('Error', 'Incorrect Email')
            else:
                messagebox.showerror('Error', 'Incorrect Email')

    def draft(self, email_title, email_from, email_to, email_cc, message):
        emails = self.conn.show_all_database(self.username_input.get())
        msg = MIMEMultipart('alternative')
        msg['Subject'] = email_title
        msg['from'] = self.username_input.get()
        msg['to'] = email_to
        msg['cc'] = email_cc
        # msg['bcc'] = email_bcc

        html = """Hello,  <br><br>""" + message
        message = MIMEText(html, 'html')
        msg.attach(message)
        msg.as_string()
        now = imaplib.Time2Internaldate(time.time())
        self.gmail_conn.append('[Gmail]/Drafts', "", now, str(msg).encode('utf-8'))

    def form(self):
        self.Form = Tk()
        self.Form.title('Email Draft Saver')
        self.Form.geometry('1300x1000')
        self.bg_icon = ImageTk.PhotoImage(file="C:/your path/your image.png")
        Form_background = Label(self.Form, image=self.bg_icon).pack()
        self.Frame = Frame(self.Form, bg='white').place(x=0, y=0)
        # -----------------Label-----------------

        self.Email_title_Label = Label(self.Frame, text='Email Title').place(relx=0.2, rely=0.1)
        self.Email_to_Label = Label(self.Frame, text='Email To').place(relx=0.2, rely=0.15)
        self.Email_cc_Label = Label(self.Frame, text='Email CC').place(relx=0.2, rely=0.2)
        self.Email_Message_Label = Label(self.Frame, text='Email Message').place(relx=0.2, rely=0.25)
        self.Email_start = Label(self.Frame, text='Email Start On').place(relx=0.2, rely=0.3)
        self.Frequency_Label = Label(self.Frame, text='Frequency').place(relx=0.2, rely=0.35)

        # ---------------Entry------------------
        self.Email_title_Entry = Entry(self.Frame, text='Email Title')
        self.Email_title_Entry.place(relx=0.3, rely=0.1,
                                     relwidth=0.5)
        self.Email_to_Entry = Entry(self.Frame, text='Email To')
        self.Email_to_Entry.place(relx=0.3, rely=0.15,
                                  relwidth=0.5)

        self.Email_cc_Entry = Entry(self.Frame, text='Email CC')
        self.Email_cc_Entry.place(relx=0.3, rely=0.2,
                                  relwidth=0.5)

        self.Email_Message_Entry = Entry(self.Frame, text='Email Message')
        self.Email_Message_Entry.place(relx=0.3, rely=0.25,
                                       relwidth=0.5)

        self.Email_start_calender = DateEntry(self.Frame, selectmode='day')
        self.Email_start_calender.place(relx=0.3, rely=0.3, relwidth=0.5)

        self.Frequency_dropdown = StringVar()
        self.Frequency_dropdown.set("Frequency")
        self.Frequency_Entry = OptionMenu(self.Frame, self.Frequency_dropdown, "Daily", "Weekly-Monday",
                                          "Weekly-Tuesday",
                                          "Weekly-Tuesday", "Weekly-Wednesday", "Weekly-Thursday", "Weekly-Friday",
                                          "Biweekly-Monday", "Biweekly-Tuesday", "Biweekly-Tuesday",
                                          "Biweekly-Wednesday", "Biweekly-Thursday", "Biweekly-Friday")
        self.Frequency_Entry.place(relx=0.3, rely=0.35, relwidth=0.5)

        self.search_bar = Entry(self.Frame, text='Search by email title')
        self.search_bar.place(relx=0.5, rely=0.41, relheight=0.03, relwidth=0.2)

        # --------------------Button---------------------------#
        run = Button(self.Frame, text='Run', command=self.run).place(relx=0.8, rely=0.5, relheight=0.05,
                                                                     relwidth=0.1)

        save_btn = Button(self.Frame, text='Save Email', command=self.save).place(relx=0.32, rely=0.5,
                                                                                  relheight=0.05, relwidth=0.1)
        delete_btn = Button(self.Frame, text='Delete All Emails', command=self.delete_all_save_report).place(
            relx=0.68, rely=0.5,
            relheight=0.05, relwidth=0.1)

        show_all_saved_emails_btn = Button(self.Frame, text='Show all emails',
                                           command=self.show_all_saved_report).place(relx=0.44, rely=0.5,
                                                                                     relheight=0.05,
                                                                                     relwidth=0.1)

        delete_one = Button(self.Frame, text='Delete one email', command=self.delete_one).place(relx=0.56,
                                                                                                rely=0.5,
                                                                                                relheight=0.05,
                                                                                                relwidth=0.1)

        update_email = Button(self.Frame, text='Update email', command=self.update).place(relx=0.20, rely=0.5,
                                                                                          relheight=0.05,
                                                                                          relwidth=0.1)

        search = Button(self.Frame, text='Search by email title', command=self.search).place(relx=0.37, rely=0.41)

        self.search = StringVar()
        self.t1 = StringVar()
        self.t2 = StringVar()
        self.t3 = StringVar()
        self.t4 = StringVar()
        self.t5 = StringVar()
        self.t6 = StringVar()
        self.t7 = StringVar()

        self.Form.mainloop()

    def run(self):
        emails = self.conn.show_all_database(self.username_input.get())
        frequency = self.conn.get_frequency(self.username_input.get())
        for email in emails:
            if (datetime.datetime.today() >= datetime.datetime.strptime(email[5], '%Y-%m-%d')) and email[6] == 'Daily':
                self.draft(email[0], self.username_input.get(), email[1], email[2], email[3])
                print('daily')

            elif (datetime.datetime.today() >= datetime.datetime.strptime(email[5], '%Y-%m-%d')) and (
                    date.today().weekday() == 0) and (email[6] == 'Weekly-Monday'):
                self.draft(email[0], self.username_input.get(), email[1], email[2], email[3])
                print('weekly monday')

            elif (datetime.datetime.today() >= datetime.datetime.strptime(email[5], '%Y-%m-%d')) and (
                    date.today().weekday() == 1) and (email[6] == 'Weekly-Tuesday'):
                self.draft(email[0], self.username_input.get(), email[1], email[2], email[3])
                print('weekly tuesday')

            elif (datetime.datetime.today() >= datetime.datetime.strptime(email[5], '%Y-%m-%d')) and (
                    date.today().weekday() == 2) and (email[6] == 'Weekly-Wednesday'):
                self.draft(email[0], self.username_input.get(), email[1], email[2], email[3])
                print('weeky wednesday')

            elif (datetime.datetime.today() >= datetime.datetime.strptime(email[5], '%Y-%m-%d')) and (
                    date.today().weekday() == 3) and (email[6] == 'Weekly-Thursday'):
                self.draft(email[0], self.username_input.get(), email[1], email[2], email[3])
                print('weekly thursday')

            elif (datetime.datetime.today() >= datetime.datetime.strptime(email[5], '%Y-%m-%d')) and (
                    date.today().weekday() == 4) and (email[6] == 'Weekly-Friday'):
                self.draft(email[0], self.username_input.get(), email[1], email[2], email[3])
                print('weekly friday')

            elif (datetime.datetime.today() - datetime.datetime.strptime(email[5], '%Y-%m-%d')) % datetime.timedelta(
                    days=14) == datetime.timedelta(days=0):
                self.draft(email[0], self.username_input.get(), email[1], email[2], email[3])
                print('biweekly', date.today(), self.Email_start_calender.get_date())
        messagebox.askyesno("Notification", "Done!")

    def save(self):
        if self.Email_title_Entry.get() != '':
            self.conn.insert(self.Email_title_Entry.get(), self.Email_to_Entry.get(), self.Email_cc_Entry.get(),
                             self.Email_Message_Entry.get(), self.username_input.get(),
                             self.Email_start_calender.get_date(), self.Frequency_dropdown.get())

            self.Email_title_Entry.delete(0, END)
            self.Email_to_Entry.delete(0, END)
            self.Email_cc_Entry.delete(0, END)
            self.Email_Message_Entry.delete(0, END)
            self.Email_start_calender.configure(validate='none')
            self.Frequency_dropdown.set("Frequency")

            self.show_all_saved_report()

        else:
            return True

    def update(self):
        pass

    def data_table(self):
        self.table = ttk.Treeview(self.Frame, columns=(1, 2, 3, 4, 5, 6, 7), show='headings')
        self.table.place(relwidth=0.9, relheight=0.3, rely=0.6, relx=0.05)
        self.table.heading(1, text='Email Title')
        self.table.column(1, width=70)

        self.table.heading(2, text='Email To')
        self.table.column(2, width=70)

        self.table.heading(3, text='Email CC')
        self.table.column(3, width=70)

        self.table.heading(4, text='Email Message')
        self.table.column(4, width=70)

        self.table.heading(5, text='Email Username')
        self.table.column(5, width=70)

        self.table.heading(6, text='Email Start On')
        self.table.column(6, width=70)

        self.table.heading(7, text='Frequency')
        self.table.column(7, width=70)

        # self.table.delete(*self.table.get_children())   #to clear out table interface
        # self.table.bind('<Double 1>', self.getrow)       # double click
        self.table.bind('<<TreeviewSelect>>', self.getrow)  # single click

    def getrow(self, event):
        rowid = self.table.identify_row(event.y)
        self.item = self.table.item(self.table.focus())

        self.t1.set(self.item['values'][0])
        self.t2.set(self.item['values'][1])
        self.t3.set(self.item['values'][2])
        self.t4.set(self.item['values'][3])
        self.t5.set(self.item['values'][4])
        self.t6.set(self.item['values'][5])
        self.t7.set(self.item['values'][6])
        print(self.item['values'][0])

    def show_all_saved_report(self):
        records = self.conn.show_all_database(self.username_input.get())

        self.data_table()
        for i in records:
            self.table.insert('', 'end', values=i)

    def search(self):
        self.data_table()
        search = self.search_bar.get()
        res = self.conn.get_report_by_reportname(search)
        for i in res:
            print(i)
            self.table.insert('', 'end', values=i)

    def delete_all_save_report(self):
        if messagebox.askyesno("Confirm Delete?", "Are you sure you want to delete ALL email draft?"):
            self.conn.delete_all_databsae(self.username_input.get())
            self.table.delete(*self.table.get_children())
        else:
            return True

    def delete_one(self):
        # delete_by_report_name = self.search_bar.get()
        try:
            selected = self.item['values'][0]
            if messagebox.askyesno("Confirm Delete?", "Are you sure you want to delete" + self.item['values'][0] + " ?"):
                res = self.conn.delete_one(selected)
                self.table.delete(*self.table.get_children())
                self.show_all_saved_report()
            else:
                pass
        except:
            messagebox.askyesno("Information", "Please click on the email you want to delete")

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


root = Tk()
image = resource_path("C:/your path/your image.png")
obj = EmailDrafter(root)

root.mainloop()

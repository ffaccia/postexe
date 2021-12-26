
import requests
import sys
import os
import tkinter as tk
import logging
from datetime import datetime as dt
from tkinter import ttk
import tkinter as tk
import sqlite3

FORMAT = "[ %(asctime)s, %(levelname)s] %(message)s" 
#logging.basicConfig(filename = logName, level=logging.DEBUG, format=FORMAT)

dbFile = None
conn = None
cur = None
CODES_OK = None
extensions = None
MULTI_FILES = False
entry_dir = None
   
def setup_profile():    
    global extensions, CODES_OK, MULTI_FILES
    url = 'https://httpbin.org/post'
    extensions = ['pdf']
    CODES_OK = [ v for k,v in requests.codes.__dict__.items() if k in ['OK','CREATED','ACCEPTED']]
    MULTI_FILES = True
    
    user_profile_dir = os.environ["USERPROFILE"]
    inner_dir = os.path.join("documents","sviluppo","python","postexe")
    entry_dir = os.path.join(user_profile_dir, inner_dir)
    if os.path.isdir(entry_dir) == False:
          msg = "Setup dir is missing: %s! " % (entry_dir)
          tk.messagebox.showerror("Setup Error", msg)
          #win32api.MessageBox(0, msg, "Critical Error", 0x00001000)         
          raise InputError(msg, msg)    
    
    os.chdir(entry_dir)
    
    resp_dir = os.path.join(".","responses")
    if os.path.isdir(resp_dir) == False:
        os.mkdir(resp_dir, 755);
    
    file_log = os.path.join(".", "responses", "response.txt")
    fh_out = open(file_log,"a")


 
class Error(Exception):
    """Base class for exceptions in this module."""
    pass

class InputError(Error):
    """Exception raised for errors in the input.

    Attributes:
        expression -- input expression in which the error occurred
        message -- explanation of the error
    """

    def __init__(self, expression, message):
        self.expression = expression
        self.message = message
        
        

def get_timestamp():
    return dt.now().strftime(format='%Y-%m-%d %H:%M:%S.%f')[:-3]
    
def setup_connection():
    global dbFile, conn, cur
    dbFile = os.path.join(".", "db", "coronavirus.db")
    conn = sqlite3.connect(dbFile)
    cur = conn.cursor()
    
    comm = """CREATE TABLE IF NOT EXISTS pdf2ws0f 
                               (id       INTEGER PRIMARY KEY, 
                                filename TEXT, 
                                dt_snd   TEXT,
                                dt_rcv   TEXT,
                                status   INTEGER)
           """
    cur.execute(comm)


def record_upload(cur, filename, dt_snd, dt_rcv, status):
    if filename == "" or \
       dt_snd == "" or \
       dt_rcv == "" or \
       status == "":
       return False 
       
    cur.execute('INSERT INTO pdf2ws0f (filename, dt_snd, dt_rcv, status) VALUES (?,?)', (filename, dt_snd, dt_rcv, status))


root = tk.Tk()
root.withdraw()

try:
    setup_profile()  
except IOError as e:
    msg = "Aborted! creating user profile! Error: %s" % str(e)  
    logging.critical(msg)  
    tk.messagebox.showerror("Setup Error", msg) 
    exit()

try:
    setup_connection()  
except sqlite3.Error as e:
    msg = "Aborted! Error connecting to db! Error: %s" % str(e)  
    logging.critical(msg)  
    tk.messagebox.showerror("Setup Error", msg) 
    exit()

        


def upload_files():
    error = False
    os.chdir(entry_dir)
    files = tk.filedialog.askopenfilenames() if MULTI_FILES else filedialog.askopenfilename()
    
    for filename in files:
        file = os.path.basename(filename).lower()
        ext = ''.join(file.split(".")[-1])
        fbase = ''.join(file.split(".")[:-1])
        print(fbase, ext)
        if ext not in extensions:
            msg = "File %s cannot be sent due to wrong extension: %s!" % (file, ext)
            tk.messagebox.showerror("File Input Error", msg)
            #win32api.MessageBox(0, msg, "Critical Error", 0x00001000)         
            raise InputError(msg, msg)
            break
            
        
        files = {'file': (file, open(filename, 'rb'), 'application/pdf', {'Expires': '0'})}
        dt_snd = get_timestamp()
        r = requests.post(url, files=files)
        dt_rcv = get_timestamp()
        
        file_response = os.path.join(resp_dir, "%s%s" % (fbase, ".response"))
        print(file_response)
        
        fh_resp = open(file_response,"w")
        
        fh_resp.write("status code: %s" % str(r.status_code))
        fh_resp.write(str(r.headers))
        fh_resp.write(r.text)
        fh_resp.close()
    
        try:
            record_upload(cur, filename, dt_snd, dt_rcv, r.status_code)
        except sqlite3.Error as e:
            msg = "Aborted! Error recording entry in table for file %s! Error: %s" % (filename, str(e))  
            logging.critical(msg)  
            tk.messagebox.showerror("Setup Error", msg) 
            error = True
            break
        
            
        if r.status_code not in CODES_OK:
            msg = "Upload failed for File %s. status_code: %s!" % (file, r.status_code)
            tk.messagebox.showerror("Failed file upload", msg)
            #win32api.MessageBox(0, msg, "Critical Error", 0x00001000) 
            fh_out.write(str(r.headers))
            r.raise_for_status()
        else:
            msg = "Upload Successfull for File %s. status_code: %s!" % (file, r.status_code)
            tk.messagebox.showinfo("File upload successfull", msg)
            #win32api.MessageBox(0, msg, "Critical Error", 0x00001000) 
            fh_out.write(str(r.headers))
               
        fh_out.write(r.text)
    
    if error:
        conn.rollback()
    else:
        conn.commit()
    conn.close()
        
    fh_out.close()   



"""

def connect():

    con1 = sqlite3.connect("<path/database_name>")

    cur1 = con1.cursor()

    cur1.execute("CREATE TABLE IF NOT EXISTS table1(id INTEGER PRIMARY KEY, First TEXT, Surname TEXT)")

    con1.commit()

    con1.close()


def View():

    con1 = sqlite3.connect("<path/database_name>")

    cur1 = con1.cursor()

    cur1.execute("SELECT * FROM <table_name>")

    rows = cur1.fetchall()    

    for row in rows:

        print(row) 

        tree.insert("", tk.END, values=row)        

    con1.close()


# connect to the database

connect() 

root = tk.Tk()

tree = ttk.Treeview(root, column=("c1", "c2", "c3"), show='headings')

tree.column("#1", anchor=tk.CENTER)

tree.heading("#1", text="ID")

tree.column("#2", anchor=tk.CENTER)

tree.heading("#2", text="FNAME")

tree.column("#3", anchor=tk.CENTER)

tree.heading("#3", text="LNAME")

tree.pack()

button1 = tk.Button(text="Display data", command=View)

button1.pack(pady=10)

root.mainloop()

"""

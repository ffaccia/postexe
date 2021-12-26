
import requests
import sys
import os
from shutil import copyfile as sh_copy
import time
import base64
from datetime import datetime as dt

from tkinter import *
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
from tkcalendar import DateEntry

import sqlite3
import logging
from PIL import ImageTk, Image


FORMAT = "[ %(asctime)s, %(levelname)s] %(message)s" 

dbFile, conn, cur = None, None, None
CODES_OK, extensions, MULTI_FILES = None, None, None
entry_dir, abs_dir, resp_dir, save_dir, file_save = None, None, None, None, None
im_checked, im_unchecked, trv = None, None, None
checkall_btn, uncheckall_btn, upload_again_btn = None, None, None
url = None

SAVE_FILE_DB = True   
SAVE_FILE_DIR = True   
w_width = 740
w_height = 600

root = Tk()
root.title("PDF 2 WS")
root.geometry("%dx%d" % (w_width, w_height))
   
   
def setup_profile():    
    global extensions, CODES_OK, MULTI_FILES, im_checked, im_unchecked
    global resp_dir, save_dir, url
    
    url = 'https://httpbin.org/post'
    extensions = ['pdf']
    CODES_OK = [ v for k,v in requests.codes.__dict__.items() if k in ['OK','CREATED','ACCEPTED']]
    MULTI_FILES = True

    user_profile_dir = os.environ["USERPROFILE"]
    inner_dir = os.path.join('documents','sviluppo','python','postexe')
    entry_dir = os.path.join(user_profile_dir, inner_dir)
    if os.path.isdir(entry_dir) == False:
          msg = "Setup dir is missing: %s! " % (entry_dir)
          tk.messagebox.showerror("Setup Error", msg)
          #win32api.MessageBox(0, msg, "Critical Error", 0x00001000)         
          raise InputError(msg, msg)    
    
    abs_dir = os.path.join(entry_dir,entry_dir) 
    #print(abs_dir)
    os.chdir(abs_dir)
    
    funchecked = os.path.join(abs_dir, 'img', 'uncheckedn.png')
    fchecked = os.path.join(abs_dir, 'img', 'checkedn.png')
    im_checked = ImageTk.PhotoImage(Image.open(fchecked))
    im_unchecked = ImageTk.PhotoImage(Image.open(funchecked))
    
    print(funchecked)
    
    
    resp_dir = os.path.join(".","responses")
    if os.path.isdir(resp_dir) == False:
        os.mkdir(resp_dir, 755);
    
    save_dir = os.path.join(".","save")
    if os.path.isdir(save_dir) == False:
        os.mkdir(save_dir, 755);
    
    logName = os.path.join(".", "responses", "responses.log")
    logging.basicConfig(filename=logName, level=logging.DEBUG, format=FORMAT)




def setup_connection():
    global dbFile, conn, cur
    dbFile = os.path.join(".", "db", "coronavirus.db")
    conn = sqlite3.connect(dbFile)
    cur = conn.cursor()
    
    comm = """CREATE TABLE IF NOT EXISTS pdf2ws0f 
                               (id       INTEGER PRIMARY KEY, 
                                filename TEXT, 
                                size     INTEGER,
                                dt_snd   TEXT,
                                dt_rcv   TEXT,
                                status   INTEGER)
           """
    cur.execute(comm)

    comm = """CREATE TABLE IF NOT EXISTS pdf2file0f 
                               (id       INTEGER, 
                                filename VARCHAR(250), 
                                data     TEXT,
                                FOREIGN KEY(id) REFERENCES pdf2ws0f(id))
           """
    cur.execute(comm)



class DateEntry(DateEntry):
    def get_date(self):
        if not self.get():
            return None
        self._validate_date()
        return self.parse_date(self.get())       


"""
DateEntry(scroll_frame, width = 25, background = 'LightCyan3',
                                             foreground ='white',borderwidth=2)
"""



                    
def search_files(qfile, qdatada, qdataa, qstatus):
    qfile = qfile.get().lower()
    qdatada = qdatada.get()
    qdataa = qdataa.get() #.replace("/")
    qstatus = qstatus.get()
    
    print(qdatada)
    
    if qstatus != "":
        try:
            qstatus = int(qstatus)
        except ValueError:
            print("valueerror")
            qstatus=0

        
    query=f""" SELECT id, filename, size, dt_snd, dt_rcv, status FROM pdf2ws0f
              WHERE 1=1 
           """
    if qfile != None:
        query += f" AND lower(filename) LIKE '{qfile}%'"       

    if qstatus != "":
        query += f" AND status == '{qstatus}'"       

    if qdatada != None:
        query += f" AND substr(dt_snd, 1, 10) >= '{qdatada}'"       
        
    if qdataa != None:
        query += f" AND substr(dt_snd, 1, 10) <= '{qdataa}'"       

    query += """ ORDER BY id DESC"""
    
    print(query)
    cur.execute(query)
    rows=cur.fetchall()
    update(rows)


    
def clear():
    populate()




def setup_frames():    
    global trv, checkall_btn, uncheckall_btn, upload_again_btn

    wrapper1 = LabelFrame(root, text="Last Sent Pdf")
    wrapper2 = LabelFrame(root,  text="Upload Pdf")
    wrapper3 = LabelFrame(root,  text="Search for Pdf")
    wrapper4 = LabelFrame(root,  text="Remove entry")
    
    wrapper1.pack(fill=tk.X, expand="yes", padx=20, pady=10)
    wrapper2.pack(fill=tk.X, expand="yes", padx=20, pady=10)
    wrapper3.pack(fill=tk.X, expand="yes", padx=20, pady=10)
    wrapper4.pack(fill=tk.X, expand="yes", padx=20, pady=10)
    
    wrapper2.pack_propagate(0)
    wrapper3.pack_propagate(0)
    wrapper4.pack_propagate(0)
    
    qfile = StringVar()
    qstatus = StringVar()
    qid = IntVar()
        
    trv = ttk.Treeview(wrapper1, columns=(1,2,3,4,5,6))
    style=ttk.Style()
    style.configure('Treeview', rowheight=20)
    

    trv.tag_configure('checked', image=im_checked)
    trv.tag_configure('unchecked', image=im_unchecked)
    trv.tag_configure('gray', background='#cc6666')
    

    trv.column("#0", minwidth=40, width=40, stretch=0)
    trv.heading('#0', text="")

    trv.column("#1", minwidth=30, width=50, stretch=0, anchor=CENTER)
    trv.heading('#1', text="Id")

    trv.column("#2", minwidth=200, width=300, stretch=1)
    trv.heading('#2', text="File")

    trv.column("#3", minwidth=50, width=60, stretch=0, anchor=CENTER)
    trv.heading('#3', text="Size")

    trv.column("#4", minwidth=150, width=200, stretch=1)
    trv.heading('#4', text="dt_snd")

    trv.column("#5", minwidth=150, width=200, stretch=1)
    trv.heading('#5', text="dt_rcv")

    trv.column("#6", minwidth=50, width=60, stretch=1)
    trv.heading('#6', text="status")

       
    trv.pack_propagate(0)
    trv.pack()
    
    checkall_btn = Button(wrapper1, text="Check all", command=lambda:toggleCheck2("checked"))
    checkall_btn.pack(side=LEFT, anchor="w", padx=10)

    uncheckall_btn = Button(wrapper1, text="Uncheck all", command=lambda:toggleCheck2("unchecked"))
    uncheckall_btn.pack(side=LEFT, anchor="s", padx=10)

    upload_again_btn = Button(wrapper1, text="Upload again", command=upload_again)
    upload_again_btn.pack(side=LEFT, anchor="s", padx=10)

    populate()
    #toggleDisableButton()
    #upload_again_btn['state'] = "disabled"
    #checkall_btn['state'] = "disabled"
    #uncheckall_btn['state'] = "disabled"
    
    #trv.bind('<Double 1>', getrow)
    trv.bind('<Button 1>', toggleCheck)


    
    lblf = Label(wrapper2, text="Upload file...")
    lblf.grid(column=0, row=0, padx=5, pady=5)
    
    upload_btn = Button(wrapper2, text="Upload", command=upload_files)
    upload_btn.grid(column=1, row=0, padx=5, pady=5)
    
    
    lblf = Label(wrapper3, text="FileName...")
    lblf.grid(column=0, row=0, padx=5, pady=5)
    entf = Entry(wrapper3, textvariable=qfile)
    entf.grid(column=1, row=0, padx=5, pady=5)
    
    lbls = Label(wrapper3, text="Status.....")
    lbls.grid(column=2, row=0, padx=5, pady=5)
    ents = Entry(wrapper3, textvariable=qstatus)
    ents.grid(column=3, row=0, padx=5, pady=5)
    
    lbld = Label(wrapper3, text="Date from...")
    lbld.grid(column=0, row=1, padx=5, pady=5)
    entda = DateEntry(wrapper3,selectmode='day', date_pattern='Y-mm-d')
    entda.grid(column=1, row=1, padx=5, pady=5)
    
    lbld = Label(wrapper3, text="Date to.....")
    lbld.grid(column=2, row=1, padx=5, pady=5)
    enta = DateEntry(wrapper3,selectmode='day', date_pattern='Y-mm-d')
    enta.grid(column=3, row=1, padx=5, pady=5)
    
    
    search_btn = Button(wrapper3, text="Search Files", command=lambda:search_files(qfile, entda, enta, qstatus))
    search_btn.grid(column=0, row=2, padx=5, pady=5)
    clear_btn = Button(wrapper3, text="Clear", command=clear)
    clear_btn.grid(column=1, row=2, padx=5, pady=5)
    
    
    lbldel = Label(wrapper4, text="Id num.....")
    lbldel.grid(column=0, row=0, padx=5, pady=5)
    entdel = Entry(wrapper4, textvariable=qid)
    entdel.grid(column=1, row=0, padx=5, pady=5)
    
    del_btn = Button(wrapper4, text="Delete id", command=delete_id)
    del_btn.grid(column=2, row=0, padx=5, pady=5)
    
    exit_btn = Button(wrapper4, text="Exit App", command=root.destroy)
    #exit_btn.grid(padx=5, pady=5)
    exit_btn.pack(side=tk.RIGHT, padx=10)
    



def upload_files(here_file=None):
    global file_save
    error = False
    save = True
    #os.chdir(entry_dir)
    
    if here_file != None:
        print("herefile ", here_file)
        files = (os.path.join(save_dir, here_file),)
        save=False
    else:
        files = tk.filedialog.askopenfilenames() if MULTI_FILES else filedialog.askopenfilename()

    print("here file ")
    print(type(files))    

    for filename in files:
        abs_file = os.path.abspath(filename)
        size = os.stat(abs_file).st_size
        file = os.path.basename(filename).lower()
        ext = ''.join(file.split(".")[-1])
        fbase = ''.join(file.split(".")[:-1])
        print(fbase, ext)
        if ext not in extensions:
            msg = "File %s cannot be sent due to wrong extension: %s!" % (file, ext)
            tk.messagebox.showerror("File Input Error", msg)
            #win32api.MessageBox(0, msg, "Critical Error", 0x00001000)         
            logging.error(msg) 
            break
            
        
        files = {'file': (file, open(filename, 'rb'), 'application/pdf', {'Expires': '0'})}
        dt_snd = get_timestamp()
        r = requests.post(url, files=files)
        dt_rcv = get_timestamp()
        
        file_response = os.path.join(resp_dir, "%s%s" % (fbase, ".response"))

        file_save = os.path.join(save_dir, file)
        print("---")
        print(file_save)
        
        
        fh_resp = open(file_response,"w")
        
        fh_resp.write("status code: %s" % str(r.status_code))
        fh_resp.write(str(r.headers))
        fh_resp.write(r.text)
        fh_resp.close()
    
        id = record_upload(os.path.basename(filename.lower()), size, dt_snd, dt_rcv, r.status_code)
            
        if r.status_code not in CODES_OK:
            msg = "Upload failed for File %s. status_code: %s!" % (file, r.status_code)
            tk.messagebox.showerror("Failed file upload", msg)
            #win32api.MessageBox(0, msg, "Critical Error", 0x00001000) 
            logging.error(msg) 
            r.raise_for_status()
        else:
            msg = "Upload Successfull for File %s. status_code: %s!" % (file, r.status_code)
            tk.messagebox.showinfo("File upload successfull", msg)
            #win32api.MessageBox(0, msg, "Critical Error", 0x00001000) 
            logging.error(msg) 
            
            if SAVE_FILE_DIR and save:
                try:
                    sh_copy(abs_file, file_save)
                except:
                    msg = "Error saving file %s!" % os.path.basename(file)
                    logging.error(msg) 
                    tk.messagebox.showerror("Saving Error", msg)
            
            if SAVE_FILE_DB:
                try:
                    file_load(id, abs_file)
                except (e):
                    msg = "Error saving file %s into db. Error %s!" % (os.path.basename(file), str(e))
                    logging.error(msg) 
                    tk.messagebox.showerror("Saving Error", msg)
            
            
    

def get_timestamp():
    return dt.now().strftime(format='%Y-%m-%d %H:%M:%S.%f')[:-3]
    

    
def update(rows):
    trv.delete(*trv.get_children())
    for i, row in enumerate(rows):
        tags=['unchecked'] if i%2 == 1 else ['unchecked','gray']
        print("inserted ", tags)
        trv.insert('', 'end', values=row, tags=tags)
    
    toggleDisableButton()
   

def toggleCheck2(opt):
    for rowid in trv.get_children():
        trv.item(rowid, tags=opt)
    toggleDisableButton()

def countChecked(opt="checked"):
    cnt=0
    for rowid in trv.get_children():
        #iid = trv.index(item)
        if trv.item(rowid, "tags")[0] == opt:
            cnt +=1
            
    print("checked: ",cnt)
    return cnt


def toggleDisableButton():
    print(trv.get_children())
    tot = len(list(trv.get_children()))
    print("tot %s" % tot)
    if tot:
        if countChecked() == tot:
            checkall_btn['state'] = "disabled"
            uncheckall_btn['state'] = "active"
            upload_again_btn['state'] = "active"
        elif countChecked() == 0:    
            uncheckall_btn['state'] = "disabled"
            checkall_btn['state'] = "active"
            upload_again_btn['state'] = "disabled"
        elif countChecked():     
            uncheckall_btn['state'] = "active"
            checkall_btn['state'] = "active"
            upload_again_btn['state'] = "active"
    else:
        checkall_btn['state'] = "disabled"
        uncheckall_btn['state'] = "disabled"
        upload_again_btn['state'] = "disabled"
            
            
    
        
def toggleCheck(event):
    rowid = trv.identify_row(event.y)
    if rowid != None and rowid != "":
        tag = trv.item(rowid, "tags")[0]
        #print("--- ", rowid)
        #print(tag)
        #tags = list(trv.item(rowid, "tags"))
        #tags.remove(tag)
        #trv.item(rowid, tags=tags)
        if tag == "checked":
            trv.item(rowid, tags="unchecked")
        else:
            trv.item(rowid, tags="checked")
        toggleDisableButton()


def upload_again():
    for row in trv.get_children():
        item = trv.item(row)
        print("qwe---")
        print(item["tags"])
        print(item)
        print(item['values'][0])
        print(item['values'][1])    
        if item["tags"][0] == "checked":
            upload_files(item['values'][1])
        

def populate():
    query="SELECT id, filename, size, dt_snd, dt_rcv, status FROM pdf2ws0f ORDER BY id DESC"
    cur.execute(query)
    rows=cur.fetchall()
    update(rows)


def record_upload(filename, size, dt_snd, dt_rcv, status):
    if filename == "" or \
       size == "" or \
       dt_snd == "" or \
       dt_rcv == "" or \
       status == "":
       return 
    
    #print(filename, dt_snd, dt_rcv, status)
    try:       
        cur.execute('INSERT INTO pdf2ws0f (filename, size, dt_snd, dt_rcv, status) VALUES (?,?,?,?,?)', (filename, size, dt_snd, dt_rcv, status))
        rows=get_affected_rows()
        rowid = cur.lastrowid
        print(type(rows))
        print(rows)
        if rows != 1:   
            msg="Inserted rows: %d" % rows
            logging.error(msg) 
            messagebox.showerror("Insert Error", msg) 
            conn.rollback()
            logging.error("rollback occurred") 
            return True
    except sqlite3.Error as e:
        conn.rollback()
        messagebox.showerror("Insert Error", str(e))
    else:
        conn.commit()
    finally:
        populate()
        return rowid
        
    
def file_load(id, file):
    
    error = False
    print("foreing")
    print(id, file)
    basename = os.path.basename(file)
    text_data = base64.b64encode(open(file,"rb").read())
    #print(text_data)
    
    try:       
        #cur.execute("INSERT INTO pdf2file0f (id, filename, data) VALUES ({0},{1},{2})".format(id, basename, text_data))
        sql="INSERT INTO pdf2file0f (id, filename, data) VALUES (?,?,?) "
        cur.execute(sql, (id, basename, text_data))
    except sqlite3.Error as e:
        print("error sql")
        error=True
        msg="Error inserting pdf2file0f filename %s, error: %s" % (basefile, str(e))
        messagebox.showerror(msg)
        print(msg)
        logging.error(msg) 
        conn.rollback()

    except e:
        conn.rollback()
        print("error sql2")
        error=True
        msg="Error inserting pdf2file0f filename %s, error: %s" % (basefile, str(e))
        logging.error(msg) 
        messagebox.showerror(msg)
    else:
        print("ok sql")
        error=False
        conn.commit()
    finally:
        print("return sql2")
        return error
        
    
    
def getrow():
    #rowid = trv.identify_row(event.y)
    item = trv.item(trv.focus())
    logging.info("%s %s %s %s %s" % (item['values'][0], item['values'][1], item['values'][2], item['values'][3], item['values'][4]))
     
    #t1.set(item['values'][0]) t2.set(item['values'][1]) t3.set(item['values'][2]) t4.set(item['values'][3]) t5.set(item['values'][4])
    
    
def get_affected_rows():
    query = "SELECT changes()"
    cur.execute(query)
    return cur.fetchone()[0]
       

def delete_id():
    id = t1.get()       

    if messagebox.askyesno("Confirm please", "Are you sure you want to delete customer %d ?" % id) == False:
        return False

    try:
        query=""" DELETE FROM pdf2ws0f WHERE id = %d """ % id
        cur.execute(query)      
        
        if get_affected_rows() != 1:   
            msg="Deleted rows: %d" % get_affected_rows()
            logging.error(msg)
            messagebox.showwarning("Delete Error", msg) 
            conn.rollback()
            logging.error("rollback occurred") 
            return True
    except sqlite3.Error as e:
        conn.rollback()
        logging.error("rollback occurred") 
        messagebox.showerror("Insert Error", str(e))
    else:
        conn.commit()
    
    return True            
 

 



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

 
 
setup_frames()


root.mainloop()






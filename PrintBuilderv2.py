##========================================
##Created by : Collin Preston
##Created date: 3/11/2021
##========================================


"""
    Print Builder +Plus Ultra

    Allows techs to quickly create print script files for devices. Techs
    can also use the built in search function to find devices by name, ip address,
    or location. After using the search feature Techs can use the selection box to
    select the device and send the selection to the entry fields on the GUI.

"""

##==========================
## Imports
##==========================
import tkinter as tk
from tkinter import *
import tkinter.font as tkFont
from tkinter import messagebox
from tkinter import filedialog 
import os, time
import pandas as pd
import xlwings as xw
import win32com.client as win32

#Fuction determines which radio button is selected, this value will be
#used by the submitValues function to create correct file type.
def sel(var):
    selection = var.get()
    return selection
                       
def img_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


#List of printer queues is store in an excel file.
#This function updates the file by opening and refreshing the document. Document will close automatically after update.
def updateExcel():
    global quelabel
    xlapp = win32.DispatchEx('Excel.Application')
    xlapp.DisplayAlerts = False
    xlapp.Visible = True

    xlbook = xlapp.Workbooks.Open(r'\\phcnas01\main\_Field Support\_Utilities\Print Queues\Queues.xlsx')

    # Refresh all pivot tables
    xlbook.RefreshAll()

    xlbook.Save()
    xlbook.Close()
    xlapp.Quit()
    quelabel.config(text = str("Updated: " + time.ctime(os.path.getmtime(quefile))))
    
    # Make sure Excel completely closes
    del xlbook
    del xlapp
    



#GUI Window


window = Tk()
window.geometry("600x600") 
window.title("Print Builder PlusUltra+")
window.minsize(width=600, height=600)
window.maxsize(width=600, height=600)
fontstyle1 = tkFont.Font(family="Lucida Grande", size=20)
fontstyle2 = tkFont.Font(family="Lucida Grande", size=10)
fontstyle3 = tkFont.Font(family="Lucida Grande", size=8)
fontstyle4 = tkFont.Font(family="Showcard Gothic", size=20)
fontstyle5 = tkFont.Font(family="Lucida Grande", size=7)
bg = PhotoImage( file = img_resource_path("backgf.gif"))
backback =Label(window, image=bg)
backback.place(x = 0,y = 0)
window.iconbitmap(default=img_resource_path('piedmonticon.ico'))

# what you see on the GUI
# Header
header = Label(window,font=fontstyle1,  text="Print Builder +Plus Ultra" )
header.grid(row=2, column=2,columnspan=2)
l1 = Label(window,font=fontstyle2, text="Computer Names (Separate by comma)")
l1.grid(row=3,column=1,sticky="w",padx=2,columnspan=2)
l1.focus()
e1 = Entry(window,show=None, bd =3, width=50)
e1.grid(row=4,column=1,sticky="w",padx=2,columnspan=2)
s2 = Entry(window,show=None, bd =3, width=5,)
s2.grid(row=6,column=1,sticky="w",padx=2)
l2 = Label(window,font=fontstyle2, text="Default Printer")
l2.grid(row=5,column=2,sticky="w",padx=2,columnspan=2)
e2 = Entry(window,show=None, bd =3, width=30,)
e2.grid(row=6,column=2,sticky="w",padx=2)
s3 = Entry(window,show=None, bd =3, width=5,)
s3.grid(row=8,column=1,sticky="w",padx=2)
l3 = Label(window,font=fontstyle2, text="Addtional Printers",)
l3.grid(row=7,column=2,sticky="w",padx=2)
e3 = Entry(window,show=None, bd =3, width=30,)
e3.grid(row=8,column=2,sticky="w",padx=2)
s4 = Entry(window,show=None, bd =3, width=5,)
s4.grid(row=9,column=1,sticky="w",padx=2)
e4 = Entry(window,show=None, bd =3, width=30,)
e4.grid(row=9,column=2,sticky="w",padx=2)
s5 = Entry(window,show=None, bd =3, width=5,)
s5.grid(row=10,column=1,sticky="w",padx=2)
e5 = Entry(window,show=None, bd =3, width=30,)
e5.grid(row=10,column=2,sticky="w",padx=2)
s6 = Entry(window,show=None, bd =3, width=5,)
s6.grid(row=11,column=1,sticky="w",padx=2)
e6 = Entry(window,show=None, bd =3, width=30,)
e6.grid(row=11,column=2,sticky="w",padx=2)
l7 = Label(window,font=fontstyle2, text="Search Name or IP Address")
l7.place(x=70,y=320)
e7 = Entry(window,show=None, bd =3, width=45,)
e7.place(x=245,y=320)
window.grid_columnconfigure(2, weight=1)
checkvar = tk.IntVar()
c1 = tk.Checkbutton(window, text='Overwrite',variable=checkvar, onvalue=1, offvalue=0, command=None)
c1.place(x=125,y=270)
servernum = Label(window,font=fontstyle3, text="Server\n #",)
servernum.grid(row=5,column=1,sticky="w",padx=2)
var = IntVar()
R1 = Radiobutton(window, text="Desktop/Laptop", variable=var, value=1,command=lambda: sel(var))
R1.place(x=300 ,y=123)
R2 = Radiobutton(window, text="Zero Client", variable=var, value=2,command=lambda: sel(var))
R2.place(x=300 ,y=143)

# button to open the piednth1\clients path
infoboxcheck = 1
infobtn = Button(window, width=6,height=1,font=fontstyle3,text="?",bg="firebrick4",fg="white",activebackground="white",activeforeground="firebrick4", command = lambda: about())
infobtn.place(x=540,y=10) 

updateexcel = Button(window, width=8,height=2,font=fontstyle3,text="Update\nQueue List",bg="red",fg="yellow",activebackground="yellow",activeforeground="red", command = lambda: updateExcel())
updateexcel.place(x=530,y=540)
#Printer Search Listbox and buttons
Lb1 = Listbox(window,height=10,width=75,selectbackground="sienna3",bg="white",bd=3)
Lb1.place(x=70, y=350)
searchbtn = Button(window, width=6,height=1,font=fontstyle4,text="SEARCH",bg="sienna1",fg="black",activebackground="black",activeforeground="sienna1", command = lambda: populatebox())
searchbtn.place(x=235, y=522) 

#Info Button, Version info, Contact information
def about3():
    global infoboxcheck
    infoboxcheck = 1

def about2(popup):
    popup.destroy()


def about():
    global infoboxcheck
    if infoboxcheck == 1:
        infoboxcheck = 0
        root_x = window.winfo_rootx()
        root_y = window.winfo_rooty()
        popup = Tk()
        win_x = root_x + 150
        win_y = root_y + 200
        popup.geometry(f'+{win_x}+{win_y}')
        popup.wm_title("Program Info")
        
        popup.configure(background='sky blue')
        infofont1 = tkFont.Font(family="Lucida Grande", size=20, weight=tkFont.BOLD)        
        infolabel = Label(popup, font = infofont1, text="Print Builder +Plus Ultra\nVersion 2.1\nCopyright 2021 C.Preston",bg="sky blue",fg="black")
        infolabel.pack(side="top", fill="x", pady=10)
        infolabel2 = Label(popup, text="Program For Use Only By Those With Permission\nSend Questions/Comments/Bug Reports to\nCollin.Preston1@piedmont.org",bg="sky blue", font=fontstyle4)
        infolabel2.pack( fill="x", pady=10)
        popup.attributes("-topmost", True)
        popup.overrideredirect(True)
        B1 = Button(popup, text="Close",bg="azure", command = lambda: [about2(popup),about3()])
        B1.pack(side="bottom")
    else:
        popup.destroy()

#List of printer queues is store in an excel file.
#Checks date last motified and updates visually on the GUI.
quefile = r'\\phcnas01\main\_Field Support\_Utilities\Print Queues\Queues.xlsx'
queprint = StringVar()
queprinta = time.ctime(os.path.getmtime(quefile))
queprintb = str("Updated: " + time.ctime(os.path.getmtime(quefile)))
quelabel = Label(window, text=str("Last Updated: " + time.ctime(os.path.getmtime(quefile))))
quelabel.place(x=380,y=580)


    
#ListBox Selection Field Buttons

def selectionbuttons():
    select1 = 1
    select2 = 2
    select3 = 3
    select4 = 4
    select5 = 5
    f1btn = Button(window, text=" ",bg="sienna1",activebackground="red",font=fontstyle5,width=2, command = lambda: searchValues(select1,0,0,0,0))
    f1btn.place(x=235, y=123)
    f2btn = Button(window, text=" ",bg="sienna1",activebackground="red",font=fontstyle5,width=2, command = lambda: searchValues(0,select2,0,0,0))
    f2btn.place(x=235, y=165)
    f3btn = Button(window, text=" ",bg="sienna1",activebackground="red",font=fontstyle5,width=2, command = lambda: searchValues(0,0,select3,0,0))
    f3btn.place(x=235, y=189)
    f4btn = Button(window, text=" ",bg="sienna1",activebackground="red",font=fontstyle5,width=2, command = lambda: searchValues(0,0,0,select4,0))
    f4btn.place(x=235, y=214)
    f5btn = Button(window, text=" ",bg="sienna1",activebackground="red",font=fontstyle5,width=2, command = lambda: searchValues(0,0,0,0,select5))
    f5btn.place(x=235, y=238)


#Search function for the listbox ( search(), populatebox(), and searchValues()
#List of printer queues is store in an excel file. Search function uses this file to populate
#the listbox.
def search():
    nope = e7.get()
    data = pd.read_excel (r'\\phcnas01\main\_Field Support\_Utilities\Print Queues\Queues.xlsx')
    df = pd.DataFrame(data, columns= ['Server Name', 'Printer Name', 'Port Name','Comment'])
    mul = df[df.apply(lambda row: row.astype(str).str.contains(nope,case=False,na=False).any(),axis=1)]
    #result = df.to_records(index=False).tolist()
    result = mul.values.tolist()
    #searched = filter(lambda x: x[0].startswith(nope), result)
    return result
    #print(searched)
def populatebox():
    Lb1.delete(0, END)
    if e7.get() != "":
        for i in search():
            Lb1.insert("end", i)
    else:
        failure = messagebox.showinfo("Error","Field Cannot Be Left Empty")
    return  
e7.bind('<Return>',lambda event=None: searchbtn.invoke())

#When selecting items in the list box, the selection can be sent entry windows using the field button
def searchValues(select1,select2,select3,select4,select5):
    it1 = Lb1.curselection()
    if(len(it1) == 0):
        pass
    else:
        items = Lb1.get(Lb1.curselection())
        item1=[0]
        item2=[1]
        sname = []
        pname = []
        for index in item1:
            sname.append(items[index])
        for index in item2:
            pname.append(items[index])        
        finalname = str("{}\{}".format(sname[0][5:],pname[0])) 
        if select1 == 1:
            e2.delete(0, END)
            e2.insert(0,pname[0])
            s2.delete(0, END)
            s2.insert(0,sname[0][5:])
        elif select2 == 2:
            e3.delete(0, END)
            e3.insert(0,pname[0])
            s3.delete(0, END)
            s3.insert(0,sname[0][5:])
        elif select3 == 3:
            e4.delete(0, END)
            e4.insert(0,pname[0])
            s4.delete(0, END)
            s4.insert(0,sname[0][5:])
        elif select4 == 4:
            e5.delete(0, END)
            e5.insert(0,pname[0])
            s5.delete(0, END)
            s5.insert(0,sname[0][5:])
        elif select5 == 5:
            e6.delete(0, END)
            e6.insert(0,pname[0])
            s6.delete(0, END)
            s6.insert(0,sname[0][5:])
    return




#Fuction used to create a file based upon user entry, radio button selection, and computer names.
#Will create a seperate file for each computer name list and store it in its directory.
def submitValues():

    #loop that cycles all the names in the entry field for Computer Names
    eall = e1.get()
    elist = eall.split(",")
    for i in range(len(elist)):
        fname = elist[i]

        #desktop and laptop content        
        if sel(var) == 1:            
            filename = r"\\piednth1\clients\{}\printers.txt".format(fname)
            if e2.get() != "":
                p2 = "\nD \\\\phcmp{}\{}\n".format(s2.get(),e2.get())
            else:
                p2 =""
            if e3.get() != "":
                p3 = "\n\\\\phcmp{}\{}\n".format(s3.get(),e3.get())
            else:
                p3 =""
            if e4.get() != "":
                p4 = "\\\\phcmp{}\{}\n".format(s4.get(),e4.get())
            else:
                p4 =""
            if e5.get() != "":
                p5 = "\\\\phcmp{}\{}\n".format(s5.get(),e5.get())
            else:
                p5 =""
            if e6.get() != "":
                p6 = "\\\\phcmp{}\{}\n".format(s6.get(),e6.get())
            else:
                p6 =""  
            content = str('{}{}{}{}{}'.format(p3,p4,p5,p6,p2))



        #zero client content
        elif sel(var) == 2:
            filename = r"\\piednth1\clients\{}\setprint.cmd".format(fname)
            if e2.get() != "":
                p2 = "\nRunDll32.EXE printui.dll,PrintUIEntry /in /n \\\\phcmp{}\{}\n".format(s2.get(),e2.get()) + "REM * Set Default Printer *\n" +"RunDll32.EXE printui.dll,PrintUIEntry /y /n \\\\phcmp{}\{}\n".format(s2.get(),e2.get())
            else:
                p2 =""
            if e3.get() != "":
                p3 = "RunDll32.EXE printui.dll,PrintUIEntry /in /n \\\\phcmp{}\{}\n".format(s3.get(),e3.get())
            else:
                p3 =""
            if e4.get() != "":
                p4 = "RunDll32.EXE printui.dll,PrintUIEntry /in /n \\\\phcmp{}\{}\n".format(s4.get(),e4.get())
            else:
                p4 =""
            if e5.get() != "":
                p5 = "RunDll32.EXE printui.dll,PrintUIEntry /in /n \\\\phcmp{}\{}\n".format(s5.get(),e5.get())
            else:
                p5 =""
            if e6.get() != "":
                p6 = "RunDll32.EXE printui.dll,PrintUIEntry /in /n \\\\phcmp{}\{}\n".format(s6.get(),e6.get())
            else:
                p6 =""  
            content = str('\n'
                          'con2prt /f\n'
                          '{}{}{}{}{}'.format(p3,p4,p5,p6,p2)
                          )
        #if no radio button is selected you will get error
        elif sel(var) == 0:
            failure = messagebox.showinfo("Error","Please Select Device Type Desktop/Laptop or Zero Client")
            return
        if e1.get() =="":
            failure = messagebox.showinfo("Error","Please Enter Computer Name")
            return
        if checkvar.get() == 1:
            if e2.get() == "":
                failure = messagebox.showinfo("Error","Please Enter Default Printer")
                return

        output(filename, content, elist)
    success = messagebox.showinfo("Success","Your File Has Been Successfully Created for the following devices:{}".format(elist))
                
#Function that creates a file from the content and saves to drive storage.
def output(filename, content, elist):
        if checkvar.get() == 0:
            os.makedirs(os.path.dirname(filename), exist_ok=True)
            f = open(filename, "a")
            print(content, file=f)
            f.close()
        else:
            os.makedirs(os.path.dirname(filename), exist_ok=True)
            f = open(filename, "w")
            print(content, file=f)
            f.close()
   


#Resets all entry fields
def clearall():
        e1.delete(0, END)
        e2.delete(0, END)
        e3.delete(0, END)
        e4.delete(0, END)
        e5.delete(0, END)
        e6.delete(0, END)
        s2.delete(0, END)
        s3.delete(0, END)
        s4.delete(0, END)
        s5.delete(0, END)
        s6.delete(0, END)
        e7.delete(0, END)
        Lb1.delete(0, END)
        var.set(0)
        checkvar.set(0)
        

#Submit and Reset buttons visually on GUI.
submit = tk.Button(window, text="Submit",command=lambda: submitValues()) 
submit.place(x=220,y=270)
clear = tk.Button(window, text="Reset",command=lambda: clearall()) 
clear.place(x=290,y=270)



selectionbuttons()
window.mainloop()

import win32com.client
import pandas as pd
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300)
canvas1.pack()

def Mailer(CC, attachment, list, word_list):

    email_list = pd.read_excel(list)
    emails = email_list['EMAIL']

    keys = word_list.split()

    
    result = ""
    for i, key in enumerate(keys):
        if key in email_list:
            result += email_list[key]
            if i < len(keys) - 1:  # Check if it's not the last key
                result += "<br><br>"
            

    subjects = email_list['SUBJECT']

    for i in range(len(emails)):

        ol=win32com.client.Dispatch("outlook.application")
        olmailitem=0x0 
        newmail=ol.CreateItem(olmailitem)
    

        email = emails[i]
        text = result[i]
        subject = subjects[i]

        newmail.Subject= subject
        newmail.To= email
        newmail.CC= CC
        newmail.Attachments.Add(attachment)
        newmail.GetInspector 

        index = newmail.HTMLbody.find('>', newmail.HTMLbody.find('<body')) 
        newmail.HTMLbody = newmail.HTMLbody[:index + 1] + text + newmail.HTMLbody[index + 1:] 

        newmail.Display()

        print ("opened " + str(i))

def browse_file(entry_widget):
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File", filetypes=[("All Files", "*.*")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)

label_kopie = tk.Label(root, text="V kopii:")
label_kopie.pack()

kopie_entry = tk.Entry(root)
kopie_entry.pack()

label_sloupce = tk.Label(root, text="Textové sloupce oddělené mezerou")
label_sloupce.pack()

sloupce_entry = tk.Entry(root)
sloupce_entry.pack()

label_attachement = tk.Label(root, text="Cesta přílohy:")
label_attachement.pack()

attachement_entry = tk.Entry(root)
attachement_entry.pack()

attachement_button = tk.Button(root, text="Browse", command=lambda: browse_file(attachement_entry))
attachement_button.pack()

label_list = tk.Label(root, text="Cesta seznamu:")
label_list.pack()

list_entry = tk.Entry(root)
list_entry.pack()

list_button = tk.Button(root, text="Browse", command=lambda: browse_file(list_entry))
list_button.pack()

button1 = tk.Button(text="Zobrazit maily", command= lambda: Mailer(kopie_entry.get(),attachement_entry.get(),list_entry.get(), sloupce_entry.get() ), bg="brown", fg="white")
canvas1.create_window(150, 150, window=button1)

root.mainloop()
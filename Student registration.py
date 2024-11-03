import tkinter
from tkinter import messagebox, ttk
import openpyxl
from openpyxl import load_workbook
root = tkinter.Tk()
root.title('Signup')
root.configure(bg="#a8d5e3")
root.geometry("300x350")
file_path = r"C:\Web Development Projects\Python\Signup data.xlsx"
A = openpyxl.load_workbook(file_path)
B = A["Signup data"]

def OnClickSubmit():
    name = name_box.get()
    email = email_box.get()
    mobno = mobno_box.get()
    branch = branch_dropdown.get()
    agree = agree_value.get()
    if name and email and mobno and branch and agree==1:
        B.append([name, email, mobno, branch, 'Yes'])
        A.save(file_path)
        messagebox.showinfo("Status", "Data Submitted")
    else:
        messagebox.showwarning('Warning', 'Field cannot be empty')

name_label = tkinter.Label(root, text="Name")
name_label.pack(anchor="w", padx=10, pady=5)
name_box = tkinter.Entry(root)
name_box.pack(anchor="w", padx=10)

email_label = tkinter.Label(root, text="Email Id")
email_label.pack(anchor="w", padx=10, pady=5)
email_box = tkinter.Entry(root)
email_box.pack(anchor="w", padx=10)

mobno_label = tkinter.Label(root, text="Mobile Number")
mobno_label.pack(anchor="w", padx=10, pady=5)
mobno_box = tkinter.Entry(root)
mobno_box.pack(anchor="w", padx=10)

choices = ['Computer Science', "IT", "Mechanical", "Electrical", "Ai", "DSA"]
branch_label = tkinter.Label(root, text="Select Branch", anchor='w')
branch_label.pack(anchor='w', padx=10)
branch_dropdown = ttk.Combobox(root, values=choices)
branch_dropdown.pack(anchor='w', padx=10)

agree_value = tkinter.IntVar()
agree_checkbox = ttk.Checkbutton(root, text="Terms & Conditions", variable=agree_value)
agree_checkbox.pack(anchor='w', padx='10', pady=5)

submit_button = tkinter.Button(root, text="Submit", command=OnClickSubmit)
submit_button.pack(anchor='w', padx=10, pady=10)

root.mainloop()
from tkinter import *
from tkinter import messagebox, ttk, filedialog, simpledialog
import pandas as pd
from tkcalendar import DateEntry
import datetime
from docxtpl import DocxTemplate
#from PIL import ImageTk,Image
import sqlite3 
from twilio.rest import Client

def open_payments():
    payments()

def open_register():
    register()

root = Tk()
root.title('DemoBuild')

mydb = sqlite3.connect('students.db')
cursor = mydb.cursor()

# Create tables
cursor.execute("""CREATE TABLE IF NOT EXISTS students (
                student_id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_name TEXT NOT NULL,
                contact INTEGER NOT NULL,
                email TEXT NOT NULL,
                class TEXT NOT NULL,
                teacher_name TEXT NOT NULL,
                subjects TEXT NOT NULL,
                timing TEXT NOT NULL,
                branch TEXT NOT NULL,
                fees INTEGER NOT NULL,
                last_paid_on TEXT,
                Fee_status TEXT
)""")
cursor.execute("""CREATE TABLE IF NOT EXISTS teachers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL
)""")
cursor.execute("""CREATE TABLE IF NOT EXISTS branches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL
)""")
cursor.execute("""CREATE TABLE IF NOT EXISTS Classes (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               name TEXT NOT NULL
)""")
cursor.execute("""CREATE TABLE IF NOT EXISTS subjects (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               name TEXT NOT NULL
)""")
cursor.execute("""CREATE TABLE IF NOT EXISTS payments (
               payment_id INTEGER PRIMARY KEY AUTOINCREMENT,
               student_id INTEGER NOT NULL,
               student_name TEXT NOT NULL,
               amount_paid INTEGER NOT NULL,
               payment_date TEXT NOT NULL,
               next_due_date TEXT NOT NULL,
               payment_method TEXT NOT NULL
)""")
mydb.commit()

def register():
    # Create a new window for app x
    register_window = Toplevel()
    
    def clear_fields():
        student_name_box.delete(0, END)
        contact_box.delete(0, END)
        email_box.delete(0, END)
        class_box.set('')
        timing_box.delete(0, END)
        branch_box.set('')
        fees_box.delete(0, END)
        for i in teacher_listbox.curselection():
            teacher_listbox.selection_clear(i)
        for i in subject_listbox.curselection():
            subject_listbox.selection_clear(i)

    def get_subjects():
        cursor.execute("SELECT name FROM subjects")
        subjects = [row[0] for row in cursor.fetchall()]
        return subjects

    def add_subject():
        new_subject_name = simpledialog.askstring("Input", "Enter new subject's name:")
        if new_subject_name:
            cursor.execute("INSERT INTO subjects (name) VALUES (?)", (new_subject_name,))
            mydb.commit()
            update_subject_list()
            messagebox.showinfo("Success", f"subject '{new_subject_name}' added successfully.")

    def update_subject_list():
        subjects = get_subjects()
        subject_listbox.delete(0, END)
        for subject in subjects:
            subject_listbox.insert(END, subject)
        subject_listbox.insert(END, "Add subject (+)")

    def on_subject_selected(event):
        selected_subject = subject_listbox.get(ACTIVE)
        if selected_subject == "Add subject (+)":
            add_subject()

    def get_teachers():
        cursor.execute("SELECT name FROM teachers")
        teachers = [row[0] for row in cursor.fetchall()]
        return teachers

    def add_teacher():
        new_teacher_name = simpledialog.askstring("Input", "Enter new teacher's name:")
        if new_teacher_name:
            cursor.execute("INSERT INTO teachers (name) VALUES (?)", (new_teacher_name,))
            mydb.commit()
            update_teacher_list()
            messagebox.showinfo("Success", f"Teacher '{new_teacher_name}' added successfully.")

    def update_teacher_list():
        teachers = get_teachers()
        teacher_listbox.delete(0, END)
        for teacher in teachers:
            teacher_listbox.insert(END, teacher)
        teacher_listbox.insert(END, "Add Teacher (+)")

    def on_teacher_selected(event):
        selected_teacher = teacher_listbox.get(ACTIVE)
        if selected_teacher == "Add Teacher (+)":
            add_teacher()  

    def get_branches():
        cursor.execute("SELECT name FROM branches")
        branches = [row[0] for row in cursor.fetchall()]
        return branches

    def add_branch():
        new_branch_box = simpledialog.askstring("Input", "Enter new branch's name:")
        if new_branch_box:
            cursor.execute("INSERT INTO branches (name) VALUES (?)", (new_branch_box,))
            mydb.commit()
            update_branch_list()
            messagebox.showinfo("Success", f"Branch '{new_branch_box}' added successfully.")

    def update_branch_list():
        branches = get_branches()
        branch_box['values'] = branches + ["Add Branch (+)"]
        branch_box.set('')

    def on_branch_selected(event):
        selected_branch = branch_box.get()
        if selected_branch == "Add Branch (+)":
            add_branch() 

    def get_classes():
        cursor.execute("SELECT name FROM classes")
        classes = [row[0] for row in cursor.fetchall()]
        return classes

    def add_class():
        new_class_box = simpledialog.askstring("Input", "Enter new Class:")
        if new_class_box:
            cursor.execute("INSERT INTO classes (name) VALUES (?)", (new_class_box,))
            mydb.commit()
            update_class_list()
            messagebox.showinfo("Success", f"Class '{new_class_box}' added successfully.")

    def update_class_list():
        classes = get_classes()
        class_box['values'] = classes + ["Add Class (+)"]
        class_box.set('')

    def on_class_selected(event):
        selected_class = class_box.get()
        if selected_class == "Add Class (+)":
            add_class()         

    def add_student():
        selected_teachers = [teacher_listbox.get(i) for i in teacher_listbox.curselection()]
        selected_teachers_str = ', '.join(selected_teachers)
        selected_subjects = [subject_listbox.get(i) for i in subject_listbox.curselection()]
        selected_subjects_str = ', '.join(selected_subjects)
        cursor.execute('INSERT INTO students (student_name, contact, email, class, teacher_name, subjects, timing, branch, fees) VALUES (?,?,?,?,?,?,?,?,?)',
                       (student_name_box.get(), contact_box.get(), email_box.get(), class_box.get(), selected_teachers_str, selected_subjects_str, timing_box.get(), branch_box.get(), fees_box.get()))
        id_of_registered_student = cursor.lastrowid
        messagebox.showinfo("Success", f"Student added successfully. Student ID: {id_of_registered_student}")
        mydb.commit()
        clear_fields()

    # Creating entry boxes for register_window
    student_name_box = Entry(register_window)
    student_name_box.grid(row=1, column=1)

    contact_box = Entry(register_window)
    contact_box.grid(row=2, column=1, pady=5)

    email_box = Entry(register_window)
    email_box.grid(row=3, column=1, pady=5)

    classes = get_classes()
    class_box = ttk.Combobox(register_window, values=classes + ["Add Class (+)"], state='readonly')
    class_box.grid(row=4, column=1, pady=5)
    class_box.bind("<<ComboboxSelected>>", on_class_selected) 

    teachers = get_teachers()
    teacher_listbox = Listbox(register_window, selectmode=MULTIPLE, exportselection=0)
    for teacher in teachers:
        teacher_listbox.insert(END, teacher)
    teacher_listbox.insert(END, "Add Teacher (+)")
    teacher_listbox.grid(row=5, column=1, pady=5)
    teacher_listbox.bind("<<ListboxSelect>>", on_teacher_selected)

    subjects = get_subjects()
    subject_listbox = Listbox(register_window, selectmode=MULTIPLE, exportselection=0)
    for subject in subjects:
        subject_listbox.insert(END, subject)
    subject_listbox.insert(END, "Add subject (+)")
    subject_listbox.grid(row=6, column=1, pady=5)
    subject_listbox.bind("<<ListboxSelect>>", on_subject_selected)

    timing_box = Entry(register_window)
    timing_box.grid(row=7, column=1, pady=5)

    branches = get_branches()
    branch_box = ttk.Combobox(register_window, values=branches + ["Add Branch (+)"], state='readonly')
    branch_box.grid(row=8, column=1, pady=5)
    branch_box.bind("<<ComboboxSelected>>", on_branch_selected)

    fees_box = Entry(register_window)
    fees_box.grid(row=9, column=1, pady=5)

    # Creating labels
    student_name_label = Label(register_window, text="Student Name").grid(row=1, column=0, sticky=W, padx=10)
    contact_label = Label(register_window, text="Phone no.").grid(row=2, column=0, sticky=W, padx=10)
    email_label = Label(register_window, text="E-mail").grid(row=3, column=0, sticky=W, padx=10)
    branch_label = Label(register_window, text="Class").grid(row=4, column=0, sticky=W, padx=10)
    teacher_label = Label(register_window, text="Teacher Name").grid(row=5, column=0, sticky=W, padx=10)
    subject_labe = Label(register_window, text="Subjects").grid(row=6, column=0, sticky=W, padx=10)
    timing_label = Label(register_window, text="Batch Timing").grid(row=7, column=0, sticky=W, padx=10)
    branch_label = Label(register_window, text="Branch").grid(row=8, column=0, sticky=W, padx=10)
    fees_label = Label(register_window, text="Monthly Fee").grid(row=9, column=0, sticky=W, padx=10)

    # Create add students button
    add_student_button = Button(register_window, text="Add Student", command=add_student)
    add_student_button.grid(row=10, column=0, padx=10, pady=10)

    # Creating clear fields button
    clear_fields_button = Button(register_window, text="Clear Fields", command=clear_fields)
    clear_fields_button.grid(row=10, column=1)

    register_window.mainloop()  

def payments():
    payments_window = Toplevel()  

    def record_payment():
        cursor.execute('INSERT INTO payments (student_id, student_name, amount_paid, payment_date, next_due_date, payment_method) VALUES (?,?,?,?,?,?)',
                        (student_id_box.get(),student_name_from_id.cget("text"),amount_paid_box.get(),payment_date_box.get(),next_due_date_box.get(),payment_method_box.get()))
        
        #updating last_paid on students table
        cursor.execute("""UPDATE students SET
                       last_paid_on = :last_paid

                       WHERE oid = :oid""",
                       {'last_paid': payment_date_box.get(),
                        'oid': student_id_box.get()
                       }
        )
        paymentID = cursor.lastrowid

        if next_due_date_box.get():
            # Compare next_due_date with current date to determine fee status
            next_due_date = datetime.datetime.strptime(next_due_date_box.get(), '%d-%m-%Y')
            if next_due_date <= datetime.datetime.today():
                fee_status = 'Due'
            else:
                fee_status = 'Paid'
            
            # Update fee_status in students table
            cursor.execute("""UPDATE students SET
                            fee_status = :fee_status
                            WHERE oid = :oid""",
                            {'fee_status': fee_status,
                                'oid': student_id_box.get()
                            }
                        )
        invoice = DocxTemplate("invoice_template.docx")
        cursor.execute("""SELECT subjects, class FROM students WHERE oid = :oid""", {
            'oid': student_id_box.get()
        })
        Records_for_invoice = cursor.fetchall()
        for record_for_invoice in Records_for_invoice:
            class_invoice = record_for_invoice[1]
            subjects_invoice = record_for_invoice[0]

        invoice.render({"student_name":student_name_from_id.cget("text"), "fee_amount":amount_paid_box.get(),"paid_on":payment_date_box.get(),"next_due_date":next_due_date_box.get(),"receipt_id":paymentID, "gstin":"09ADPFS7434Q1ZY", "subjects":subjects_invoice, "class":class_invoice})
        invoice_name = simpledialog.askstring("Input", "Save Invoice as:")
        invoice.save(invoice_name + ".docx")

        mydb.commit()
        messagebox.showinfo("Success", f"Payment recorded successfully. Payment ID:{paymentID}")

    #creating labels
    student_id_label = Label(payments_window, text="Student ID").grid(row=1, column=0, sticky=W, padx=10)
    def search_student(event):
        student_id = student_id_box.get()
        cursor.execute("SELECT student_name, fees FROM students WHERE student_id = ?", (student_id,))
        student_name = cursor.fetchone()

        if student_name:
            student_name_text =  student_name[0]
        else:
            student_name_text = "No student found"

        # Update the labels
        student_name_from_id.config(text=f"{student_name_text}")

    student_name_from_id_label = Label(payments_window,text="Students Name")
    student_name_from_id_label.grid(row=2,column=0, sticky=W, padx=10)

    student_name_from_id = Label(payments_window, text="")
    student_name_from_id.grid(row=2,column=1)
    amount_paid_label = Label(payments_window, text="Amount Paid").grid(row=3, column=0, sticky=W, padx=10)
    payment_date_label = Label(payments_window, text="Payment Date").grid(row=4,column=0, sticky=W, padx=10)
    next_due__date_label = Label(payments_window, text="Next Due Date").grid(row=5,column=0, sticky=W, padx=10)
    payment_method_label = Label(payments_window, text="Payment Method").grid(row=6,column=0, sticky=W, padx=10)

    #Creating fields under payments_window
    student_id_box = Entry(payments_window)
    student_id_box.grid(row=1,column=1)
    student_id_box.bind("<KeyRelease>", search_student)

    amount_paid_box = Entry(payments_window)
    amount_paid_box.grid(row=3,column=1)

    payment_date_box = DateEntry(payments_window, date_pattern='dd-mm-yyyy')
    payment_date_box.grid(row=4, column=1)

    next_due_date_box = DateEntry(payments_window, date_pattern='dd-mm-yyyy')
    next_due_date_box.grid(row=5, column=1)

    payment_method_box = ttk.Combobox(payments_window, values=(' UPI',' Net Banking',' Cash'), state='readonly')
    payment_method_box.grid(row=6, column=1)

    gen_invoice_btn = Button(payments_window, text="Record Payment & Get Invoice", command=record_payment)  
    gen_invoice_btn.grid(row=7,column=0, columnspan=2) 

    # record_payment_btn = Button(payments_window, text="Record Payments", command=record_payment)  
    # record_payment_btn.grid(row=7,column=1)  

    cursor.execute("SELECT * FROM payments")
    payment_records = cursor.fetchall()
    for payment_record in payment_records:
            print(payment_record)

    payments_window.mainloop()    

#export function
def export_student_data():
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                title="Save Excel File As")
            if file_path:
                    df= pd.read_sql_query('SELECT * FROM students', mydb)     
                    df.to_excel(file_path, index=False)  

def export_payments_record():
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                title="Save Excel File As")
            if file_path:
                    df= pd.read_sql_query('SELECT * FROM payments', mydb)     
                    df.to_excel(file_path, index=False)  


# Define function to Edit Students data
def edit_student_data():
     editor = Toplevel()

     def delete_student():
          id_to_delete = student_id_editor.get()
          cursor.execute("DELETE FROM students WHERE oid=" + id_to_delete)
          mydb.commit()
          messagebox.showinfo("Success", "Student has been Removed.")

     def save_changes(): 
          
          selected_teachers = [teacher_listbox.get(i) for i in teacher_listbox.curselection()]
          selected_teachers_str = ', '.join(selected_teachers)
          selected_subjects = [subject_listbox.get(i) for i in subject_listbox.curselection()]
          selected_subjects_str = ', '.join(selected_subjects)

          cursor.execute("""UPDATE students SET
                         student_name = :student_name,
                         contact = :contact,
                         email = :email,
                         class = :class,
                         teacher_name = :teacher_name,
                         subjects = :subjects,
                         timing = :timing,
                         branch = :branch,
                         fees = :fees

                         WHERE oid = :oid""",
                         {
                            'student_name':student_name_editor.get(),
                            'contact':contact_editor.get(),
                            'email':email_editor.get(),
                            'class':class_editor.get(),
                            'teacher_name':selected_teachers_str,
                            'subjects': selected_subjects_str,
                            'timing':timing_editor.get(),
                            'branch':branch_editor.get(),
                            'fees':fees_editor.get(),

                            'oid':student_id_editor.get()
                         }
                         )
          mydb.commit()
          messagebox.showinfo("Success", "Changes saved successfully.")

     def get_teachers():
        cursor.execute("SELECT name FROM teachers")
        teachers = [row[0] for row in cursor.fetchall()]
        return teachers   

     def get_subjects():
        cursor.execute("SELECT name FROM subjects")
        subjects = [row[0] for row in cursor.fetchall()]
        return subjects

     def display_data(event):
          id_to_edit = student_id_editor.get()
          cursor.execute("SELECT * FROM branches")
          cursor.execute("SELECT * FROM students where student_id =" + id_to_edit)
          records = cursor.fetchall()
          #loop thru
          for record in records:
               student_name_editor.insert(0,record[1])
               contact_editor.insert(0,record[2])
               email_editor.insert(0,record[3])
               class_editor.set(record[4])
               timing_editor.insert(0,record[7])
               branch_editor.set(record[8])
               fees_editor.insert(0,record[9])

     def get_branches():
        cursor.execute("SELECT name FROM branches")
        branches = [row[0] for row in cursor.fetchall()]
        return branches 
     def get_classes():
        cursor.execute("SELECT name FROM classes")
        classes = [row[0] for row in cursor.fetchall()]
        return classes   

     #Create textBoxes
     student_id_editor = Entry(editor)
     student_id_editor.grid(row=0,column=1)
     student_id_editor.bind("<KeyRelease>", display_data)
     student_name_editor = Entry(editor)
     student_name_editor.grid(row=1,column=1)
     contact_editor = Entry(editor)
     contact_editor.grid(row=2,column=1)
     email_editor = Entry(editor)
     email_editor.grid(row=3,column=1)
     classes = get_classes()
     class_editor = ttk.Combobox(editor, values=classes , state='readonly')
     class_editor.grid(row=4,column=1)
     teachers = get_teachers()
     teacher_listbox = Listbox(editor, selectmode=MULTIPLE, exportselection=0)
     for teacher in teachers:
         teacher_listbox.insert(END, teacher)
     teacher_listbox.insert(END)
     teacher_listbox.grid(row=5, column=1, pady=5)
     subjects = get_subjects()
     subject_listbox = Listbox(editor, selectmode=MULTIPLE, exportselection=0)
     for subject in subjects:
         subject_listbox.insert(END, subject)
     subject_listbox.insert(END)
     subject_listbox.grid(row=6, column=1, pady=5)
     timing_editor = Entry(editor)
     timing_editor.grid(row=7,column=1)
     branches = get_branches()
     branch_editor = ttk.Combobox(editor, values=branches , state='readonly')
     branch_editor.grid(row=8,column=1)
     fees_editor = Entry(editor)
     fees_editor.grid(row=9,column=1)
    #  last_paid_on_editor = Entry(editor)
    #  last_paid_on_editor.grid(row=7,column=1)
    #  fee_status_editor = Entry(editor)
    #  fee_status_editor.grid(row=8,column=1)

     #creating labels
     student_id_label = Label(editor,text="Enter Student ID").grid(row=0, column=0, sticky=W, padx=10)
     student_name_label = Label(editor, text="Student Name").grid(row=1, column=0, sticky=W, padx=10)
     contact_label = Label(editor, text="Phone no.").grid(row=2, column=0, sticky=W, padx=10)
     email_label = Label(editor, text="E-mail").grid(row=3, column=0, sticky=W, padx=10)
     Class_label = Label(editor, text="Class").grid(row=4, column=0, sticky=W, padx=10)
     teacher_label = Label(editor, text="Teacher Name").grid(row=5, column=0, sticky=W, padx=10)
     subject_labe = Label(editor, text="Subjects").grid(row=6, column=0, sticky=W, padx=10)
     timing_label = Label(editor, text="Batch Timing").grid(row=7, column=0, sticky=W, padx=10)
     branch_label = Label(editor, text="Branch").grid(row=8, column=0, sticky=W, padx=10)
     fees_label = Label(editor, text="Monthly Fee").grid(row=9, column=0, sticky=W, padx=10)
     #creating save changes btn
     save_changes_btn = Button(editor, text="Save Changes", command=save_changes)
     save_changes_btn.grid(row=10, column=0,columnspan=2, sticky="nsew")

     remove_student_btn = Button(editor, text="Remove Student", command=delete_student)
     remove_student_btn.grid(row=11, column=0,columnspan=2, sticky="nsew")


def open_edit_student_data():
     edit_student_data()


# Create Manage Students section
manage_students_frame = Frame(root, bd=2, relief=GROOVE)
manage_students_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

Label(manage_students_frame, text="Manage Students", font=('Arial', 14, 'bold')).grid(row=0,sticky=W)

# Add buttons for Manage Students section
def open_register():
    register()

register_button = Button(manage_students_frame, text="Register Student", command=open_register)
register_button.grid(row=1,column=0,sticky="nsew")

edit_student_btn = Button(manage_students_frame, text="Edit Student Data", command=open_edit_student_data)
edit_student_btn.grid(row=1,column=1,sticky="nsew")

# search_label = Label(manage_students_frame, text="Search Student").grid(row=2,column=0,sticky=W)

# search_via_dropdown = ttk.Combobox(manage_students_frame, values=('Name','Phone no','ID'))
# search_via_dropdown.grid(row=2, column=1)

# search_for_box = Entry(manage_students_frame)
# search_for_box.grid(row=2,column=2,sticky="nsew")
# Create Payments section
payments_frame = Frame(root, bd=2, relief=GROOVE)
payments_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

Label(payments_frame, text="Payments", font=('Arial', 14, 'bold')).grid(row=0, sticky=W)

# Add buttons for Payments section
def open_payments():
    payments()

def send_reminder():
    account_sid = '...'
    auth_token = '...'
    client = Client(account_sid, auth_token)

    cursor.execute("SELECT student_name, fee_status, contact FROM students")
    students = cursor.fetchall()

    for student_name, fee_status, contact in students:
        if fee_status != 'Paid':
            message = client.messages.create(
            from_='whatsapp:',
            body=f"Dear parents, your Ward's fees is due. Please pay by 10th of this month.",
            to=f'whatsapp:+91{contact}'
            )    

def refresh_fee_status():
    cursor.execute("SELECT student_id, next_due_date FROM payments")
    students = cursor.fetchall()

    for student_id, next_due_date in students: 
        try:
            next_due_date_dt = datetime.datetime.strptime(next_due_date, '%d/%m/%y')
        except ValueError:
            try:
                next_due_date_dt = datetime.datetime.strptime(next_due_date, '%d-%m-%Y')
            except ValueError:
                continue

        if next_due_date_dt <= datetime.datetime.today():
            fee_status = 'Due'
        else:
            fee_status = 'Paid'
                
        # Update fee_status in students table
        cursor.execute("""UPDATE students SET
                            fee_status = :fee_status
                            WHERE student_id = :student_id""",
                        {'fee_status': fee_status, 'student_id': student_id}
                    )

    mydb.commit()
    messagebox.showinfo("Success", "Fee status updated for all students.")

payment_button = Button(payments_frame, text="Record Payments or Generate Invoices", command=open_payments)
payment_button.grid(row=1 , column=0, columnspan=2,sticky=NSEW)

refresh_fee_status_btn = Button(payments_frame, text="Refresh Fee Status", command=refresh_fee_status)
refresh_fee_status_btn.grid(row=2,column=0,sticky="nsew")

send_reminders_btn = Button(payments_frame, text="Send Fee Reminders", command=send_reminder)
send_reminders_btn.grid(row=2,column=1,sticky="nsew")
# Create Export section
export_frame = Frame(root, bd=2, relief=GROOVE)
export_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

Label(export_frame, text="Export", font=('Arial', 14, 'bold')).pack(pady=5)

# Add buttons for Export section
def export_student_data():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                             title="Save Excel File As")
    if file_path:
        df = pd.read_sql_query('SELECT * FROM students', mydb)     
        df.to_excel(file_path, index=False)  

def export_payments_record():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                             title="Save Excel File As")
    if file_path:
        df = pd.read_sql_query('SELECT * FROM payments', mydb)     
        df.to_excel(file_path, index=False)  

export_student_data_button = Button(export_frame, text="Export Students database to Excel", command=export_student_data)
export_student_data_button.pack(fill=X, padx=10, pady=5)

export_payments_data_button = Button(export_frame, text="Export Payments Record to Excel", command=export_payments_record)
export_payments_data_button.pack(fill=X, padx=10, pady=5)

# Configure rows and columns to expand with the window
root.columnconfigure((0, 1, 2), weight=1)
root.rowconfigure(0, weight=1)

root.mainloop()
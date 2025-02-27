import tkinter as tk
from db import connect_db
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
from tkcalendar import Calendar
import spire.doc as spd
import os

def fetch_leave_data(tree):
    """Fetch leave data from the database and display it."""
    db = connect_db()
    cursor = db.cursor()
    cursor.execute("""
        SELECT conges.id, employes.name, conges.start_date, conges.end_date, conges.left_day, conges.remaining_day, conges.year
        FROM conges 
        JOIN employes ON conges.employee_id = employes.id
    """)
    
    rows = cursor.fetchall()
    db.close()

    tree.delete(*tree.get_children())
    for row in rows:
        tree.insert("", "end", values=row)



def fill_leave_fields(tree, employee_id, start_date, end_date, left_day, remaining_day, year):
    """Fill fields with selected leave data."""
    selected_item = tree.selection()
    if not selected_item:
        return

    leave_data = tree.item(selected_item[0], "values")

    employee_id.set(leave_data[1]) 
    start_date.set(leave_data[2])
    end_date.set(leave_data[3])
    left_day.set(leave_data[4])
    remaining_day.set(leave_data[5])
    year.set(leave_data[6])

# def update_leave(tree, start_date, end_date, left_day, remaining_day):
#     """Update selected leave record."""
#     selected_item = tree.selection()
#     if not selected_item:
#         messagebox.showwarning("Erreur", "Sélectionnez un congé à modifier")
#         return

#     leave_id = tree.item(selected_item[0], "values")[0]
#     start = start_date.get()
#     end = end_date.get()
#     left_day = left_day.get()
#     remaining_day = remaining_day.get()

#     if not start or not end:
#         messagebox.showwarning("Erreur", "Tous les champs sont requis")
#         return

#     db = connect_db()
#     cursor = db.cursor()
#     cursor.execute("UPDATE conges SET start_date=%s, end_date=%s, left_day=%s, remaining_day=%s WHERE id=%s",
#                    (start, end, left_day, remaining_day, leave_id))
    
#     db.commit()
#     db.close()

#     messagebox.showinfo("Succès", "Congé modifié avec succès")
#     fetch_leave_data(tree)

def delete_leave(tree):
    """Delete selected leave record."""
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("خطا", "حدد العطلة المراد حذفها")
        return

    leave_id = tree.item(selected_item[0], "values")[0]

    db = connect_db()
    cursor = db.cursor()
    cursor.execute("DELETE FROM conges WHERE id=%s", (leave_id,))
    db.commit()
    db.close()

    messagebox.showinfo("نجاح", "تم إلغاء الإجازة")
    fetch_leave_data(tree)


def add_leave(tree, employee_id, start_date, end_date, left_day, remaining_day):
    """Adds a new leave record for an employee."""
    if not employee_id:
        messagebox.showerror("خطا", "يرجى تحديد موظف")
        return
    start = start_date.get().strip()
    end = end_date.get().strip()
    left_day = left_day.get().strip()
    remaining_day = remaining_day.get().strip()

    if not start or not end or not left_day or not remaining_day:
        messagebox.showerror("خطا", "يرجى إكمال جميع الحقول.")
        return

    current_year = datetime.strptime(start, "%Y/%m/%d").year
    db = connect_db()
    cursor = db.cursor()
    cursor.execute("INSERT INTO conges (employee_id, start_date, end_date, left_day, remaining_day, year) VALUES (%s, %s, %s, %s, %s, %s)",
                   (employee_id, start, end, left_day, remaining_day,current_year))
    db.commit()
    db.close()

    messagebox.showinfo("النجاح", "تمت إضافة الإجازة بنجاح.")
    start_date.set("")
    end_date.set("")
        
    if isinstance(left_day, tk.StringVar): 
            left_day.set("")
    else:
            left_day = tk.StringVar()  

    if isinstance(remaining_day, tk.StringVar):
            remaining_day.set("")
    else:
            remaining_day = tk.StringVar()
    fetch_leave_data(tree)

def extract_employee_id(employee_text):
    """Extracts employee ID from the formatted text."""
    try:
        return int(employee_text.split("(رقم: ")[1].strip(")")) 
    except (IndexError, ValueError):
        return None 

def get_remaining_days(employee_id):
    # db = connect_db()
    # cursor = db.cursor()
    # id = extract_employee_id(employee_id) 
    # query = "SELECT remaining_day FROM conges WHERE employee_id = %s ORDER BY id DESC LIMIT 1"
    # cursor.execute(query, (id,))
    # result = cursor.fetchone()
    # return result[0] if result else 22
    db = connect_db()
    cursor = db.cursor()

    current_year = datetime.now().year
    last_year = current_year - 1
    id = extract_employee_id(employee_id) 
    cursor.execute("""
        SELECT remaining_day FROM conges
        WHERE employee_id = %s AND year = %s
        ORDER BY id DESC LIMIT 1
    """, (id, last_year))

    last_year_remaining = cursor.fetchone()

    last_year_remaining = last_year_remaining[0] if last_year_remaining else 0

    cursor.execute("""
        SELECT SUM(left_day) FROM conges
        WHERE employee_id = %s AND year = %s
    """, (id, current_year))

    used_days = cursor.fetchone()[0]
    used_days = used_days if used_days else 0

    total_leave = 22 + last_year_remaining  

    remaining_days = max(0, total_leave - used_days)

    db.close()
    return remaining_days

def get_holidays():
    """Fetch holiday dates from the database."""
    try:
        db = connect_db()
        cursor = db.cursor()
        cursor.execute("SELECT date_holiday FROM holidays") 
        holidays = {datetime.strptime(row[0], "%Y-%m-%d").strftime("%Y-%m-%d") for row in cursor.fetchall()}  
        db.close()
        return holidays

    except ValueError:
        return set("خطا")
    
def calculate_dates(selected_employee_id, start_date, left_day, end_date, remaining_day):
    """Fetch input values, calculate end date and remaining days, and update fields while excluding weekends & holidays."""
    
    start = start_date.get().strip()
    left = left_day.get().strip()

    if start and left.isdigit():
        try:
            start_dt = datetime.strptime(start, "%Y/%m/%d") 
            left_days = int(left)

            emp_id = selected_employee_id.get()
            current_remaining_days = get_remaining_days(emp_id)

            if left_days > current_remaining_days:
                messagebox.showerror("خطأ", "لم يتبق ما يكفي من أيام الإجازة!")
                return
            
            holidays = get_holidays()

            days_added = 0
            current_date = start_dt
            while days_added < left_days:
                current_date += timedelta(days=1)
                
                if current_date.weekday() in (5, 6) or current_date.strftime("%Y-%m-%d") in holidays:
                    continue  
                
                days_added += 1  

            new_remaining_days = max(0, current_remaining_days - left_days)

            end_date.set(current_date.strftime("%Y/%m/%d"))
            remaining_day.set(str(new_remaining_days))

        except ValueError:
            end_date.set("")
            remaining_day.set("")
    else:
        end_date.set("")
        remaining_day.set("")

def open_calendar(start_date):
    """Open a popup calendar to select a date."""
    def set_date():
        start_date.set(cal.selection_get().strftime("%Y/%m/%d"))
        top.destroy()

    top = tk.Toplevel()
    top.title("اختيار التاريخ")

    cal = Calendar(top, date_pattern="yyyy/mm/dd") 
    cal.pack(pady=10)

    tk.Button(top, text="اختيار", command=set_date).pack(pady=5)

def search_conge(tree, search_entry):
    """Filters employee data based on the search term."""
    db = connect_db()
    cursor = db.cursor()
    search_text = search_entry.get().strip()
    try:
        query = "SELECT c.id, e.name, c.start_date, c.end_date, c.left_day, c.remaining_day, c.year " \
                    "FROM conges c JOIN employes e ON c.employee_id = e.id " \
                    "WHERE e.name LIKE %s"
        cursor.execute(query, (f"%{search_text}%",))
            
        rows = cursor.fetchall()
        db.close()

        for item in tree.get_children():
            tree.delete(item)

        for row in rows:
            tree.insert("", "end", values=row)

    except ValueError:
        messagebox.showerror("هناك خطا")

def reset_search(tree, search_var):
    """Reset search field and reload all leave records."""
    search_var.set("")  
    fetch_leave_data(tree)

def leave_data_word(name):
    """Fetch leave data from the database for a specific employee."""
    db = connect_db()
    cursor = db.cursor()
    query1 = """
        SELECT e.cin,e.lease_number, e.name_frame, e.nameFr, e.nameFrameFr
        FROM employes e
        WHERE e.name = %s
    """
    cursor.execute(query1,(name,))
    data = cursor.fetchone()

    db.close()
    return data

def generate_leave_document(tree, output_filename="generate.docx"):
    """Generate a leave document for an employee using the Word template."""
    selected_item = tree.selection()

    name = tree.item(selected_item[0], "values")[1]
    dateStart = tree.item(selected_item[0], "values")[2]
    leftDay = tree.item(selected_item[0], "values")[4]
 
    doc = spd.Document()
    doc.LoadFromFile("template.docx")

    leave_data = leave_data_word(name)

    if not leave_data:
        print("No leave data found for this employee.")
        return

    cin = leave_data[0]
    lease_number = leave_data[1]
    frame = leave_data[2]
    frameFr = leave_data[4]
    nameFr = leave_data[3]

    year = datetime.now().year
    dateNow = datetime.now().strftime("%d/%m/%Y")

    doc.Replace("name", name, True, True)
    doc.Replace("cin", cin, True, True)
    doc.Replace("frame", frame, True, True)
    doc.Replace("lease", lease_number, True, True)
    doc.Replace("leftDay", leftDay, True, True)
    doc.Replace("dateStart", dateStart, True, True)
    doc.Replace("nameFr", nameFr, True, True)
    doc.Replace("frameFr", frameFr, True, True)
    doc.Replace("year", str(year), True, True)
    doc.Replace("dateNow", dateNow, True, True)

    desktop_path = get_desktop_path()
    output_path = os.path.join(desktop_path, "output/"+output_filename)
    doc.SaveToFile(output_path, spd.FileFormat.Docx2016)
    doc.Close()
    os.startfile(output_path)
    print(f"Document generated: {output_path}")

def get_desktop_path():
    """Get the Desktop path for different operating systems."""
    return os.path.join(os.path.expanduser("~"), "bureau")

def open_conge_management():
    """Opens Leave Management Window."""
    conge_window = tk.Toplevel()
    conge_window.title("ادارة اجازات الموظفين")
    conge_window.geometry("900x750")
    conge_window.config(bg="#f4f4f4")

    label = tk.Label(conge_window, text="ادارة اجازات الموظفين", font=("Arial", 14, "bold"), bg="#f4f4f4")
    label.pack(pady=10)

    search_var = tk.StringVar()
    search_frame = tk.Frame(conge_window, bg="#f4f4f4")
    search_frame.pack(pady=5)

    tk.Label(search_frame, text="أسم الموظف", bg="#f4f4f4").pack(side="left", padx=5)
    search_entry = tk.Entry(search_frame, textvariable=search_var, width=25)
    search_entry.pack(side="left")

    search_btn = tk.Button(search_frame, text="بحث", bg="#2196F3", fg="white",
                           command=lambda: search_conge(tree, search_var))
    search_btn.pack(side="left", padx=5)

    reset_btn = tk.Button(search_frame, text="إعادة تعيين", bg="#FF5733", fg="white",
                          command=lambda: reset_search(tree, search_var))
    reset_btn.pack(side="left", padx=5)

    columns = ("رقم", "الاسم", "تاريخ المغادرة", "تاريخ الاستئناف", "مدة العطلة", "مدة متبقية", "في سنة")
    tree = ttk.Treeview(conge_window, columns=columns, show="headings")

    for col in columns:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=130, anchor="center")
    tree.pack(pady=10, fill="both", expand=True)

    fetch_leave_data(tree)

    frame = tk.Frame(conge_window, bg="#f4f4f4")
    frame.pack(pady=10)

    # Employee Selection
    tk.Label(frame, text="الاسم:", bg="#f4f4f4").grid(row=0, column=0)
    selected_employee = tk.StringVar()
    employee_entry = tk.Entry(frame, textvariable=selected_employee, state="readonly", width=25)
    employee_entry.grid(row=0, column=1)

    

    select_employee_btn = tk.Button(frame, text="اختيار الموظف", bg="#2196F3", fg="white",
                                    command=lambda: open_employee_selection(employee_entry))
    select_employee_btn.grid(row=0, column=2, padx=5)

    
    # Leave Date Inputs
    tk.Label(frame, text="تاريخ المغادرة", bg="#f4f4f4").grid(row=1, column=0)
    start_date = tk.StringVar()
    start_date_entry = tk.Entry(frame, textvariable=start_date)
    start_date_entry.grid(row=1, column=1)

    tk.Button(frame, text="📅", command=lambda: open_calendar(start_date)).grid(row=2, column=2, padx=5)

    # Left Days (Days taken)
    tk.Label(frame, text="مدة العطلة", bg="#f4f4f4").grid(row=2, column=0)
    left_day = tk.StringVar()
    left_day_entry = tk.Entry(frame, textvariable=left_day)
    left_day_entry.grid(row=2, column=1)

    # End Date (Calculated)
    tk.Label(frame, text="تاريخ الاستئناف", bg="#f4f4f4").grid(row=3, column=0)
    end_date = tk.StringVar()
    end_date_entry = tk.Entry(frame, textvariable=end_date, state="readonly")
    end_date_entry.grid(row=3, column=1)

    # Remaining Days (Calculated)
    tk.Label(frame, text="مدة متبقية", bg="#f4f4f4").grid(row=4, column=0)
    remaining_day = tk.StringVar()
    remaining_day_entry = tk.Entry(frame, textvariable=remaining_day, state="readonly")
    remaining_day_entry.grid(row=4, column=1)

    calcul_btn = tk.Button(frame, text="احتساب المدة", bg="#2196F3", fg="white",
                                    command=lambda: calculate_dates(employee_entry,start_date, left_day, end_date, remaining_day))
    calcul_btn.grid(row=1, column=2, padx=5)

    # Buttons
    btn_frame = tk.Frame(conge_window, bg="#f4f4f4")
    btn_frame.pack(pady=10)

    btn_add = tk.Button(btn_frame, text="اضافة", bg="#4CAF50", fg="white", 
                        command=lambda: add_leave(tree, extract_employee_id(selected_employee.get()), start_date, end_date, left_day, remaining_day))
    btn_add.grid(row=0, column=0, padx=5)

    # btn_update = tk.Button(btn_frame, text="تعديل", bg="#2196F3", fg="white",
    #                        command=lambda: update_leave(tree, start_date, end_date, left_day, remaining_day))
    # btn_update.grid(row=0, column=1, padx=5)

    btn_delete = tk.Button(btn_frame, text="حذف", bg="#FF5733", fg="white",
                           command=lambda: delete_leave(tree))
    btn_delete.grid(row=0, column=2, padx=5)

    btn_delete = tk.Button(btn_frame, text="طبع", bg="#2196F3", fg="white",
                           command=lambda: generate_leave_document(tree))
    btn_delete.grid(row=0, column=3, padx=5)

    year = tk.StringVar()
    tree.bind("<ButtonRelease-1>", lambda event: fill_leave_fields(tree, selected_employee, start_date, end_date, left_day, remaining_day, year))

    btn_close = tk.Button(conge_window, text="خروج", bg="gray", fg="white", command=conge_window.destroy)
    btn_close.pack(pady=10)

def search_employee(tree, search_entry, search_by):
    """Filters employee data based on the search term."""
    search_term = search_entry.get().strip().lower()

    for row in tree.get_children():
        tree.delete(row)  

    db = connect_db()
    cursor = db.cursor()

    if search_by == "Name":
        cursor.execute("SELECT * FROM employes WHERE LOWER(name) LIKE %s", (f"%{search_term}%",))
    elif search_by == "CIN":
        cursor.execute("SELECT * FROM employes WHERE LOWER(cin) LIKE %s", (f"%{search_term}%",))
    elif search_by == "Lease Number":
        cursor.execute("SELECT * FROM employes WHERE LOWER(lease_number) LIKE %s", (f"%{search_term}%",))
    elif search_by == "Name Frame":
        cursor.execute("SELECT * FROM employes WHERE LOWER(name_frame) LIKE %s", (f"%{search_term}%",))
    else:
        cursor.execute("SELECT * FROM employes") 

    rows = cursor.fetchall()
    db.close()

    for row in rows:
        tree.insert("", tk.END, values=row)

def fetch_employees(tree):
    """Fetch employees from database and display in table."""
    db = connect_db()
    cursor = db.cursor()
    cursor.execute("SELECT e.id, e.name, e.cin, e.lease_number, e.name_frame FROM employes e")
    rows = cursor.fetchall()
    db.close()

    for item in tree.get_children():
        tree.delete(item)

    for row in rows:
        tree.insert("", "end", values=row)

def open_employee_selection(employee_entry):
    """Open a window to select an employee from a datatable."""
    emp_window = tk.Toplevel()
    emp_window.title("اختيار الموظف")
    emp_window.geometry("600x500")
    emp_window.config(bg="#f4f4f4")

    label = tk.Label(emp_window, text="اختيار الموظف", font=("Arial", 14, "bold"), bg="#f4f4f4")
    label.pack(pady=10)

    filter_frame = tk.Frame(emp_window, bg="#f4f4f4")
    filter_frame.pack(pady=5)

    tk.Label(filter_frame, text="بحث", bg="#f4f4f4").grid(row=0, column=0, padx=5)
    search_entry = tk.Entry(filter_frame)
    search_entry.grid(row=0, column=1, padx=5)

    tk.Label(filter_frame, text="تصفية حسب", bg="#f4f4f4").grid(row=0, column=2, padx=5)
    filter_options = ["Name", "CIN", "Lease Number", "Name Frame"]
    selected_filter = tk.StringVar(value=filter_options[0])
    filter_dropdown = ttk.Combobox(filter_frame, textvariable=selected_filter, values=filter_options, state="readonly")
    filter_dropdown.grid(row=0, column=3, padx=5)

    columns = ("رقم", "الاسم", "رقم البطاقة الوطنية", "رقم التاجير", "اسم الاطار")
    tree = ttk.Treeview(emp_window, columns=columns, show="headings")

    for col in columns:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=150, anchor="center")
    tree.pack(pady=10, fill="both", expand=True)

    def search_employee():
        """Filters employee data based on the search term."""
        search_term = search_entry.get().strip().lower()
        
        for row in tree.get_children():
            tree.delete(row)  

        db = connect_db()
        cursor = db.cursor()

        if selected_filter.get() == "Name":
            cursor.execute("SELECT * FROM employes WHERE LOWER(name) LIKE %s", (f"%{search_term}%",))
        elif selected_filter.get() == "CIN":
            cursor.execute("SELECT * FROM employes WHERE LOWER(cin) LIKE %s", (f"%{search_term}%",))
        elif selected_filter.get() == "Lease Number":
            cursor.execute("SELECT * FROM employes WHERE LOWER(lease_number) LIKE %s", (f"%{search_term}%",))
        elif selected_filter.get() == "Name Frame":
            cursor.execute("SELECT * FROM employes WHERE LOWER(name_frame) LIKE %s", (f"%{search_term}%",))
        else:
            cursor.execute("SELECT * FROM employes")  

        rows = cursor.fetchall()
        db.close()

        for row in rows:
            tree.insert("", tk.END, values=row)

    btn_search = tk.Button(filter_frame, text="بحث", bg="#2196F3", fg="white",
                           command=search_employee)
    btn_search.grid(row=0, column=4, padx=5)

    btn_reset = tk.Button(filter_frame, text="إعادة تعيين", bg="#FF5733", fg="white",
                          command=lambda: fetch_employees(tree))
    btn_reset.grid(row=0, column=5, padx=5)

    fetch_employees(tree)

    def select_employee():
        """Get selected employee and update main leave form."""
        selected_item = tree.selection()
        if selected_item:
            values = tree.item(selected_item, "values")
            if values:
                employee_entry.config(state="normal") 
                employee_entry.delete(0, tk.END)
                employee_entry.insert(0, f"{values[1]} (رقم: {values[0]})")
                employee_entry.config(state="readonly") 
                emp_window.destroy()
            else:
                print("No employee selected!") 

    btn_select = tk.Button(emp_window, text="اختيار", bg="#4CAF50", fg="white", command=select_employee)
    btn_select.pack(pady=10)


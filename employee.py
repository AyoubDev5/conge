import tkinter as tk
from db import connect_db
from tkinter import ttk, messagebox

page_number = 1
rows_per_page = 20

def fetch_employees(tree):
    """Fetch employees from database and display in table."""
    db = connect_db()
    cursor = db.cursor()
    cursor.execute("SELECT * FROM employes")
    rows = cursor.fetchall()
    db.close()

    for item in tree.get_children():
        tree.delete(item)

    for row in rows:
        tree.insert("", "end", values=row)

def add_employee(tree, name, cin, lease_number, name_frame):
    """Add a new employee to the database."""
    name = name.get()
    cin = cin.get()
    lease_number = lease_number.get()
    name_frame = name_frame.get()

    if not name or not cin or not lease_number or not name_frame:
        messagebox.showwarning("خطأ في الإدخال", "جميع الحقول مطلوبة")
        return

    db = connect_db()
    cursor = db.cursor()
    cursor.execute("INSERT INTO employes (name, cin, lease_number, name_frame) VALUES (%s, %s, %s, %s)", 
                   (name, cin, lease_number, name_frame))
    db.commit()
    db.close()

    messagebox.showinfo("النجاح", "موظف تمت إضافته بنجاح")
    fetch_employees(tree) 

def clear_fields(name, cin, lease_number, name_frame):
    """Clears input fields."""
    name.delete(0, tk.END)
    cin.delete(0, tk.END)
    lease_number.delete(0, tk.END)
    name_frame.delete(0, tk.END)

def fill_update_fields(tree, name, cin, lease_number, name_frame, selected_id):
    """Fills input fields with selected employee's data or clears them if deselected."""
    selected_item = tree.selection()
    if not selected_item:

        clear_fields(name, cin, lease_number, name_frame)
        selected_id.set("")
        return

    employee_data = tree.item(selected_item[0], "values")

    if selected_id.get() == employee_data[0]:  
        tree.selection_remove(selected_item[0])
        clear_fields(name, cin, lease_number, name_frame)
        selected_id.set("")
    else:
        name.delete(0, tk.END)
        name.insert(0, employee_data[1])

        cin.delete(0, tk.END)
        cin.insert(0, employee_data[2])

        lease_number.delete(0, tk.END)
        lease_number.insert(0, employee_data[3])

        name_frame.delete(0, tk.END)
        name_frame.insert(0, employee_data[4])

        selected_id.set(employee_data[0])


def update_employee(tree, name, cin, lease_number, name_frame):
    """Update an existing employee."""
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("خطأ في الاختيار", "تحديد موظف لتحديثه")
        return

    employee_id = tree.item(selected_item[0], "values")[0]
    name = name.get()
    cin = cin.get()
    lease_number = lease_number.get()
    name_frame = name_frame.get()

    if not name or not cin or not lease_number or not name_frame:
        messagebox.showwarning("خطأ في الإدخال", "جميع الحقول مطلوبة")
        return

    db = connect_db()
    cursor = db.cursor()
    cursor.execute("UPDATE employes SET name=%s, cin=%s, lease_number=%s,  name_frame=%s WHERE id=%s",
                   (name, cin, lease_number, name_frame, employee_id))
    db.commit()
    db.close()

    messagebox.showinfo("النجاح", "تم تحديث الموظف بنجاح")
    fetch_employees(tree)

def delete_employee(tree):
    """Delete an employee from the database."""
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("خطأ في الاختيار", "حدد موظفاً لحذفه")
        return

    employee_id = tree.item(selected_item[0], "values")[0]

    db = connect_db()
    cursor = db.cursor()
    cursor.execute("DELETE FROM employes WHERE id=%s", (employee_id,))
    db.commit()
    db.close()

    messagebox.showinfo("النجاح", "تم حذف الموظف بنجاح")
    fetch_employees(tree)

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

def fetch_paginated_employees(tree, page_label, btn_next, btn_prev, search_query=None, filter_by=None):
    """Fetches paginated employee data from the database."""
    global page_number

    db = connect_db()
    cursor = db.cursor()

    query = "SELECT id, name, cin, lease_number, name_frame FROM employes"
    params = []

    if search_query and filter_by:
        query += f" WHERE {filter_by.lower().replace(' ', '_')} LIKE %s"
        params.append(f"%{search_query}%")

    cursor.execute(query, params)
    total_rows = len(cursor.fetchall())

    query += " LIMIT %s OFFSET %s"
    params.extend([rows_per_page, (page_number - 1) * rows_per_page])

    cursor.execute(query, params)
    rows = cursor.fetchall()
    db.close()

    for row in tree.get_children():
        tree.delete(row)

    for row in rows:
        tree.insert("", "end", values=row)

    total_pages = max(1, (total_rows + rows_per_page - 1) // rows_per_page)
    page_label.config(text=f"صفحة {page_number} على {total_pages}")

    if (page_number * rows_per_page) >= total_rows:
        btn_next.config(state="disabled")
    else:
        btn_next.config(state="normal")

    if page_number == 1:
        btn_prev.config(state="disabled")
    else:
        btn_prev.config(state="normal")

def update_pagination(tree, total_rows, page_label):
    """Updates pagination buttons and label."""
    total_pages = max(1, (total_rows + rows_per_page - 1) // rows_per_page)
    page_label.config(text=f"Page {page_number} of {total_pages}")

def next_page(tree, page_label, btn_next, btn_prev, search_entry, filter_dropdown):
    """Go to the next page."""
    global page_number
    page_number += 1
    fetch_paginated_employees(tree, page_label, btn_next, btn_prev, search_entry.get(), filter_dropdown.get())

def prev_page(tree, page_label, btn_next, btn_prev, search_entry, filter_dropdown):
    """Go to the previous page."""
    global page_number
    if page_number > 1:
        page_number -= 1
    fetch_paginated_employees(tree, page_label, btn_next, btn_prev, search_entry.get(), filter_dropdown.get())

def open_employee_management():
    """Opens Employee Management Window."""
    global page_number

    emp_window = tk.Toplevel()
    emp_window.title("إدارة الموظفين")
    emp_window.geometry("600x500")
    emp_window.config(bg="#f4f4f4")

    label = tk.Label(emp_window, text="إدارة شؤون الموظفين", font=("Arial", 14, "bold"), bg="#f4f4f4")
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

    btn_search = tk.Button(filter_frame, text="بحث", bg="#2196F3", fg="white",
                           command=lambda: search_employee(tree, search_entry, selected_filter.get()))
    btn_search.grid(row=0, column=4, padx=5)

    btn_reset = tk.Button(filter_frame, text="إعادة تعيين", bg="#FF5733", fg="white",
                          command=lambda: fetch_employees(tree))
    btn_reset.grid(row=0, column=5, padx=5)

    columns = ("الصف", "الاسم", "رقم البطاقة الوطنية", "رقم التاجير", "اسم الاطار")
    tree = ttk.Treeview(emp_window, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=150, anchor="center")
    tree.pack(pady=10, fill="both", expand=True)

    fetch_employees(tree) 

    pagination_frame = tk.Frame(emp_window, bg="#f4f4f4")
    pagination_frame.pack(pady=5)

    btn_prev = tk.Button(pagination_frame, text="<< السابق", bg="#2196F3", fg="white", state="disabled",
                         command=lambda: prev_page(tree, page_label, btn_next, btn_prev, search_entry, filter_dropdown))
    btn_prev.grid(row=0, column=0, padx=5)

    page_label = tk.Label(pagination_frame, text="صفحة 1", bg="#f4f4f4", font=("Arial", 10, "bold"))
    page_label.grid(row=0, column=1, padx=10)

    btn_next = tk.Button(pagination_frame, text="التالي >>", bg="#2196F3", fg="white",
                         command=lambda: next_page(tree, page_label, btn_next, btn_prev, search_entry, filter_dropdown))
    btn_next.grid(row=0, column=2, padx=5)

    frame = tk.Frame(emp_window, bg="#f4f4f4")
    frame.pack(pady=10)

    tk.Label(frame, text="الاسم", bg="#f4f4f4").grid(row=0, column=0)
    name = tk.Entry(frame)
    name.grid(row=0, column=1) 

    tk.Label(frame, text="رقم البطاقة الوطنية", bg="#f4f4f4").grid(row=1, column=0)
    cin = tk.Entry(frame)
    cin.grid(row=1, column=1)

    tk.Label(frame, text="رقم التاجير", bg="#f4f4f4").grid(row=2, column=0)
    lease_number = tk.Entry(frame)
    lease_number.grid(row=2, column=1)

    tk.Label(frame, text="اسم الاطار", bg="#f4f4f4").grid(row=3, column=0)
    name_frame = tk.Entry(frame)
    name_frame.grid(row=3, column=1)

    btn_frame = tk.Frame(emp_window, bg="#f4f4f4")
    btn_frame.pack(pady=10)

    btn_add = tk.Button(btn_frame, text="إضافة", bg="#4CAF50", fg="white",
                        command=lambda: add_employee(tree, name, cin, lease_number, name_frame))
    btn_add.grid(row=0, column=0, padx=5)

    btn_update = tk.Button(btn_frame, text="تحديث", bg="#2196F3", fg="white",
                           command=lambda: update_employee(tree, name, cin, lease_number, name_frame))
    btn_update.grid(row=0, column=1, padx=5)

    btn_delete = tk.Button(btn_frame, text="حذف", bg="#FF5733", fg="white",
                           command=lambda: delete_employee(tree))
    btn_delete.grid(row=0, column=2, padx=5)

    selected_id = tk.StringVar()
    tree.bind("<ButtonRelease-1>", lambda event: fill_update_fields(tree, name, cin, lease_number, name_frame, selected_id))

    btn_close = tk.Button(emp_window, text="إغلاق", bg="gray", fg="white", command=emp_window.destroy)
    btn_close.pack(pady=10)

    fetch_paginated_employees(tree, page_label, btn_next, btn_prev)

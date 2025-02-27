import tkinter as tk
from db import connect_db
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
from tkcalendar import Calendar
import spire.doc as spd
import os

def fetch_attestation_data(tree):
    """Récupérer et afficher les attestations."""
    db = connect_db()
    cursor = db.cursor()
    cursor.execute("""
        SELECT attestation.numeroSerie, employes.name, attestation.date, attestation.objectif, 
               attestation.langageAttestation, attestation.remis
        FROM attestation 
        JOIN employes ON attestation.employeeId = employes.id
    """)
    
    rows = cursor.fetchall()
    db.close()

    tree.delete(*tree.get_children())
    for row in rows:
        tree.insert("", "end", values=row)



def fill_leave_fields(tree, employee_id, objectif, langageAttestation, remis):
    """Fill fields with selected leave data."""
    selected_item = tree.selection()
    if not selected_item:
        return

    leave_data = tree.item(selected_item[0], "values")

    employee_id.set(leave_data[1]) 
    objectif.set(leave_data[3])
    langageAttestation.set(leave_data[4])
    remis.set(leave_data[5])

def delete_leave(tree):
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("خطأ", "الرجاء تحديد الشهادة المراد حذفها.")
        return

    attestation_id = tree.item(selected_item[0], "values")[0]

    db = connect_db()
    cursor = db.cursor()
    cursor.execute("DELETE FROM attestation WHERE numeroSerie=%s", (attestation_id,))
    db.commit()
    db.close()

    messagebox.showinfo("النجاح", "تمت إزالة الشهادة بنجاح.")
    fetch_attestation_data(tree)

def add_attestation(tree, employee_id, objectif, langageAttestation, remis):
    """Ajouter une nouvelle attestation."""
    if not employee_id:
        messagebox.showerror("خطأ", "يرجى اختيار موظف.")
        return

    current_date = datetime.today()
    current_year = datetime.now().year

    db = connect_db()
    cursor = db.cursor()

    cursor.execute("""
        SELECT numeroSerie FROM attestation 
        WHERE YEAR(date) = %s 
        ORDER BY numeroSerie DESC 
        LIMIT 1
    """, (current_year,))
    
    last_numero = cursor.fetchone()

    if last_numero:
        next_numero = int(last_numero[0]) + 1
    else:
        next_numero = 1 

    cursor.execute("""
        INSERT INTO attestation (numeroSerie, date, objectif, langageAttestation, remis, employeeId) 
        VALUES (%s, %s, %s, %s, %s, %s)
    """, (next_numero, current_date, objectif.get(), langageAttestation.get(), remis.get(), employee_id))

    db.commit()
    db.close()

    messagebox.showinfo("النجاح", f"تمت إضافة الشهادة بنجاح. الرقم التسلسلي: {next_numero}")
    fetch_attestation_data(tree)

def search_attestation(tree, search_entry):
    """Rechercher les attestations par nom d'employé."""
    db = connect_db()
    cursor = db.cursor()
    search_text = search_entry.get().strip()

    query = """
        SELECT attestation.numeroSerie, employes.name, attestation.date, attestation.objectif, 
               attestation.langageAttestation, attestation.remis
        FROM attestation 
        JOIN employes ON attestation.employeeId = employes.id
        WHERE employes.name LIKE %s
    """
    
    cursor.execute(query, (f"%{search_text}%",))
    rows = cursor.fetchall()
    db.close()

    tree.delete(*tree.get_children())
    for row in rows:
        tree.insert("", "end", values=row)

def reset_search(tree, search_var):
    """Reset search field and reload all leave records."""
    search_var.set("")  
    fetch_attestation_data(tree)

def leave_data_word(name):
    """Fetch leave data from the database for a specific employee."""
    db = connect_db()
    cursor = db.cursor()
    query1 = """
        SELECT e.cin,e.lease_number, e.name_frame, e.nameFr, e.nameFrameFr
        FROM employes e
        WHERE e.name = ?
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

def extract_employee_id(employee_text):
    """Extracts employee ID from the formatted text."""
    try:
        return int(employee_text.split("(رقم: ")[1].strip(")")) 
    except (IndexError, ValueError):
        return None 

def open_attestation_management():
    """Opens Leave Management Window."""
    attestation_window = tk.Toplevel()
    attestation_window.title("ادارة شهادات العمل")
    attestation_window.geometry("900x750")
    attestation_window.config(bg="#f4f4f4")

    label = tk.Label(attestation_window, text="ادارة شهادات العمل", font=("Arial", 14, "bold"), bg="#f4f4f4")
    label.pack(pady=10)

    search_var = tk.StringVar()
    search_frame = tk.Frame(attestation_window, bg="#f4f4f4")
    search_frame.pack(pady=5)

    tk.Label(search_frame, text="أسم الموظف", bg="#f4f4f4").pack(side="left", padx=5)
    search_entry = tk.Entry(search_frame, textvariable=search_var, width=25)
    search_entry.pack(side="left")

    search_btn = tk.Button(search_frame, text="بحث", bg="#2196F3", fg="white",
                           command=lambda: search_attestation(tree, search_entry))
    search_btn.pack(side="left", padx=5)

    reset_btn = tk.Button(search_frame, text="إعادة تعيين", bg="#FF5733", fg="white",
                          command=lambda: reset_search(tree, search_var))
    reset_btn.pack(side="left", padx=5)

    columns = ("الرقم التسلسلي", "الاسم", "تاريخ", "الغاية", "لغة الشهادة", "المسلم له")
    tree = ttk.Treeview(attestation_window, columns=columns, show="headings")

    for col in columns:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=130, anchor="center")
    tree.pack(pady=10, fill="both", expand=True)

    fetch_attestation_data(tree)

    frame = tk.Frame(attestation_window, bg="#f4f4f4")
    frame.pack(pady=10)

    tk.Label(frame, text="الاسم:", bg="#f4f4f4").grid(row=0, column=0)
    selected_employee = tk.StringVar()
    employee_entry = tk.Entry(frame, textvariable=selected_employee, state="readonly", width=25)
    employee_entry.grid(row=0, column=1)

    select_employee_btn = tk.Button(frame, text="اختيار الموظف", bg="#2196F3", fg="white",
                                    command=lambda: open_employee_selection(employee_entry))
    select_employee_btn.grid(row=0, column=2, padx=5)

    tk.Label(frame, text=" الغاية", bg="#f4f4f4").grid(row=1, column=0)
    objectif = tk.StringVar()
    objectif_entry = tk.Entry(frame, textvariable=objectif)
    objectif_entry.grid(row=1, column=1)

    tk.Label(frame, text=" لغة الشهادة", bg="#f4f4f4").grid(row=2, column=0)
    langageAttestation = tk.StringVar(value="عربية") 

    tk.Radiobutton(frame, text="فرنسية", variable=langageAttestation, value="فرنسية", bg="#f4f4f4").grid(row=2, column=1, sticky="w")
    tk.Radiobutton(frame, text="عربية", variable=langageAttestation, value="عربية", bg="#f4f4f4").grid(row=2, column=2, sticky="w")

    tk.Label(frame, text="المسلم له", bg="#f4f4f4").grid(row=3, column=0)
    remis = tk.StringVar()
    remis_entry = tk.Entry(frame, textvariable=remis)
    remis_entry.grid(row=3, column=1)

    # Buttons
    btn_frame = tk.Frame(attestation_window, bg="#f4f4f4")
    btn_frame.pack(pady=10)

    btn_add = tk.Button(btn_frame, text="اضافة", bg="#4CAF50", fg="white", 
                        command=lambda: add_attestation(tree, extract_employee_id(selected_employee.get()), objectif, langageAttestation, remis))
    btn_add.grid(row=0, column=0, padx=5)

    btn_delete = tk.Button(btn_frame, text="حذف", bg="#FF5733", fg="white",
                           command=lambda: delete_leave(tree))
    btn_delete.grid(row=0, column=2, padx=5)

    btn_delete = tk.Button(btn_frame, text="طبع", bg="#2196F3", fg="white",
                           command=lambda: generate_leave_document(tree))
    btn_delete.grid(row=0, column=3, padx=5)

    tree.bind("<ButtonRelease-1>", lambda event: fill_leave_fields(tree, selected_employee, objectif, langageAttestation, remis))

    btn_close = tk.Button(attestation_window, text="خروج", bg="gray", fg="white", command=attestation_window.destroy)
    btn_close.pack(pady=10)

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
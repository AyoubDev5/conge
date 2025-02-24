import tkinter as tk
from navigation import open_page
from employee import open_employee_management
from conge import open_conge_management

def open_dashboard(root):
    """Admin dashboard with only Employee Management and Leave Management."""
    dashboard = tk.Toplevel()
    dashboard.title("لوحة تحكم المسؤول")
    dashboard.geometry("500x400")
    dashboard.config(bg="#f4f4f4")

    dashboard.protocol("WM_DELETE_WINDOW", lambda: logout(dashboard, root))

    label_dashboard = tk.Label(dashboard, text="لوحة تحكم المسؤول", font=("Arial", 14, "bold"), bg="#f4f4f4")
    label_dashboard.pack(pady=20)

    # Employee Management Button
    btn_employee = tk.Button(dashboard, text="إدارة الموظفين", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white",
                             command=open_employee_management)
    btn_employee.pack(pady=10)

    # Employee Leave Management Button
    btn_conge = tk.Button(dashboard, text="الإجازة ", font=("Arial", 12, "bold"), bg="#2196F3", fg="white",
                          command=open_conge_management)
    btn_conge.pack(pady=10)

    # Logout Button
    btn_logout = tk.Button(dashboard, text="تسجيل الخروج", font=("Arial", 12, "bold"), bg="#FF5733", fg="white",
                           command=lambda: logout(dashboard, root))
    btn_logout.pack(pady=20)

def logout(dashboard_window, root):
    """Logout and return to login screen."""
    dashboard_window.destroy()
    root.deiconify()

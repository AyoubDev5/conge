import tkinter as tk
from tkinter import messagebox
from tkinter import font
from PIL import Image, ImageTk
from dashboard import open_dashboard
import os
import sys


if __name__ == "__main__":
    ADMIN_USERNAME = "sara"
    ADMIN_PASSWORD = "user.123"

    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS 
        except Exception:
            base_path = os.path.abspath(".")  
        return os.path.join(base_path, relative_path)

    def loginWindow(root, entry_username, entry_password):
        """Validate login credentials and navigate to the dashboard."""
        
        username = entry_username.get().strip()  
        password = entry_password.get().strip()

        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            root.withdraw()  
            open_dashboard(root)  
        else:
            messagebox.showerror("فشل تسجيل الدخول", "اسم المستخدم أو كلمة المرور غير صحيحة!")

    def close_login(root):
        """Close login window properly."""
        root.destroy()  

    
    root = tk.Tk()
    root.title("تسجيل دخول المسؤول")
    root.geometry("400x350") 
    root.config(bg="#f4f4f4")  


    logo_path = resource_path("logo.png")
    logo = Image.open(logo_path)  
    logo = logo.resize((100, 100)) 
    logo_img = ImageTk.PhotoImage(logo)

    label_logo = tk.Label(root, image=logo_img, bg="#f4f4f4")
    label_logo.pack(pady=20)


    label_font = font.Font(family="Helvetica", size=12, weight="bold")
    entry_font = font.Font(family="Helvetica", size=12)

  
    label_username = tk.Label(root, text="اسم المستخدم", font=label_font, bg="#f4f4f4", anchor="center")
    label_username.pack(pady=5, padx=20, fill="x")

    entry_username = tk.Entry(root, font=entry_font)
    entry_username.pack(pady=5, padx=20, fill="x")

  
    label_password = tk.Label(root, text="كلمة المرور", font=label_font, bg="#f4f4f4", anchor="center")
    label_password.pack(pady=5, padx=20, fill="x")

    entry_password = tk.Entry(root, show="*", font=entry_font)
    entry_password.pack(pady=5, padx=20, fill="x")

 
    login_button = tk.Button(root, text="تسجيل الدخول", font=label_font, bg="#4CAF50", fg="white",
                            command=lambda: loginWindow(root, entry_username, entry_password))
    login_button.pack(pady=20)


    root.protocol("WM_DELETE_WINDOW", lambda: close_login(root))

    root.mainloop()

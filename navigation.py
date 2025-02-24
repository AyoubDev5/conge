import tkinter as tk

def open_page(page_name):
    page = tk.Toplevel()
    page.title(page_name)
    page.geometry("400x300")
    page.config(bg="#f4f4f4")

    label = tk.Label(page, text=f"Welcome to {page_name} Page", font=("Arial", 14), bg="#f4f4f4")
    label.pack(pady=20)

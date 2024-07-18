import requests
import json
import tkinter as tk
from tkinter import messagebox
from test import windows
from ship_co import get_ebay_data

def auth_data_save():
    ebay_id = ebay_id_entry.get()
    ebay_pass = ebay_pass_entry.get()
    ship_id = ship_id_entry.get()
    ship_pass = ship_pass_entry.get()
    product_sheet_id = product_manager_sheet_entry.get()
    order_sheet_id = order_manager_sheet_entry.get()
    total_data = total_check_box.get()
    recent_data = recent_check_box.get()
    checkbox_val = ""
    if ebay_id == '':
        messagebox.showwarning("警告", "ebayユーザーIDは必須です。")
        return
    if ebay_pass == '':
        messagebox.showwarning("警告", "ebayユーザーPASSは必須です。")
        return
    if product_sheet_id == '':
        messagebox.showwarning("警告", "商品管理シートIDは必須です。")
        return
    if order_sheet_id == '':
        messagebox.showwarning("警告", "注文管理シートIDは必須です。")
        return
    if total_data == True:
        checkbox_val = 0
    if recent_data == True:
        checkbox_val = 1
    if total_data == False and recent_data == False:
        messagebox.showwarning("警告", "データ取得期間は必須です。")
        return
    if total_data == True and recent_data == True:
        messagebox.showwarning("警告", "1つの項目だけを選択してください。")
        return
    data = {
        "ebay_id": ebay_id,
        "ebay_pass": ebay_pass,
        "ship_id": ship_id,
        "ship_pass": ship_pass,
        "product_sheet_id": product_sheet_id,
        "order_sheet_id": order_sheet_id,
        "total_data": total_data,
        "recent_data": recent_data
    }

    # Save the data to a JSON file
    with open("auth_data.json", "w") as f:
        json.dump(data, f, indent=4)
    messagebox.showinfo("OK", "保存されました。")
def read_auth_data():
    try:
        with open("auth_data.json", "r") as f:
            data = json.load(f)
    except FileNotFoundError:
        print("auth_data.json file not found. Please save the authentication data first.")
        return None
    return data

def submit_login():
    ebay_id = ebay_id_entry.get()
    ebay_pass = ebay_pass_entry.get()
    ship_id = ship_id_entry.get()
    ship_pass = ship_pass_entry.get()
    product_sheet_id = product_manager_sheet_entry.get()
    order_sheet_id = order_manager_sheet_entry.get()
    total_data = total_check_box.get()
    recent_data = recent_check_box.get()
    checkbox_val = ""
    if ebay_id == '':
        messagebox.showwarning("警告", "ebayユーザーIDは必須です。")
        return
    if ebay_pass == '':
        messagebox.showwarning("警告", "ebayユーザーPASSは必須です。")
        return
    if product_sheet_id == '':
        messagebox.showwarning("警告", "商品管理シートIDは必須です。")
        return
    if order_sheet_id == '':
        messagebox.showwarning("警告", "注文管理シートIDは必須です。")
        return
    if total_data == True:
        checkbox_val = 0
    if recent_data == True:
        checkbox_val = 1
    if total_data == False and recent_data == False:
        messagebox.showwarning("警告", "データ取得期間は必須です。")
        return
    if total_data == True and recent_data == True:
        messagebox.showwarning("警告", "1つの項目だけを選択してください。")
        return

    if not(ship_id == '' or ship_pass == ''):
        windows(ebay_id, ebay_pass, ship_id, ship_pass, product_sheet_id, order_sheet_id, checkbox_val)
    else:
        get_ebay_data(ebay_id, ebay_pass, product_sheet_id, order_sheet_id, checkbox_val)

# Create the main window
login_window = tk.Tk()
login_window.title("販売管理シートツール")
login_window.geometry('300x400')
# ebayTitle
ebay_title = tk.Label(login_window, text="ebayアカウント")
ebay_title.pack()
ebay_title.place(x=10, y=30)
# ebayId input
ebay_id_label = tk.Label(login_window, text="ユーザーID :")
ebay_id_label.pack()
ebay_id_label.place(x=20, y=60)
ebay_id_entry = tk.Entry(login_window)
ebay_id_entry.pack()
ebay_id_entry.place(x=90, y=60, width=185)
# ebayPASS input
ebay_pass_label = tk.Label(login_window, text="ユーザーPASS :")
ebay_pass_label.pack()
ebay_pass_label.place(x=20, y=90)
ebay_pass_entry = tk.Entry(login_window, show="*")
ebay_pass_entry.pack()
ebay_pass_entry.place(x=90, y=90, width=185)
# ship&coTitle
ebay_title = tk.Label(login_window, text="ship&coアカウント")
ebay_title.pack()
ebay_title.place(x=10, y=120)
# ship&coId input
ship_id_label = tk.Label(login_window, text="ユーザーID :")
ship_id_label.pack()
ship_id_label.place(x=20, y=150)
ship_id_entry = tk.Entry(login_window)
ship_id_entry.pack()
ship_id_entry.place(x=90, y=150, width=185)
# ship&coPASS input
ship_pass_label = tk.Label(login_window, text="ユーザーPASS :")
ship_pass_label.pack()
ship_pass_label.place(x=20, y=180)
ship_pass_entry = tk.Entry(login_window, show="*")
ship_pass_entry.pack()
ship_pass_entry.place(x=90, y=180, width=185)
# product manager sheet input
product_manager_sheet_label = tk.Label(login_window, text="商品管理シートID :")
product_manager_sheet_label.pack()
product_manager_sheet_label.place(x=20, y=230)
product_manager_sheet_entry = tk.Entry(login_window)
product_manager_sheet_entry.pack()
product_manager_sheet_entry.place(x=20, y=250, width=260)
# order manager sheet input
order_manager_sheet_label = tk.Label(login_window, text="注文管理シートID :")
order_manager_sheet_label.pack()
order_manager_sheet_label.place(x=20, y=280)
order_manager_sheet_entry = tk.Entry(login_window)
order_manager_sheet_entry.pack()
order_manager_sheet_entry.place(x=20, y=300, width=260)
# total
total_check_box = tk.BooleanVar()
total_checkbox = tk.Checkbutton(login_window, text="すべてのデータを取得", variable=total_check_box)
total_checkbox.place(x=16, y=330)
# recent
recent_check_box = tk.BooleanVar()
recent_checkbox = tk.Checkbutton(login_window, text="最近月データを取得", variable=recent_check_box)
recent_checkbox.place(x=160, y=330)
def insert_values():
    auth_data = read_auth_data()
    if auth_data is not None:
        ebay_id_entry.insert(0, auth_data["ebay_id"])
        ebay_pass_entry.insert(0, auth_data["ebay_pass"])
        ship_id_entry.insert(0, auth_data["ship_id"])
        ship_pass_entry.insert(0, auth_data["ship_pass"])
        product_manager_sheet_entry.insert(0, auth_data["product_sheet_id"])
        order_manager_sheet_entry.insert(0, auth_data["order_sheet_id"])
        total_check_box.set(auth_data["total_data"])
        recent_check_box.set(auth_data["recent_data"])
insert_values()
# Submit button
save_button = tk.Button(login_window, text="保存", command=auth_data_save)
save_button.pack()
save_button.place(x=200, y=360)
submit_button = tk.Button(login_window, text="起動", command=submit_login)
submit_button.pack()
submit_button.place(x=245, y=360)
login_window.mainloop()
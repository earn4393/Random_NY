import tkinter as tk
from tkinter import filedialog
import pandas as pd
import random
import time  # เพิ่ม import สำหรับ time.sleep()


# อ่านไฟล์ Excel
def read_excel_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")])
    df = pd.read_excel(file_path)
    return df

# เพิ่มไฟล์ Excel ใหม่
def add_new_excel_file():
    global employee_df, employee_names
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")])
    employee_df = pd.read_excel(file_path)
    employee_names = list(employee_df["ชื่อพนักงาน"])
    result_label.config(text="")
    remaining_label.config(text=f"รายชื่อที่เหลือ: {len(employee_names)} คน")


# สุ่มรหัสพนักงาน
def generate_random_code():
    generate_button.config(state="disabled")  # กลบปุ่มสุ่ม
    time.sleep(2)  # ล่าช้า 2 วินาที
    if len(employee_names) > 0:
        random_employee = random.choice(employee_names)
        employee_names.remove(random_employee)
        result_label.config(text=random_employee)
        remaining_label.config(
            text=f"รายชื่อที่เหลือ: {len(employee_names)} คน")
    else:
        result_label.config(text="ไม่มีรายชื่อพนักงานที่เหลือ")
    generate_button.config(state="active")  # เปิดปุ่มสุ่มอีกครั้ง

# เริ่มต้น GUI
root = tk.Tk()
root.title("แจกรางวัลปีใหม่")

# เลือกธีม
root.tk_setPalette(background="#f0f0f0")

# การจัดวาง
label = tk.Label(root, text="สุ่มผู้โชคดีรับรางวัลฉลองวันปีใหม่",
                 font=("Helvetica", 24))
label.pack(padx=20, pady=10)
label.config(foreground="blue", background="white")

# รูปแบบองค์ประกอบ
frame = tk.Frame(root, borderwidth=2, relief="groove")
frame.pack(padx=20, pady=10)

# รูปแบบทั่วไป
root.geometry("500x500")
root.configure(bg="white")

# อ่านไฟล์ Excel
employee_df = read_excel_file()
employee_names = list(employee_df["ชื่อพนักงาน"])

# ปุ่ม 1
browse_button = tk.Button(
    root, text="เลือกไฟล์ Excel", command=add_new_excel_file)
browse_button.config(font=("Helvetica", 12))
browse_button.pack(padx=20, pady=10)
browse_button.config(width=15, height=1)


# ปุ่ม 2
generate_button = tk.Button(root, text="สุ่มรหัส",
                            command=generate_random_code)
generate_button.config(foreground="white", background="green",
                       font=("Helvetica", 12))
generate_button.pack(padx=20, pady=2)
generate_button.config(width=15, height=1)

# ผลลัพธ์การสุ่ม 
result_label = tk.Label(root, text="", font=(
    "Helvetica", 20), background='white')
result_label.pack(padx=20, pady=20)

remaining_label = tk.Label(
    root, text=f"รายชื่อที่เหลือ: {len(employee_names)} คน", font=("Helvetica", 14))
remaining_label.pack(padx=20, pady=10)

root.mainloop()

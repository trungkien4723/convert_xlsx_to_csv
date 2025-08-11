import pandas as pd
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Hàm chuyển XLSX -> CSV theo mẫu
def convert_xlsx_to_csv():
    try:
        xlsx_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not xlsx_path:
            return
        
        csv_template_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not csv_template_path:
            return

        # Đọc file mẫu CSV để lấy header
        df_template = pd.read_csv(csv_template_path, nrows=0)
        headers = df_template.columns.tolist()

        # Đọc dữ liệu từ XLSX
        df_xlsx = pd.read_excel(xlsx_path, dtype=str).fillna("")

        # Tạo DataFrame mới theo mẫu
        df_new = pd.DataFrame(columns=headers)
        for col in headers:
            if col in df_xlsx.columns:
                df_new[col] = df_xlsx[col]
            else:
                df_new[col] = ""

        # Chọn nơi lưu CSV
        save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")])
        if not save_path:
            return

        # Xuất CSV với UTF-8 + BOM và bao toàn bộ giá trị trong dấu nháy
        df_new.to_csv(save_path, index=False, quoting=csv.QUOTE_ALL, encoding="utf-8-sig")

        messagebox.showinfo("Hoàn tất", f"Đã lưu CSV tại:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))


# Hàm chuyển CSV -> XLSX
def convert_csv_to_xlsx():
    try:
        csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not csv_path:
            return

        df_csv = pd.read_csv(csv_path, dtype=str).fillna("")

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            return

        df_csv.to_excel(save_path, index=False)

        messagebox.showinfo("Hoàn tất", f"Đã lưu XLSX tại:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))


# Giao diện Tkinter
root = tk.Tk()
root.title("Excel - CSV Converter")
root.geometry("500x300")  # rộng hơn để không che nút

tk.Label(root, text="Chuyển đổi định dạng giữa XLSX và CSV", font=("Arial", 14)).pack(pady=20)

btn1 = tk.Button(root, text="XLSX ➜ CSV (theo mẫu)", font=("Arial", 12), command=convert_xlsx_to_csv)
btn1.pack(pady=10, ipadx=10, ipady=5, fill="x", padx=50)

btn2 = tk.Button(root, text="CSV ➜ XLSX", font=("Arial", 12), command=convert_csv_to_xlsx)
btn2.pack(pady=10, ipadx=10, ipady=5, fill="x", padx=50)

tk.Button(root, text="Thoát", font=("Arial", 12), command=root.quit).pack(pady=20, ipadx=10, ipady=5, fill="x", padx=50)

root.mainloop()

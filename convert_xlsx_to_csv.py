import pandas as pd
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

# Bật DPI Awareness trên Windows để UI hiển thị đủ kích cỡ trên màn hình 4K/scale
try:
    if sys.platform.startswith("win"):
        try:
            from ctypes import windll
            # Windows 8.1+ (SetProcessDpiAwareness)
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                # Fallback (Vista+)
                windll.user32.SetProcessDPIAware()
            except Exception:
                pass
except Exception:
    # Không chặn chương trình nếu DPI không set được
    pass

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
        # Chuẩn hoá header mẫu: trim khoảng trắng, giữ nguyên thứ tự hiển thị
        template_headers = [str(h).strip() for h in df_template.columns.tolist()]

        # Đọc dữ liệu từ XLSX
        df_xlsx = pd.read_excel(xlsx_path, dtype=str).fillna("")
        # Tạo map header đã chuẩn hoá (case-insensitive) -> tên cột gốc trong XLSX
        xlsx_header_map = {}
        for original_col_name in df_xlsx.columns:
            normalized = str(original_col_name).strip().lower()
            if normalized not in xlsx_header_map:
                xlsx_header_map[normalized] = original_col_name

        # Tạo DataFrame mới theo mẫu
        df_new = pd.DataFrame(columns=template_headers)
        # Gán dữ liệu theo tên cột (ưu tiên khớp tên sau khi chuẩn hoá), nếu không có thì fallback theo vị trí
        for idx, header in enumerate(template_headers):
            norm = str(header).strip().lower()
            if norm in xlsx_header_map:
                df_new[header] = df_xlsx[xlsx_header_map[norm]]
            elif idx < df_xlsx.shape[1]:
                # Fallback: lấy theo vị trí nếu số cột đủ
                df_new[header] = df_xlsx.iloc[:, idx]
            else:
                # Không có dữ liệu tương ứng -> cột rỗng
                df_new[header] = ""

        # Chọn nơi lưu CSV
        save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv")])
        if not save_path:
            return

        # Xuất CSV với UTF-8 + BOM và bao toàn bộ giá trị trong dấu nháy
        df_new.to_csv(
            save_path,
            index=False,
            quoting=csv.QUOTE_ALL,
            encoding="utf-8-sig",
            lineterminator="\n",
        )

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
# Cho phép thay đổi kích thước cửa sổ và đặt kích thước tối thiểu để không che nút
root.geometry("640x360")
root.minsize(600, 340)
root.resizable(True, True)

tk.Label(root, text="Chuyển đổi định dạng giữa XLSX và CSV", font=("Arial", 14)).pack(pady=20)

btn1 = tk.Button(root, text="XLSX ➜ CSV (theo mẫu)", font=("Arial", 12), command=convert_xlsx_to_csv)
btn1.pack(pady=10, ipadx=10, ipady=5, fill="x", padx=50)

btn2 = tk.Button(root, text="CSV ➜ XLSX", font=("Arial", 12), command=convert_csv_to_xlsx)
btn2.pack(pady=10, ipadx=10, ipady=5, fill="x", padx=50)

tk.Button(root, text="Thoát", font=("Arial", 12), command=root.quit).pack(pady=20, ipadx=10, ipady=5, fill="x", padx=50)

root.mainloop()

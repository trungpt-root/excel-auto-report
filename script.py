import pandas as pd
import glob
import os

print("=== BẮT ĐẦU ===")

files = glob.glob("data/*.xlsx")

if not files:
    raise Exception("KHÔNG có file Excel nào trong thư mục data")

print(f"Có {len(files)} file")

df_list = []

for file in files:
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        df_list.append(df)
        print(f"OK: {file}")
    except Exception as e:
        print(f"LỖI: {file} - {e}")

merged = pd.concat(df_list, ignore_index=True)

print("Đã gộp xong")

# ====== SỬA PHẦN NÀY THEO FILE CỦA BẠN ======
# Ví dụ chuẩn

report = merged.groupby("Tên nhân viên")["Doanh số"].sum().reset_index()

# ===========================================

if os.path.exists("report.xlsx"):
    os.remove("report.xlsx")

report.to_excel("report.xlsx", index=False)

print("✅ DONE")

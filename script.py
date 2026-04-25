import pandas as pd
import glob
import os

print("=== AUTO PROCESS EXCEL ===")

files = glob.glob("data/*.xlsx")

if not files:
    raise Exception("Không có file Excel nào!")

print(f"Tìm thấy {len(files)} file")

df_list = []

for file in files:
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        df["__source_file"] = os.path.basename(file)  # thêm nguồn file
        df_list.append(df)
        print(f"OK: {file}")
    except Exception as e:
        print(f"LỖI: {file} - {e}")

merged = pd.concat(df_list, ignore_index=True)

print("Đã gộp xong")

# ===== AUTO DETECT =====

# tìm cột dạng số
numeric_cols = merged.select_dtypes(include="number").columns.tolist()

# tìm cột dạng text
text_cols = merged.select_dtypes(include="object").columns.tolist()

print("Cột số:", numeric_cols)
print("Cột text:", text_cols)

# ===== LOGIC THÔNG MINH =====

if numeric_cols and text_cols:
    try:
        # lấy cột text đầu tiên làm group
        group_col = text_cols[0]
        
        # lấy cột số đầu tiên để sum
        value_col = numeric_cols[0]

        print(f"Auto group theo: {group_col}")
        print(f"Auto sum: {value_col}")

        report = merged.groupby(group_col)[value_col].sum().reset_index()

    except:
        print("Không group được → xuất raw")
        report = merged

else:
    print("Không đủ dữ liệu để group → xuất raw")
    report = merged

# ======================

if os.path.exists("report.xlsx"):
    os.remove("report.xlsx")

report.to_excel("report.xlsx", index=False)

print("✅ DONE")

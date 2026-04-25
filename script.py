import pandas as pd
import glob
import os

print("=== CLEAN & REPORT SYSTEM ===")

files = glob.glob("data/*.xlsx")

if not files:
    raise Exception("❌ Không có file Excel trong thư mục data")

df_list = []

for file in files:
    try:
        df = pd.read_excel(file)

        # ===== CLEAN DATA =====
        required_cols = [
            "No.",
            "Vietnamese Description",
            "Item Category Code",
            "Inventory",
            "Base Unit of Measure"
        ]

        # kiểm tra cột
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            print(f"⚠️ File {file} thiếu cột: {missing} → bỏ qua")
            continue

        df = df[required_cols]

        df.columns = [
            "Mã hàng",
            "Tên hàng",
            "Loại hàng",
            "Tồn kho",
            "Đơn vị"
        ]

        # convert số cho chắc chắn
        df["Tồn kho"] = pd.to_numeric(df["Tồn kho"], errors="coerce").fillna(0)

        df["Nguồn file"] = os.path.basename(file)

        df_list.append(df)

        print(f"✅ OK: {file}")

    except Exception as e:
        print(f"❌ Lỗi file {file}: {e}")

# kiểm tra có data không
if not df_list:
    raise Exception("❌ Không có dữ liệu hợp lệ")

# ===== GỘP =====
merged = pd.concat(df_list, ignore_index=True)

# ===== BÁO CÁO 1: TỒN KHO THEO LOẠI =====
report_category = merged.groupby("Loại hàng")["Tồn kho"].sum().reset_index()

# lọc bỏ tồn kho = 0
report_category = report_category[report_category["Tồn kho"] > 0]

# format số đẹp
report_category["Tồn kho"] = report_category["Tồn kho"].round(0).astype(int)

# ===== BÁO CÁO 2: TOP HÀNG =====
report_top = merged.sort_values(by="Tồn kho", ascending=False).head(20)
report_top["Tồn kho"] = report_top["Tồn kho"].round(0).astype(int)

# ===== SAVE =====
os.makedirs("output", exist_ok=True)

merged.to_excel("output/data_clean.xlsx", index=False)
report_category.to_excel("output/report_by_category.xlsx", index=False)
report_top.to_excel("output/top_inventory.xlsx", index=False)

print("✅ Đã xuất Excel")

# ===== DASHBOARD HTML (ĐẸP) =====

html_category = report_category.to_html(index=False)
html_top = report_top.to_html(index=False)

html_page = f"""
<html>
<head>
<meta charset="UTF-8">
<title>Inventory Dashboard</title>

<style>
body {{
    font-family: Arial;
    padding: 30px;
    background: #f5f7fa;
}}

h1 {{
    text-align: center;
}}

table {{
    border-collapse: collapse;
    width: 80%;
    margin: 20px auto;
    background: white;
}}

th {{
    background: #4CAF50;
    color: white;
    padding: 10px;
}}

td {{
    padding: 8px;
    text-align: center;
}}

tr:nth-child(even) {{
    background: #f2f2f2;
}}

.section {{
    margin-bottom: 50px;
}}
</style>
</head>

<body>

<h1>📊 INVENTORY DASHBOARD</h1>

<div class="section">
<h2>📦 Tồn kho theo loại</h2>
{html_category}
</div>

<div class="section">
<h2>🏆 Top 20 tồn kho</h2>
{html_top}
</div>

</body>
</html>
"""

with open("output/dashboard.html", "w", encoding="utf-8") as f:
    f.write(html_page)

print("✅ Dashboard ready")

import pandas as pd
import glob
import os

print("=== CLEAN & REPORT SYSTEM ===")

files = glob.glob("data/*.xlsx")

df_list = []

for file in files:
    df = pd.read_excel(file)

    # ===== CLEAN DATA =====
    df = df[[
        "No.",
        "Vietnamese Description",
        "Item Category Code",
        "Inventory",
        "Base Unit of Measure"
    ]]

    df.columns = [
        "Mã hàng",
        "Tên hàng",
        "Loại hàng",
        "Tồn kho",
        "Đơn vị"
    ]

    df["Nguồn file"] = os.path.basename(file)

    df_list.append(df)

# Gộp
merged = pd.concat(df_list, ignore_index=True)

# ===== BÁO CÁO 1: TỒN KHO THEO LOẠI =====
report_category = merged.groupby("Loại hàng")["Tồn kho"].sum().reset_index()

# ===== BÁO CÁO 2: TOP HÀNG TỒN KHO =====
report_top = merged.sort_values(by="Tồn kho", ascending=False).head(20)

# ===== SAVE =====
os.makedirs("output", exist_ok=True)

merged.to_excel("output/data_clean.xlsx", index=False)
report_category.to_excel("output/report_by_category.xlsx", index=False)
report_top.to_excel("output/top_inventory.xlsx", index=False)

print("✅ Đã xuất Excel")

# ===== DASHBOARD HTML =====

html = f"""
<h2>📊 Tồn kho theo loại</h2>
{report_category.to_html(index=False)}

<h2>🏆 Top 20 tồn kho</h2>
{report_top.to_html(index=False)}
"""

html_page = f"""
<html>
<head>
<style>
body {{ font-family: Arial; padding: 20px; }}
table {{ border-collapse: collapse; width: 100%; margin-bottom: 40px; }}
th, td {{ border: 1px solid #ddd; padding: 8px; }}
th {{ background-color: #f2f2f2; }}
</style>
</head>
<body>
<h1>📊 INVENTORY DASHBOARD</h1>
{html}
</body>
</html>
"""

with open("output/dashboard.html", "w", encoding="utf-8") as f:
    f.write(html_page)

print("✅ Dashboard ready")

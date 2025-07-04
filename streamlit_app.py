import streamlit as st
import pandas as pd
import math
from pathlib import Path
import plotly.express as px
import os
os.system("pip install openpyxl")

# Set the title and favicon that appear in the Browser's tab bar.
st.set_page_config(
    page_title='KPI dashboard Tool',
    page_icon=':earth_americas:',
)
import streamlit as st
import pandas as pd
import re
import os
os.system("pip install openpyxl")

st.set_page_config(page_title="Đọc tên nhân viên", page_icon="📊")

# =====================
# Hàm chuẩn hóa tên nhân viên
def clean_employee_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"\n.*", "", name)  # Xoá phần sau xuống dòng nếu có
    name = re.sub(r"\(.*?\)", "", name)  # Xoá ghi chú trong ngoặc ()
    name = re.sub(r"\s+", " ", name)  # Chuẩn hoá khoảng trắng
    return name.strip()

# =====================
# Hàm trích xuất từng sheet
def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    current_nv = None
    rows = sheet_df.shape[0]

    i = 3  # bắt đầu từ dòng 4
    while i < rows:
        row = sheet_df.iloc[i]
        name_cell = str(row[1]).strip() if pd.notna(row[1]) else ""

        if name_cell and name_cell.lower() not in ["nan", "组员名字", "表格不要做任何调整，除前两列，其余全是公式"]:
            current_nv = name_cell
            for j in range(i, i + 6):
                if j >= rows:
                    break
                sub_row = sheet_df.iloc[j]
                nguon = sub_row[2]
                if pd.isna(nguon) or str(nguon).strip() in ["", "0"]:
                    break
                data.append({
                    "Nhân viên": current_nv.strip(),
                    "Nguồn": str(nguon).strip(),
                    "Sheet": sheet_name
                })
            i += 6
        else:
            i += 1
    return data

# =====================
# Hàm đọc toàn bộ file Excel
def extract_all_data(file):
    xls = pd.ExcelFile(file)
    all_rows = []

    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            if df.shape[0] < 10 or df.shape[1] < 5:
                continue
            extracted = extract_data_from_sheet(df, sheet_name)
            all_rows.extend(extracted)
        except Exception as e:
            st.warning(f"❌ Lỗi ở sheet '{sheet_name}': {e}")

    return pd.DataFrame(all_rows)

# =====================
# Upload file
uploaded_files = st.file_uploader("📥 Kéo nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        st.write(f"📂 Đang xử lý: `{file.name}`")
        df = extract_all_data(file)
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    # Chuẩn hoá tên nhân viên
    df_all["Nhân viên chuẩn"] = df_all["Nhân viên"].apply(clean_employee_name)

    st.subheader("✅ Danh sách nhân viên đã chuẩn hóa")
    st.dataframe(df_all[["Nhân viên", "Nhân viên chuẩn", "Sheet"]].drop_duplicates(), use_container_width=True)

    st.success(f"Tổng số dòng dữ liệu: {len(df_all)} — Nhân viên duy nhất: {df_all['Nhân viên chuẩn'].nunique()}")

else:
    st.info("📎 Vui lòng upload file Excel để bắt đầu.")

import streamlit as st
import pandas as pd
import math
from pathlib import Path
import plotly.express as px
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

def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    current_nv = None
    rows = sheet_df.shape[0]
    i = 3  # bỏ qua 3 dòng đầu

    while i < rows:
        row = sheet_df.iloc[i]
        # Nếu có tên mới thì cập nhật current_nv
        if pd.notna(row[1]) and str(row[1]).strip().lower() not in ["", "nan", "组员名字", "表格不要 làm gì"]:
            current_nv = re.sub(r"\(.*?\)", "", str(row[1])).strip()

        empty_count = 0
        j = i
        while j < rows:
            sub_row = sheet_df.iloc[j]

            # Nếu có tên nhân viên mới ở dòng này, cập nhật lại current_nv
            if pd.notna(sub_row[1]) and str(sub_row[1]).strip().lower() not in ["", "nan", "组员名字", "表格不要 làm gì"]:
                current_nv = re.sub(r"\(.*?\)", "", str(sub_row[1])).strip()

            nguon = str(sub_row[2]).strip() if pd.notna(sub_row[2]) else ""

            if nguon == "" or nguon.lower() == "nan":
                empty_count += 1
                if empty_count >= 2:
                    break
            else:
                empty_count = 0
                data.append({
                    "Nhân viên": current_nv,
                    "Nguồn": nguon,
                    "Tương tác ≥10 câu": pd.to_numeric(sub_row[15], errors="coerce"),
                    "Group Zalo": pd.to_numeric(sub_row[18], errors="coerce"),
                    "Kết bạn trong ngày": pd.to_numeric(sub_row[12], errors="coerce"),
                    "Sheet": sheet_name
                })
            j += 1
        i = j

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

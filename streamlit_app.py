# 🔄 Code cập nhật: thêm cột 'Xuất hiện ở các sheet'

import unicodedata
import re
import pandas as pd
import streamlit as st
from collections import defaultdict

# ✅ Hàm chuẩn hóa text
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = text.strip().lower()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = re.sub(r'\s+', ' ', text)
    return text

# ✅ Chuẩn hóa tên nhân viên
def normalize_name(name):
    if not isinstance(name, str):
        return ""
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r'\s+', ' ', name).strip().title()
    return name

# ✅ Lấy tên nhân viên theo block merge
def extract_data_with_staff(df, staff_col_index=1):
    df = df.copy()
    df = df.dropna(how='all')
    df.columns = [f"col_{i}" for i in range(len(df.columns))]
    staff_col = f"col_{staff_col_index}"

    current_name = ""
    empty_count = 0
    stop_index = None

    for i, val in enumerate(df[staff_col]):
        val = str(val).strip()
        if val:
            current_name = val
            df.at[i, staff_col] = current_name
            empty_count = 0
        else:
            df.at[i, staff_col] = current_name
            empty_count += 1

        if empty_count >= 2:
            stop_index = i
            break

    if stop_index:
        df = df.iloc[:stop_index]

    df[staff_col] = df[staff_col].apply(normalize_name)
    df.rename(columns={staff_col: "Tên nhân viên"}, inplace=True)
    return df

# ✅ Giao diện Streamlit
st.set_page_config(page_title="📊 KPI Dashboard", layout="wide")
st.title("📋 Danh sách Nhân Viên từ File Excel")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    name_to_sheets = defaultdict(set)

    for file in uploaded_files:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            try:
                raw_df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
                df = extract_data_with_staff(raw_df, staff_col_index=1)
                st.caption(f"📄 Sheet: `{sheet_name}` — Cột: {list(df.columns)}")

                for name in df['Tên nhân viên']:
                    if name and normalize_text(name) not in ["nan", "", "zuoyuan", "zuoyuan mingzi"]:
                        name_to_sheets[normalize_name(name)].add(sheet_name)

            except Exception as e:
                st.warning(f"❗ Sheet {sheet_name} lỗi: {e}")

    # ======= Hiển thị bảng tổng hợp
    if name_to_sheets:
        data = []
        for name, sheets in name_to_sheets.items():
            data.append({
                "Tên nhân viên chuẩn hóa": name,
                "Xuất hiện ở các sheet": ", ".join(sorted(sheets)),
                "Số lần xuất hiện": len(sheets)
            })
        df_result = pd.DataFrame(data).sort_values("Tên nhân viên chuẩn hóa")

        st.dataframe(df_result, use_container_width=True)
        st.success(f"✅ Tổng cộng có {len(df_result)} nhân viên duy nhất sau chuẩn hóa.")

        st.download_button("📥 Tải danh sách nhân viên", data=df_result.to_csv(index=False).encode('utf-8-sig'), file_name="danh_sach_nhan_vien.csv", mime="text/csv")
    else:
        st.error("❌ Không tìm được nhân viên nào.")

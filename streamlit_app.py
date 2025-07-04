import streamlit as st
import unicodedata
import re
import pandas as pd


# ✅ Chuẩn hóa tên nhân viên
def normalize_name(name):
    if not isinstance(name, str):
        return ""
    name = name.strip()
    name = re.sub(r'\s+', ' ', name)
    name = name.title()
    return name

# ✅ Hàm lọc tên nhân viên từ cột B, dừng khi gặp 2 dòng trống liên tiếp
def extract_staff_names(df, staff_col_index=1):
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
    return df[["Tên nhân viên"]]

# ✅ Giao diện Streamlit
st.set_page_config(page_title="📋 Danh sách Nhân viên", layout="wide")
st.title("📋 Upload & Chuẩn hóa Tên Nhân Viên từ File Excel")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_names = []
    for file in uploaded_files:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            try:
                raw_df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
                df = extract_staff_names(raw_df, staff_col_index=1)
                all_names.append(df)
            except Exception as e:
                st.warning(f"❗ Sheet {sheet_name} lỗi: {e}")

    if all_names:
        combined = pd.concat(all_names, ignore_index=True)
        unique_staff = combined.drop_duplicates().reset_index(drop=True)
        st.success(f"✅ Có tổng cộng {len(unique_staff)} nhân viên khác nhau")
        st.dataframe(unique_staff, use_container_width=True)
        st.download_button("📥 Tải danh sách nhân viên", data=unique_staff.to_csv(index=False).encode('utf-8-sig'), file_name="danh_sach_nhan_vien.csv", mime="text/csv")
    else:
        st.error("❌ Không có dữ liệu nhân viên nào được trích xuất.")

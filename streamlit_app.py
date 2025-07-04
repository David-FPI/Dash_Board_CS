import streamlit as st
import pandas as pd
import re
import os

os.system("pip install openpyxl")

st.set_page_config(page_title="📊 Đọc tên nhân viên", page_icon="👩‍💼")

# =====================
# 🔧 Hàm chuẩn hóa tên nhân viên
def clean_employee_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"\n.*", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip().title()


# =====================
# 📥 Đọc từng sheet
def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    rows = sheet_df.shape[0]

    sheet_df[1] = sheet_df[1].fillna(method='ffill')  # fill tên nhân viên từ merge
    current_nv = None
    empty_count = 0

    for i in range(3, rows):  # bỏ 3 dòng đầu
        row = sheet_df.iloc[i]

        # Xác định tên nhân viên từ cột B
        if pd.notna(row[1]):
            name_cell = str(row[1]).strip()
            if name_cell.lower() in ["组员名字", "统计", "表格不要 làm gì", "tổng"]:
                continue
            current_nv = re.sub(r"\(.*?\)", "", name_cell).strip()

        # Nếu không có tên thì bỏ qua
        if not current_nv:
            continue

        # Xác định nguồn từ cột C
        nguon = str(row[2]).strip() if pd.notna(row[2]) else ""
        if nguon == "" or nguon.lower() == "nan":
            empty_count += 1
            if empty_count >= 2:
                break  # kết thúc khối dữ liệu nếu trống liên tiếp 2 dòng
            continue
        else:
            empty_count = 0

        # Lưu lại dòng hợp lệ
        data.append({
            "Nhân viên": current_nv,
            "Nguồn": nguon,
            "Tương tác ≥10 câu": pd.to_numeric(row[15], errors="coerce"),
            "Group Zalo": pd.to_numeric(row[18], errors="coerce"),
            "Kết bạn trong ngày": pd.to_numeric(row[12], errors="coerce"),
            "Sheet": sheet_name
        })

    return data


# =====================
# 📤 Đọc toàn bộ file Excel
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
# Giao diện upload
st.title("📥 Đọc Tên Nhân Viên Từ File Excel Báo Cáo")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        st.write(f"📂 Đang xử lý: `{file.name}`")
        df = extract_all_data(file)
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    if not df_all.empty:
        # Chuẩn hóa tên nhân viên
        df_all["Nhân viên chuẩn"] = df_all["Nhân viên"].apply(clean_employee_name)

        st.subheader("✅ Danh sách Nhân viên đã chuẩn hóa")
        st.dataframe(df_all[["Nhân viên", "Nhân viên chuẩn", "Sheet"]].drop_duplicates(), use_container_width=True)

        st.success(f"✅ Tổng số dòng dữ liệu: {len(df_all)} — 👩‍💻 Nhân viên duy nhất: {df_all['Nhân viên chuẩn'].nunique()}")
    else:
        st.warning("❗ Không có dữ liệu nào được trích xuất. Vui lòng kiểm tra lại file.")
else:
    st.info("📎 Vui lòng upload file Excel để bắt đầu.")

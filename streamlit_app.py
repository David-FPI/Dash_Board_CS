import streamlit as st
import pandas as pd
import re
import os
from unidecode import unidecode

st.set_page_config(page_title="📊 Đọc tên nhân viên & Tính KPI", page_icon="👩‍💼")

# =====================
# 🔧 Chuẩn hóa tên nhân viên
def clean_employee_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"\n.*", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip().title()

# =====================
# 🔧 Tiên xử lý tiêu đề header (dò theo keyword linh hoạt)
def detect_kpi_columns(header_row):
    mapping = {}
    for i, col in enumerate(header_row):
        text = unidecode(str(col)).lower()
        text = re.sub(r"\s+", " ", text.replace("\n", " ").replace("\t", " ")).strip()

        if ">=10" in text:
            mapping["Tương tác ≥10 câu"] = i
        elif ("group" in text and "zalo" in text):
            mapping["Lượng tham gia group Zalo"] = i
        elif ("ket ban" in text and "trong ngay" in text):
            mapping["Tổng số kết bạn trong ngày"] = i

    return mapping

# =====================
# 📅 Đọc 1 sheet duy nhất
def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    sheet_df = sheet_df.drop([0, 1])  # bỏ dòng 1, 2
    header_row = sheet_df.iloc[0]
    sheet_df = sheet_df[1:].reset_index(drop=True)
    kpi_columns = detect_kpi_columns(header_row)

    if len(kpi_columns) < 3:
        st.warning(f"⚠️ Sheet {sheet_name} không đủ cột KPI — dò được {list(kpi_columns.keys())}")
        return []

    sheet_df[1] = sheet_df[1].fillna(method='ffill')
    current_nv = None
    empty_count = 0

    for _, row in sheet_df.iterrows():
        if pd.notna(row[1]):
            name_cell = str(row[1]).strip()
            if name_cell.lower() in ["组员名字", "统计", "表格不要做什么", "tổng"]:
                continue
            current_nv = re.sub(r"\(.*?\)", "", name_cell).strip()
        if not current_nv:
            continue

        nguon = str(row[2]).strip() if pd.notna(row[2]) else ""
        if nguon == "" or nguon.lower() == "nan":
            empty_count += 1
            if empty_count >= 2:
                break
            continue
        else:
            empty_count = 0

        data.append({
            "Nhân viên": current_nv,
            "Nguồn": nguon,
            "Tương tác ≥10 câu": row[kpi_columns["Tương tác ≥10 câu"]],
            "Lượng tham gia group Zalo": row[kpi_columns["Lượng tham gia group Zalo"]],
            "Tổng số kết bạn trong ngày": row[kpi_columns["Tổng số kết bạn trong ngày"]],
            "Sheet": sheet_name
        })

    return data

# =====================
# 📅 Xử lý nhiều file

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
            st.warning(f"❌ Lỗi sheet '{sheet_name}': {e}")

    return pd.DataFrame(all_rows)

# =====================
# 📁 Giao diện Streamlit
st.title("📥 Đọc Tên Nhân Viên & Tính KPI Từ File Excel Báo Cáo")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        st.write(f"📂 Đang xử lý: `{file.name}`")
        df = extract_all_data(file)
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    if not df_all.empty:
        df_all["Nhân viên chuẩn"] = df_all["Nhân viên"].apply(clean_employee_name)

        st.subheader("✅ Danh sách Nhân viên đã chuẩn hóa")
        st.dataframe(df_all[["Nhân viên", "Nhân viên chuẩn", "Sheet"]].drop_duplicates(), use_container_width=True)

        st.success(f"✅ Tổng số dòng dữ liệu: {len(df_all)} — 👩‍💼 Nhân viên duy nhất: {df_all['Nhân viên chuẩn'].nunique()}")
    else:
        st.warning("❗ Không có dữ liệu hợp lệ. Kiểm tra file.")
else:
    st.info("📎 Upload file Excel để bắt đầu.")

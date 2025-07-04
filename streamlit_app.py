import streamlit as st
import pandas as pd
import re
import os
from unidecode import unidecode

st.set_page_config(page_title="📅 Đọc Tên Nhân Viên & Tính KPI", page_icon="💼")

# =====================
# 🔧 Tự động cài package (nếu chưa có)
os.system("pip install openpyxl unidecode")

# =====================
# 🔧 Chuẩn hóa text để so sánh

def normalize_text(text):
    text = str(text).lower()
    text = re.sub(r"[\n\r]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = unidecode(text.strip())
    return text

# =====================
# 🖊️ Từ điển keyword để mapping cột
COLUMN_MAPPING_KEYWORDS = {
    "Tương tác ≥10 câu": ["10 cau", ">=10", "tuong tac", "so cau tuong tac"],
    "Lượng tham gia group Zalo": ["group zalo", "tham gia zalo", "nhom zalo", "zalo group", "join group", "zalo"],
    "Tổng số kết bạn trong ngày": ["ket ban", "tong so ket ban", "ket ban trong ngay", "add zalo"]
}

# =====================
# 📂 Trích xuất dữ liệu từ sheet

def extract_data_from_sheet(df, sheet_name):
    data = []
    rows = df.shape[0]

    df = df.iloc[2:].reset_index(drop=True)
    df.columns = [normalize_text(col) for col in df.iloc[0]]
    df = df[1:].reset_index(drop=True)

    col_mapping = {}
    for standard_name, keywords in COLUMN_MAPPING_KEYWORDS.items():
        for col in df.columns:
            for keyword in keywords:
                if keyword in col:
                    col_mapping[standard_name] = col
                    break
            if standard_name in col_mapping:
                break

    found_cols = list(col_mapping.keys())
    if len(found_cols) < 3:
        st.warning(f"⚠️ Sheet {sheet_name} không đủ cột KPI — dò được {found_cols}")
        return []

    if 1 in df.columns:
        df[1] = df[1].fillna(method='ffill')

    current_nv = None
    empty_count = 0
    for _, row in df.iterrows():
        if pd.notna(row[1]):
            name_cell = str(row[1]).strip()
            if name_cell.lower() in ["组员名字", "统计", "bảng tổng", "tổng"]:
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
            "Sheet": sheet_name,
            **{k: pd.to_numeric(row[col_mapping[k]], errors="coerce") for k in col_mapping}
        })

    return data

# =====================
# 📃 Xử lý toàn bộ file

def extract_all_data(file):
    xls = pd.ExcelFile(file)
    all_rows = []

    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            if df.shape[0] < 10:
                continue
            extracted = extract_data_from_sheet(df, sheet_name)
            all_rows.extend(extracted)
        except Exception as e:
            st.warning(f"❌ Lỗi sheet {sheet_name}: {e}")

    return pd.DataFrame(all_rows)

# =====================
# 📅 Giao diện Streamlit

st.title("📅 Đọc Tên Nhân Viên & Tính KPI Từ File Excel Báo Cáo")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        st.write(f"📂 Đang xử lý: `{file.name}`")
        df = extract_all_data(file)
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    if not df_all.empty:
        df_all["Nhân viên chuẩn"] = df_all["Nhân viên"].apply(lambda x: re.sub(r"\(.*?\)", "", str(x)).strip().title())

        st.subheader("✅ Danh sách Nhân viên đã chuẩn hóa")
        st.dataframe(df_all[["Nhân viên", "Nhân viên chuẩn", "Sheet"]].drop_duplicates(), use_container_width=True)

        tong_dong = len(df_all)
        so_nv = df_all["Nhân viên chuẩn"].nunique()
        st.success(f"✅ Tổng số dòng dữ liệu: {tong_dong} — 👩‍💼 Nhân viên duy nhất: {so_nv}")

        # ========== KPI Tuỳ Biến ==========
        st.header("📊 KPI Dashboard - Tính KPI Tuỳ Biến")
        st.subheader("🔢 Dữ liệu tổng hợp ban đầu")
        st.dataframe(df_all.head(), use_container_width=True)

        st.subheader("⚙️ Cấu hình KPI Tuỳ Biến")
        kpi_col1 = st.selectbox("Chọn cột A", df_all.columns)
        operator = st.selectbox("Phép toán", ["/", "*", "+", "-"])
        kpi_col2 = st.selectbox("Chọn cột B", df_all.columns)
        kpi_name = st.text_input("Tên chỉ số KPI mới", "Hiệu suất (%)")

        if st.button("✅ Tính KPI"):
            try:
                if operator == "/":
                    df_all[kpi_name] = df_all[kpi_col1] / df_all[kpi_col2] * 100
                elif operator == "*":
                    df_all[kpi_name] = df_all[kpi_col1] * df_all[kpi_col2]
                elif operator == "+":
                    df_all[kpi_name] = df_all[kpi_col1] + df_all[kpi_col2]
                elif operator == "-":
                    df_all[kpi_name] = df_all[kpi_col1] - df_all[kpi_col2]
                st.success(f"✅ KPI mới đã được tính: {kpi_name}")
                st.dataframe(df_all[["Nhân viên chuẩn", kpi_name, "Sheet"]], use_container_width=True)
            except Exception as e:
                st.error(f"⚠️ Lỗi khi tính KPI: {e}")
    else:
        st.warning("❗ Không có dữ liệu nào được trích xuất. Vui lòng kiểm tra file.")
else:
    st.info("📁 Vui lòng upload file Excel để bắt đầu.")

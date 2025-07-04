import streamlit as st
import pandas as pd
import re
from unidecode import unidecode

st.set_page_config(page_title="📊 Đọc tên nhân viên & Tính KPI", page_icon="👩‍💼")

# =====================
# 🔧 Các keyword linh hoạt để match các cột
COLUMN_KEYWORDS = {
    "Tương tác ≥10 câu": [">=10", "≥10", "tuong tac", "so tuong tac", "tương tác"],
    "Lượng tham gia group Zalo": ["group zalo", "tham gia group", "luong tham gia", "zalo nhom", "zalo group", "nhom zalo", "zalo", "tham gia zalo", "zalo tham gia", "zalo group join", "nhậu zalo", "加入zalo群数量"],
    "Tổng số kết bạn trong ngày": ["ket ban", "so ket ban", "tong ket ban", "tong so ket ban", "ket ban trong ngay", "zalo", "ket ban zalo", "ngay", "ketban", "当天加zalo"]
}

# =====================
# 🔧 Chuẩn hóa text để match header

def normalize_text(text):
    text = str(text).replace("\n", " ").replace("\r", " ")
    text = unidecode(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()

# =====================
# 🔍 Tìm vị trí cột theo keyword

def match_column_indices(header_row):
    mapping = {}
    for idx, col in enumerate(header_row):
        col_clean = normalize_text(col)
        for target_name, keyword_list in COLUMN_KEYWORDS.items():
            if any(kw in col_clean for kw in keyword_list):
                mapping[target_name] = idx
    return mapping

# =====================
# 📅 Đọc dữ liệu từ sheet

def extract_data_from_sheet(df, sheet_name):
    data = []
    if df.shape[0] < 5:
        return data, []

    df.columns = range(df.shape[1])  # reset column index
    df = df.drop([0, 1])  # Bỏ 2 dòng đầu
    header = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    col_map = match_column_indices(header)

    if len(col_map) < 2:
        return [], list(col_map.keys())

    df[1] = df[1].fillna(method="ffill")
    current_nv = None
    empty_count = 0

    for i in range(df.shape[0]):
        row = df.iloc[i]

        if pd.notna(row[1]):
            name_cell = str(row[1]).strip()
            if name_cell.lower() in ["nhan vien", "tong", "stat"]:
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
            "Tương tác ≥10 câu": row.get(col_map.get("Tương tác ≥10 câu")),
            "Lượng tham gia group Zalo": row.get(col_map.get("Lượng tham gia group Zalo")),
            "Tổng số kết bạn trong ngày": row.get(col_map.get("Tổng số kết bạn trong ngày")),
            "Sheet": sheet_name
        })
    return data, list(col_map.keys())

# =====================
# 📂 Đọc toàn bộ file Excel

def extract_all_data(file):
    xls = pd.ExcelFile(file)
    all_rows = []
    warnings = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        records, found_cols = extract_data_from_sheet(df, sheet)
        all_rows.extend(records)
        if len(found_cols) < 2:
            warnings.append((sheet, found_cols))
    return pd.DataFrame(all_rows), warnings

# =====================
# 🔍 App
st.title("📅 Đọc Tên Nhân Viên & Tính KPI Từ File Excel Báo Cáo")
uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    all_warnings = []
    for file in uploaded_files:
        st.write(f"📂 Đang xử lý: `{file.name}`")
        df, warns = extract_all_data(file)
        all_data.append(df)
        all_warnings.extend(warns)

    df_all = pd.concat(all_data, ignore_index=True)

    if not df_all.empty:
        df_all["Nhân viên chuẩn"] = df_all["Nhân viên"].apply(lambda x: str(x).strip().title())

        st.subheader("✅ Danh sách Nhân viên đã chuẩn hóa")
        st.dataframe(df_all[["Nhân viên", "Nhân viên chuẩn", "Sheet"]].drop_duplicates(), use_container_width=True)

        st.success(f"✅ Tổng số dòng dữ liệu: {len(df_all)} — 👩‍💼 Nhân viên duy nhất: {df_all['Nhân viên chuẩn'].nunique()}")

        # Cảnh báo sheet bị thiếu KPI
        for sheet, found in all_warnings:
            st.warning(f"⚠️ Sheet {sheet} không đủ cột KPI — dò được {found}")
    else:
        st.error("❌ Không có dữ liệu hợp lệ. Kiểm tra file.")

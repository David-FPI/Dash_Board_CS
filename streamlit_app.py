import unicodedata
import re
import pandas as pd
import streamlit as st

# ✅ Hàm chuẩn hóa text: bỏ dấu, lowercase, bỏ khoảng trắng thừa
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = text.strip().lower()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = re.sub(r'\s+', ' ', text)
    return text

# ✅ Danh sách keyword cho cột "Tương tác ≥10 câu"
KEYWORDS_TUONG_TAC = [
    "≥10", ">=10", "10 câu", "tuong tac", "số lượng tương tác", "tương tác 10 câu", "tương tác",
    "互动", "số câu hỏi", "tương tác với khách", "≥10句", "互动次数"
]

# ✅ Danh sách keyword cho cột "Lượng tham gia group Zalo"
KEYWORDS_GROUP_ZALO = [
    "group zalo", "zalo group", "tham gia group", "tham gia zalo", "nhóm zalo", "zalo nhóm",
    "zalo tham gia", "加zalo群", "加入zalo群数量", "vào group zalo", "vào nhóm zalo"
]

# ✅ Hàm nhận diện cột theo từ khóa
def is_tuong_tac_column(col):
    normalized = normalize_text(col)
    return any(keyword.lower() in normalized for keyword in KEYWORDS_TUONG_TAC)

def is_group_zalo_column(col):
    normalized = normalize_text(col)
    return any(keyword.lower() in normalized for keyword in KEYWORDS_GROUP_ZALO)

# ✅ Hàm dò và gán nhãn KPI từ danh sách tiêu đề
def detect_kpi_columns(columns):
    result = {}
    for col in columns:
        if is_tuong_tac_column(col):
            result["Tương tác ≥10 câu"] = col
        elif is_group_zalo_column(col):
            result["Lượng tham gia group Zalo"] = col
    return result

# ✅ Hàm chuẩn hóa tên nhân viên
def normalize_name(name):
    if not isinstance(name, str):
        return ""
    name = name.strip()
    name = re.sub(r'\s+', ' ', name)
    name = name.title()
    return name

# ✅ Tổng hợp dữ liệu KPI từ nhiều sheet
def summarize_kpi_across_sheets(sheet_data_list):
    all_data = []

    for sheet_data in sheet_data_list:
        df = sheet_data['data']
        kpi_columns = sheet_data['kpi_columns']
        staff_column = sheet_data['staff_column']

        if not staff_column or not kpi_columns:
            continue

        # Chỉ lấy 2 KPI quan trọng
        selected_kpi = {
            "Tương tác ≥10 câu": kpi_columns.get("Tương tác ≥10 câu"),
            "Lượng tham gia group Zalo": kpi_columns.get("Lượng tham gia group Zalo")
        }
        selected_kpi = {k: v for k, v in selected_kpi.items() if v}

        if not selected_kpi:
            continue

        columns_to_keep = [staff_column] + list(selected_kpi.values())
        df_filtered = df[columns_to_keep].copy()

        # Chuẩn hóa tên nhân viên
        df_filtered[staff_column] = df_filtered[staff_column].apply(normalize_name)

        # Đổi tên cột KPI theo chuẩn
        df_filtered = df_filtered.rename(columns={v: k for k, v in selected_kpi.items()})

        all_data.append(df_filtered)

    if not all_data:
        return pd.DataFrame()

    combined_df = pd.concat(all_data, ignore_index=True)
    summary = combined_df.groupby(staff_column, dropna=False).sum(numeric_only=True).reset_index()
    return summary

# ✅ Giao diện chạy trực tiếp bằng Streamlit
st.set_page_config(page_title="📊 KPI Dashboard", layout="wide")
st.title("📊 Dashboard KPI Nhân Viên từ File Excel")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    sheet_data_list = []
    for file in uploaded_files:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
                columns = df.columns.tolist()
                kpi_cols = detect_kpi_columns(columns)
                # Tìm cột nhân viên (ưu tiên cột có chữ "tên nhân viên" hoặc tên ở cột B)
                staff_col = df.columns[1] if len(df.columns) > 1 else None
                sheet_data_list.append({
                    'data': df,
                    'kpi_columns': kpi_cols,
                    'staff_column': staff_col
                })
            except Exception as e:
                st.warning(f"❗ Sheet {sheet_name} lỗi: {e}")

    result_df = summarize_kpi_across_sheets(sheet_data_list)
    if not result_df.empty:
        st.success("✅ Đã tổng hợp xong dữ liệu KPI")
        st.dataframe(result_df, use_container_width=True)
        st.download_button("📥 Tải về file tổng hợp", data=result_df.to_csv(index=False).encode('utf-8-sig'), file_name="kpi_tong_hop.csv", mime="text/csv")
    else:
        st.error("❌ Không có dữ liệu nào phù hợp để tổng hợp.")

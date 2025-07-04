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
    ">=10", "≥10", "10 cau", "tuong tac", "so luong tuong tac", "tuong tac 10 cau",
    "tuong tac voi khach", "so cau hoi", "互动", "互动次数", "≥10句"
]

# ✅ Danh sách keyword cho cột "Lượng tham gia group Zalo"
KEYWORDS_GROUP_ZALO = [
    "group zalo", "zalo group", "tham gia group", "tham gia zalo", "nhom zalo",
    "zalo nhom", "zalo tham gia", "vao group zalo", "vao nhom zalo",
    "加zalo群", "加入zalo群数量"
]

# ✅ Hàm nhận diện cột theo từ khóa
def is_tuong_tac_column(normalized_col):
    return any(keyword in normalized_col for keyword in KEYWORDS_TUONG_TAC)

def is_group_zalo_column(normalized_col):
    return any(keyword in normalized_col for keyword in KEYWORDS_GROUP_ZALO)

# ✅ Hàm dò và gán nhãn KPI từ danh sách tiêu đề
def detect_kpi_columns(columns):
    result = {}
    for col in columns:
        if not isinstance(col, str):
            continue
        norm = normalize_text(col)
        if is_tuong_tac_column(norm) and "Tương tác ≥10 câu" not in result:
            result["Tương tác ≥10 câu"] = col
        elif is_group_zalo_column(norm) and "Lượng tham gia group Zalo" not in result:
            result["Lượng tham gia group Zalo"] = col
    return result

# ✅ Chuẩn hóa tên nhân viên
def normalize_name(name):
    if not isinstance(name, str):
        return ""
    name = name.strip()
    name = re.sub(r'\s+', ' ', name)
    name = name.title()
    return name

# ✅ Hàm lọc tên nhân viên từ cột B, dừng khi gặp 2 dòng trống liên tiếp
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

    # Lọc bỏ tên không hợp lệ
    invalid_names = ["组员", "组员名字", ""]
    df = df[~df["Tên nhân viên"].isin(invalid_names)]

    return df

# ✅ Tổng hợp KPI

def summarize_kpi_across_sheets(sheet_data_list):
    all_data = []

    for sheet_data in sheet_data_list:
        df = sheet_data['data']
        kpi_columns = sheet_data['kpi_columns']

        if not kpi_columns:
            continue

        selected_kpi = {
            "Tương tác ≥10 câu": kpi_columns.get("Tương tác ≥10 câu"),
            "Lượng tham gia group Zalo": kpi_columns.get("Lượng tham gia group Zalo")
        }
        selected_kpi = {k: v for k, v in selected_kpi.items() if v}

        if not selected_kpi:
            continue

        columns_to_keep = ["Tên nhân viên"] + list(selected_kpi.values())
        df_filtered = df[columns_to_keep].copy()

        df_filtered = df_filtered.rename(columns={v: k for k, v in selected_kpi.items()})
        all_data.append(df_filtered)

    if not all_data:
        return pd.DataFrame()

    combined_df = pd.concat(all_data, ignore_index=True)
    kpi_fields = [col for col in ["Tương tác ≥10 câu", "Lượng tham gia group Zalo"] if col in combined_df.columns]
    summary = combined_df.groupby("Tên nhân viên", dropna=False)[kpi_fields].sum(numeric_only=True).reset_index()
    return summary

# ✅ Giao diện Streamlit
st.set_page_config(page_title="📊 KPI Dashboard", layout="wide")
st.title("📊 Dashboard KPI Nhân Viên từ File Excel")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    sheet_data_list = []
    for file in uploaded_files:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            try:
                raw_df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
                df = extract_data_with_staff(raw_df, staff_col_index=1)
                st.caption(f"📄 Sheet: `{sheet_name}` — Cột: {list(df.columns)}")
                columns = df.columns.tolist()
                kpi_cols = detect_kpi_columns(columns)

                sheet_data_list.append({
                    'data': df,
                    'kpi_columns': kpi_cols
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

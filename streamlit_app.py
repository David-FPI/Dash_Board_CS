import unicodedata
import re
import pandas as pd
import streamlit as st
from collections import defaultdict
from io import BytesIO

# ✅ Chuẩn hóa tên
def normalize_name(name):
    if not isinstance(name, str):
        return ""
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name).strip().title()
    return name

# ✅ Chuẩn hóa text để so sánh
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = text.strip().lower()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = re.sub(r'\s+', ' ', text)
    return text

# ✅ Tách dữ liệu nhân viên từ cột B (index = 1), kết thúc khi có 2 dòng trống liên tiếp
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

# ✅ Tạo bảng tổng hợp nhân viên + sheet xuất hiện
def build_staff_sheet_summary(sheet_data_list):
    staff_sheets = defaultdict(set)

    for item in sheet_data_list:
        df = item['data']
        sheet_name = item['sheet_name']
        for name in df["Tên nhân viên"]:
            if not name or normalize_text(name) in ["", "组员", "组员名字", "nan"]:
                continue
            staff_sheets[name].add(sheet_name)

    rows = []
    for idx, (name, sheets) in enumerate(sorted(staff_sheets.items()), start=1):
        sheet_list = sorted(list(sheets))
        rows.append({
            "STT": idx,
            "Tên nhân viên": name,
            "Xuất hiện ở các sheet": ", ".join(sheet_list),
            "Số lần xuất hiện": len(sheets)
        })

    return pd.DataFrame(rows)

# ✅ Chuyển DataFrame thành file tải về
def to_excel_download(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='DanhSachNhanVien')
    return buffer.getvalue()

# ✅ Hàm chuẩn hóa tiêu đề
def clean_text(text):
    if not isinstance(text, str):
        return ""
    text = text.strip().lower()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = re.sub(r'[\n\r\t]+', ' ', text)  # xóa xuống dòng/tab
    text = re.sub(r'\s+', ' ', text)
    return text

# ✅ Keyword liên quan đến "Tương tác ≥10 câu"
KEYWORDS_TUONG_TAC = [
    "≥10", "tuong tac", "10 cau", "≥10 cau", "≥10 câu",
    "trao doi", "interaction", "tuong tac (≥10 cau)",
    "tuong tac >=10"
]

def find_column_index_tuong_tac(file, sheet_name):
    """
    Đọc dòng thứ 3 (index=2) của sheet, dò các keyword tương tác
    Trả về: chỉ số cột nếu tìm thấy, None nếu không
    """
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=3)
        header_row3 = df_raw.iloc[2]  # dòng thứ 3 (index=2)
        for idx, val in enumerate(header_row3):
            col_clean = clean_text(str(val))
            for keyword in KEYWORDS_TUONG_TAC:
                if clean_text(keyword) in col_clean:
                    return idx
    except:
        return None
    return None

# ✅ Thống kê tổng tương tác ≥10 câu theo từng nhân viên
def summarize_interaction_by_staff(sheet_data_list):
    rows = []
    for item in sheet_data_list:
        df = item['data']
        if "Tương tác ≥10 câu" not in df.columns:
            continue

        for _, row in df.iterrows():
            name = row.get("Tên nhân viên", "")
            if not name or normalize_text(name) in ["", "组员", "组员名字", "nan"]:
                continue

            value = row.get("Tương tác ≥10 câu", 0)
            try:
                count = int(value)
            except:
                try:
                    count = float(str(value).replace(",", "."))
                except:
                    count = 0

            rows.append({
                "Tên nhân viên": name,
                "Số tương tác ≥10 câu": count
            })

    df_all = pd.DataFrame(rows)
    if df_all.empty:
        return pd.DataFrame()

    df_grouped = df_all.groupby("Tên nhân viên").sum().reset_index()
    df_grouped = df_grouped.sort_values(by="Số tương tác ≥10 câu", ascending=False).reset_index(drop=True)
    df_grouped.index += 1
    df_grouped.insert(0, "STT", df_grouped.index)
    return df_grouped



# ✅ Streamlit UI
st.set_page_config(page_title="📊 Danh sách Nhân Viên", layout="wide")
st.title("📋 Danh sách Nhân Viên từ File Excel")

uploaded_files = st.file_uploader("📁 Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    sheet_data_list = []
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        xls = pd.ExcelFile(uploaded_file)
        for sheet in xls.sheet_names:
            try:
                raw_df = pd.read_excel(xls, sheet_name=sheet, skiprows=2)
                df = extract_data_with_staff(raw_df, staff_col_index=1)
                                # ✅ Tìm cột tương tác ≥10 câu
                # ✅ Tìm cột tương tác bằng dòng 3 thật sự (không skip)
                col_index = find_column_index_tuong_tac(uploaded_file, sheet)
                if col_index is not None:
                    col_name = raw_df.columns[col_index]
                    raw_df["Tương tác ≥10 câu"] = raw_df[col_name]
                    st.info(f"📌 Sheet `{sheet}` có cột tương tác: `{col_name}`")
                else:
                    st.warning(f"⚠️ Sheet `{sheet}` không tìm thấy cột Tương tác ≥10 câu.")

                
                st.caption(f"📄 File: `{file_name}` — Sheet: `{sheet}` — {df.shape[0]} dòng")
                sheet_data_list.append({
                    'data': df,
                    'sheet_name': sheet


                })
            except Exception as e:
                st.warning(f"⚠️ Sheet `{sheet}` lỗi: {e}")

    df_summary = build_staff_sheet_summary(sheet_data_list)
    # ✅ Bảng tổng hợp tương tác ≥10 câu
    df_interaction = summarize_interaction_by_staff(sheet_data_list)

    if not df_interaction.empty:
        st.subheader("📈 Tổng số Tương tác ≥10 câu theo Nhân viên")
        st.dataframe(df_interaction, use_container_width=True)

        st.download_button(
            label="📥 Tải bảng Tương tác ≥10 câu",
            data=to_excel_download(df_interaction),
            file_name="tong_tuong_tac_10_cau.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if not df_summary.empty:
        st.success(f"✅ Tổng cộng có {df_summary.shape[0]} nhân viên duy nhất sau chuẩn hóa.")
        st.dataframe(df_summary, use_container_width=True)

        st.download_button(
            label="📥 Tải danh sách nhân viên",
            data=to_excel_download(df_summary),
            file_name="tong_hop_nhan_vien.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ Không tìm được nhân viên hợp lệ.")

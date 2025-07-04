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

# ✅ Chuẩn hóa tên nhân viên
def normalize_name(name):
    if not isinstance(name, str):
        return ""
    # Bỏ phần trong dấu ngoặc như (Event), (Note), v.v.
    name = re.sub(r"\(.*?\)", "", name)
    # Chuẩn hóa khoảng trắng và viết hoa chữ cái đầu
    name = re.sub(r'\s+', ' ', name).strip().title()
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
    return df

# ✅ Trích xuất danh sách tên nhân viên duy nhất
def extract_unique_staff_names(sheet_data_list):
    all_names = []

    for sheet_data in sheet_data_list:
        df = sheet_data['data']
        if "Tên nhân viên" not in df.columns:
            continue
        names = df["Tên nhân viên"].dropna().tolist()
        names = [normalize_name(name) for name in names if isinstance(name, str) and name.strip()]
        all_names.extend(names)

    # Lọc các tên không hợp lệ
    invalid_keywords = ["组员", "组员名字", "Nan", ""]
    all_names = [name for name in all_names if normalize_text(name) not in [normalize_text(x) for x in invalid_keywords]]

    # Trả ra tên unique và sắp xếp
    unique_names = sorted(set(all_names))
    return unique_names

# ✅ Giao diện Streamlit
st.set_page_config(page_title="📊 KPI Dashboard", layout="wide")
st.title("📋 Danh sách Nhân Viên từ File Excel")

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
                sheet_data_list.append({
                    'data': df
                })
            except Exception as e:
                st.warning(f"❗ Sheet {sheet_name} lỗi: {e}")

    unique_names = extract_unique_staff_names(sheet_data_list)

    if unique_names:
        df_names = pd.DataFrame({"Tên nhân viên chuẩn hóa": unique_names})
        st.dataframe(df_names, use_container_width=True)
        st.success(f"✅ Tổng cộng có {len(unique_names)} nhân viên duy nhất sau chuẩn hóa.")
        st.download_button("📥 Tải danh sách nhân viên", data=df_names.to_csv(index=False).encode('utf-8-sig'), file_name="danh_sach_nhan_vien.csv", mime="text/csv")
    else:
        st.error("❌ Không tìm được nhân viên nào.")

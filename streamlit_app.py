import streamlit as st
import pandas as pd
import re
import os
from unidecode import unidecode

os.system("pip install openpyxl")

st.set_page_config(page_title="📥 Đọc Nhân Viên & Tính KPI", page_icon="📊")

# ========== Hàm chuẩn hóa tiêu đề ==========
def clean_col_name(col):
    col = str(col)
    col = re.sub(r"\s+", " ", col.replace("\n", " "))  # bỏ xuống dòng và khoảng trắng
    col = unidecode(col).lower().strip()
    return col

# ========== Dò cột theo keyword ==========
def map_columns(cols):
    mapping = {}
    for i, col in enumerate(cols):
        col_clean = clean_col_name(col)
        if "≥10" in col_clean or ">=10" in col_clean:
            mapping["Tương tác ≥10 câu"] = i
        elif "group zalo" in col_clean or "zalo group" in col_clean:
            mapping["Lượng tham gia group Zalo"] = i
        elif "ket ban" in col_clean and "trong ngay" in col_clean:
            mapping["Tổng số kết bạn trong ngày"] = i
    return mapping

# ========== Chuẩn hóa tên nhân viên ==========
def clean_employee_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"\n.*", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip().title()

# ========== Đọc từng sheet ==========
def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    sheet_df = sheet_df.drop([0,1])  # Bỏ dòng 1 và 2
    sheet_df = sheet_df.reset_index(drop=True)

    header = sheet_df.iloc[0]
    sheet_df = sheet_df[1:]
    sheet_df.columns = header

    if sheet_df.shape[0] < 5:
        return []

    col_mapping = map_columns(sheet_df.columns)

    if len(col_mapping) < 3:
        st.warning(f"⚠️ Sheet `{sheet_name}` không đủ cột KPI — dò được {list(col_mapping.keys())}")
        return []

    sheet_df = sheet_df.reset_index(drop=True)
    sheet_df["NV"] = sheet_df.iloc[:,1].fillna(method="ffill")

    current_nv = None
    empty_count = 0

    for idx, row in sheet_df.iterrows():
        name_cell = str(row["NV"]).strip()
        if name_cell.lower() in ["组员名字", "统计", "表格不要 làm gì", "tổng"]:
            continue
        current_nv = clean_employee_name(name_cell)

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
            "Tương tác ≥10 câu": row[col_mapping["Tương tác ≥10 câu"]],
            "Lượng tham gia group Zalo": row[col_mapping["Lượng tham gia group Zalo"]],
            "Tổng số kết bạn trong ngày": row[col_mapping["Tổng số kết bạn trong ngày"]],
            "Sheet": sheet_name
        })

    return data

# ========== Đọc file ==========
def extract_all_data(file):
    xls = pd.ExcelFile(file)
    all_rows = []
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            extracted = extract_data_from_sheet(df, sheet_name)
            all_rows.extend(extracted)
        except Exception as e:
            st.warning(f"❌ Lỗi sheet {sheet_name}: {e}")
    return pd.DataFrame(all_rows)

# ========== Giao diện ==========
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
        st.success(f"✅ Tổng số dòng dữ liệu: {len(df_all)} — 👩‍💻 Nhân viên duy nhất: {df_all['Nhân viên chuẩn'].nunique()}")

        # KPI Dashboard
        st.markdown("### 📊 KPI Dashboard - Tính KPI Tùy Biến")
        st.markdown("#### 🔢 Dữ liệu tổng hợp ban đầu")
        st.dataframe(df_all.head(20), use_container_width=True)

        st.markdown("#### ⚙️ Cấu hình KPI Tuỳ Biến")
        col1, col2, col3 = st.columns(3)
        with col1:
            col_a = st.selectbox("Chọn cột A", df_all.columns[2:5])
        with col2:
            operation = st.selectbox("Phép toán", ["/", "*", "+", "-"])
        with col3:
            col_b = st.selectbox("Chọn cột B", df_all.columns[2:5])
        new_kpi = st.text_input("Tên chỉ số KPI mới", "Hiệu suất (%)")

        if st.button("✅ Tính KPI"):
            try:
                df_all[new_kpi] = eval(f"df_all['{col_a}'] {operation} df_all['{col_b}']")
                st.success(f"✅ Đã tính KPI mới: {new_kpi}")
                st.dataframe(df_all[[col_a, col_b, new_kpi, "Nhân viên chuẩn"]].head(20), use_container_width=True)
            except Exception as e:
                st.error(f"❌ Lỗi khi tính KPI: {e}")
    else:
        st.warning("❗ Không có dữ liệu nào hợp lệ.")
else:
    st.info("📎 Vui lòng upload file Excel để bắt đầu.")

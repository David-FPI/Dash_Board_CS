import streamlit as st
import pandas as pd
import re
import os

os.system("pip install openpyxl")

st.set_page_config(page_title="📊 Đọc tên nhân viên & Tính KPI", page_icon="👩‍💼")

# =====================
# 🔧 Hàm chuẩn hóa tên nhân viên
def clean_employee_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"\n.*", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip().title()

# =====================
# 📥 Dò cột từ dòng tiêu đề bằng keyword

def get_column_mapping(header_row):
    mapping = {}
    for idx, col in enumerate(header_row):
        col_clean = str(col).lower().replace("\n", " ").strip()
        if "\u226510" in col_clean or ">=10" in col_clean:
            mapping["Tương tác ≥10 câu"] = idx
        elif "group zalo" in col_clean:
            mapping["Lượng tham gia group Zalo"] = idx
        elif "kết bạn trong ngày" in col_clean:
            mapping["Tổng số kết bạn trong ngày"] = idx
    return mapping


# =====================
# 📥 Đọc từng sheet

def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    rows = sheet_df.shape[0]

    sheet_df = sheet_df.copy()
    sheet_df[1] = sheet_df[1].fillna(method='ffill')  # fill tên nhân viên từ merge

    if rows < 4:
        return data

    header_row = sheet_df.iloc[2]  # dùng dòng thứ 3 làm tiêu đề
    col_map = get_column_mapping(header_row)

    current_nv = None
    empty_count = 0

    for i in range(3, rows):  # bắt đầu từ dòng 4 trở đi
        row = sheet_df.iloc[i]

        # Xác định tên nhân viên từ cột B
        if pd.notna(row[1]):
            name_cell = str(row[1]).strip()
            if name_cell.lower() in ["组员名字", "统计", "表格不要 làm gì", "tổng"]:
                continue
            current_nv = re.sub(r"\(.*?\)", "", name_cell).strip()

        if not current_nv:
            continue

        # Xác định nguồn từ cột C
        nguon = str(row[2]).strip() if pd.notna(row[2]) else ""
        if nguon == "" or nguon.lower() == "nan":
            empty_count += 1
            if empty_count >= 2:
                break
            continue
        else:
            empty_count = 0

        data_row = {
            "Nhân viên": current_nv,
            "Nguồn": nguon,
            "Sheet": sheet_name
        }

        # Thêm các cột KPI nếu có
        for kpi_name, idx in col_map.items():
            value = pd.to_numeric(row[idx], errors="coerce")
            data_row[kpi_name] = value

        data.append(data_row)

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

        # 📊 KPI Dashboard - Tổng hợp và tính KPI
        st.subheader("📊 KPI Dashboard - Tính KPI Tùy Biến")

        st.markdown("### 🔢 Dữ liệu tổng hợp ban đầu")
        st.dataframe(df_all, use_container_width=True)

        st.markdown("### ⚙️ Cấu hình KPI Tuỳ Biến")

        kpi_cols = [col for col in df_all.columns if col not in ["Nhân viên", "Nguồn", "Sheet", "Nhân viên chuẩn"]]

        col_a = st.selectbox("Chọn cột A", kpi_cols)
        operation = st.selectbox("Phép toán", ["/", "*", "+", "-"])
        col_b = st.selectbox("Chọn cột B", kpi_cols)
        new_kpi_name = st.text_input("Tên chỉ số KPI mới", "Hiệu suất (%)")

        if st.button("✅ Tính KPI"):
            try:
                if operation == "/":
                    df_all[new_kpi_name] = df_all[col_a] / df_all[col_b]
                elif operation == "*":
                    df_all[new_kpi_name] = df_all[col_a] * df_all[col_b]
                elif operation == "+":
                    df_all[new_kpi_name] = df_all[col_a] + df_all[col_b]
                elif operation == "-":
                    df_all[new_kpi_name] = df_all[col_a] - df_all[col_b]

                st.success(f"✅ Đã tính KPI mới: {new_kpi_name}")
                st.dataframe(df_all[["Nhân viên chuẩn", col_a, col_b, new_kpi_name]], use_container_width=True)
            except Exception as e:
                st.error(f"❌ Lỗi khi tính KPI: {e}")

    else:
        st.warning("❗ Không có dữ liệu nào được trích xuất. Vui lòng kiểm tra lại file.")
else:
    st.info("📎 Vui lòng upload file Excel để bắt đầu.")

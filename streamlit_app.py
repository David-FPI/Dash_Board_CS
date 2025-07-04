import streamlit as st
import pandas as pd
import re
import os

os.system("pip install openpyxl")

st.set_page_config(page_title="📥 Đọc Tên Nhân Viên & Tính KPI", page_icon="👩‍💼")

# =====================
# 🔧 Hàm chuẩn hoá tên nhân viên
def clean_employee_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"\n.*", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip().title()

# =====================
# 🔧 Chuẩn hóa tiêu đề cột
def normalize_header(header):
    header = str(header).lower()
    header = re.sub(r"\s+", " ", header)  # Xoá khoảng trắng thừa & xuống dòng
    return header.strip()

# =====================
# 📥 Đọc từng sheet
def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    rows = sheet_df.shape[0]

    if rows < 3:
        return []

    # Xoá dòng 1 và 2 → Lấy dòng 3 làm header
    sheet_df.columns = sheet_df.iloc[2]
    df = sheet_df[3:].reset_index(drop=True)

    # Chuẩn hoá tiêu đề & dò vị trí cột
    header_map = {}
    for col in df.columns:
        col_clean = normalize_header(col)
        if "≥10" in col_clean:
            header_map["Tương tác ≥10 câu"] = col
        elif "group zalo" in col_clean:
            header_map["Lượng tham gia group Zalo"] = col
        elif "kết bạn trong ngày" in col_clean:
            header_map["Tổng số kết bạn trong ngày"] = col

    if len(header_map) == 0:
        return []

    # Fill tên nhân viên từ cột B (index 1)
    df.iloc[:, 1] = df.iloc[:, 1].fillna(method='ffill')

    current_nv = None
    empty_count = 0

    for _, row in df.iterrows():
        name_cell = str(row.iloc[1]).strip()
        if name_cell.lower() in ["组员名字", "统计", "表格不要 làm gì", "tổng"]:
            continue
        if name_cell:
            current_nv = re.sub(r"\(.*?\)", "", name_cell).strip()
        if not current_nv:
            continue

        nguon = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
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
            "Tương tác ≥10 câu": pd.to_numeric(row.get(header_map.get("Tương tác ≥10 câu")), errors="coerce") if "Tương tác ≥10 câu" in header_map else None,
            "Lượng tham gia group Zalo": pd.to_numeric(row.get(header_map.get("Lượng tham gia group Zalo")), errors="coerce") if "Lượng tham gia group Zalo" in header_map else None,
            "Tổng số kết bạn trong ngày": pd.to_numeric(row.get(header_map.get("Tổng số kết bạn trong ngày")), errors="coerce") if "Tổng số kết bạn trong ngày" in header_map else None,
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
# 🚀 Giao diện Streamlit
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

        # =====================
        # 🎯 KPI Dashboard - Tính KPI Tùy Biến
        st.markdown("---")
        st.header("📊 KPI Dashboard - Tính KPI Tùy Biến")

        st.subheader("🔢 Dữ liệu tổng hợp ban đầu")
        st.dataframe(df_all, use_container_width=True)

        st.subheader("⚙️ Cấu hình KPI Tuỳ Biến")

        kpi_cols = ["Tương tác ≥10 câu", "Lượng tham gia group Zalo", "Tổng số kết bạn trong ngày"]
        col1 = st.selectbox("Chọn cột A", kpi_cols)
        operation = st.selectbox("Phép toán", ["/", "*", "+", "-"])
        col2 = st.selectbox("Chọn cột B", kpi_cols)
        kpi_name = st.text_input("Tên chỉ số KPI mới", "Hiệu suất (%)")

        if st.button("✅ Tính KPI"):
            try:
                if operation == "/":
                    df_all[kpi_name] = df_all[col1] / df_all[col2]
                elif operation == "*":
                    df_all[kpi_name] = df_all[col1] * df_all[col2]
                elif operation == "+":
                    df_all[kpi_name] = df_all[col1] + df_all[col2]
                elif operation == "-":
                    df_all[kpi_name] = df_all[col1] - df_all[col2]
                st.success(f"✅ Đã tính KPI mới: {kpi_name}")
                st.dataframe(df_all[[col1, col2, kpi_name, "Nhân viên chuẩn", "Sheet"]], use_container_width=True)
            except Exception as e:
                st.error(f"Lỗi khi tính KPI: {e}")
    else:
        st.warning("❗ Không có dữ liệu nào được trích xuất. Vui lòng kiểm tra lại file.")
else:
    st.info("📎 Vui lòng upload file Excel để bắt đầu.")

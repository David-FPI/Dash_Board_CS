import streamlit as st
import pandas as pd
import re
import os

os.system("pip install openpyxl")

st.set_page_config(page_title="📊 Đọc tên nhân viên", page_icon="👩‍💼")

# =====================
# 🔧 Hàm chuẩn hóa tên nhân viên
def clean_employee_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"\n.*", "", name)
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip().title()


# =====================
# 📥 Đọc từng sheet
def extract_data_from_sheet(sheet_df, sheet_name):
    data = []
    rows = sheet_df.shape[0]

    sheet_df[1] = sheet_df[1].fillna(method='ffill')  # fill tên nhân viên từ merge
    current_nv = None
    empty_count = 0

    for i in range(3, rows):  # bỏ 3 dòng đầu
        row = sheet_df.iloc[i]

        # Xác định tên nhân viên từ cột B
        if pd.notna(row[1]):
            name_cell = str(row[1]).strip()
            if name_cell.lower() in ["组员名字", "统计", "表格不要 làm gì", "tổng"]:
                continue
            current_nv = re.sub(r"\(.*?\)", "", name_cell).strip()

        # Nếu không có tên thì bỏ qua
        if not current_nv:
            continue

        # Xác định nguồn từ cột C
        nguon = str(row[2]).strip() if pd.notna(row[2]) else ""
        if nguon == "" or nguon.lower() == "nan":
            empty_count += 1
            if empty_count >= 2:
                break  # kết thúc khối dữ liệu nếu trống liên tiếp 2 dòng
            continue
        else:
            empty_count = 0

        # Lưu lại dòng hợp lệ
        data.append({
            "Nhân viên": current_nv,
            "Nguồn": nguon,
            "Tương tác ≥10 câu": pd.to_numeric(row[15], errors="coerce"),
            "Group Zalo": pd.to_numeric(row[18], errors="coerce"),
            "Kết bạn trong ngày": pd.to_numeric(row[12], errors="coerce"),
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
# Giao diện upload
st.title("📥 Đọc Tên Nhân Viên Từ File Excel Báo Cáo")

uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        st.write(f"📂 Đang xử lý: `{file.name}`")
        df = extract_all_data(file)
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    if not df_all.empty:
        # Chuẩn hóa tên nhân viên
        df_all["Nhân viên chuẩn"] = df_all["Nhân viên"].apply(clean_employee_name)

        st.subheader("✅ Danh sách Nhân viên đã chuẩn hóa")
        st.dataframe(df_all[["Nhân viên", "Nhân viên chuẩn", "Sheet"]].drop_duplicates(), use_container_width=True)

        st.success(f"✅ Tổng số dòng dữ liệu: {len(df_all)} — 👩‍💻 Nhân viên duy nhất: {df_all['Nhân viên chuẩn'].nunique()}")

                # =====================
        # 📊 KPI Dashboard - Tính KPI Tùy Biến
        st.header("📊 KPI Dashboard - Tính KPI Tùy Biến")
    
        st.markdown("### 🔢 Dữ liệu tổng hợp ban đầu")
        grouped_df = df_all.groupby("Nhân viên chuẩn").agg({
            "Tương tác ≥10 câu": "sum",
            "Group Zalo": "sum",
            "Kết bạn trong ngày": "sum"
        }).reset_index()
    
        # Đổi tên cột "Kết bạn trong ngày" thành "Lượng tham gia group Zalo"
        grouped_df.rename(columns={"Kết bạn trong ngày": "Lượng tham gia group Zalo"}, inplace=True)
    
        st.dataframe(grouped_df, use_container_width=True)
    
        st.markdown("### ⚙️ Cấu hình KPI Tuỳ Biến")
    
        col1, col2, col3 = st.columns(3)
    
        with col1:
            col_a = st.selectbox("Chọn cột A", grouped_df.columns[1:], key="col_a")
        with col2:
            operation = st.selectbox("Phép toán", ["/", "*", "+", "-"], key="operation")
        with col3:
            col_b = st.selectbox("Chọn cột B", grouped_df.columns[1:], key="col_b")
    
        kpi_name = st.text_input("Tên chỉ số KPI mới", value="Hiệu suất (%)")
    
        if st.button("✅ Tính KPI"):
            try:
                # Tính KPI
                if operation == "/" and (grouped_df[col_b] == 0).any():
                    st.warning("⚠️ Có giá trị chia cho 0, KPI có thể không chính xác.")
                grouped_df[kpi_name] = grouped_df[col_a].astype(float)
    
                if operation == "+":
                    grouped_df[kpi_name] = grouped_df[col_a] + grouped_df[col_b]
                elif operation == "-":
                    grouped_df[kpi_name] = grouped_df[col_a] - grouped_df[col_b]
                elif operation == "*":
                    grouped_df[kpi_name] = grouped_df[col_a] * grouped_df[col_b]
                elif operation == "/":
                    grouped_df[kpi_name] = grouped_df[col_a] / grouped_df[col_b]
    
                # Nếu tên KPI có "%", thì nhân 100 và làm tròn
                if "%" in kpi_name:
                    grouped_df[kpi_name] = (grouped_df[kpi_name] * 100).round(2)
    
                st.success(f"✅ Đã tính KPI mới: `{kpi_name}`")
                st.dataframe(grouped_df, use_container_width=True)
            except Exception as e:
                st.error(f"❌ Lỗi khi tính KPI: {e}")


    else:
        st.warning("❗ Không có dữ liệu nào được trích xuất. Vui lòng kiểm tra lại file.")
else:
    st.info("📎 Vui lòng upload file Excel để bắt đầu.")

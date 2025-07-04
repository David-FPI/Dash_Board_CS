import streamlit as st
import pandas as pd
import re
from io import BytesIO
#os.system("pip install xlsxwriter")
st.set_page_config(page_title="📋 Danh sách Nhân Viên", layout="wide")
st.title("📋 Danh sách Nhân Viên từ File Excel")

# ===== Hàm chuẩn hóa tên nhân viên =====
def normalize_name(name):
    if pd.isna(name) or not isinstance(name, str) or name.strip() == "":
        return None
    name = re.sub(r"\(.*?\)", "", name)  # Xóa (Event), (abc) các kiểu
    name = re.sub(r"[^\w\sÀ-ỹ]", "", name)  # Xóa kí tự đặc biệt
    name = re.sub(r"\s+", " ", name).strip()
    name = name.title()  # Viết hoa đầu từ
    if name.lower() in ["nan", "组员名字", "组员"]:
        return None
    return name

# ===== Load nhiều file Excel =====
uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_names = set()
    sheet_presence = {}  # Dict: {sheet_name: [list nhân viên chuẩn hóa]}
    
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        excel_data = pd.ExcelFile(uploaded_file)
        
        for sheet in excel_data.sheet_names:
            try:
                df = excel_data.parse(sheet, skiprows=2)
            except:
                continue

            # Cố gắng tìm cột 'Tên nhân viên'
            col_match = [col for col in df.columns if 'tên' in str(col).lower() and 'nhân viên' in str(col).lower()]
            if not col_match:
                continue
            col_nv = col_match[0]

            names = df[col_nv].dropna().apply(normalize_name).dropna().unique()
            clean_names = set(names)
            
            # Lưu lại để thống kê
            sheet_presence[sheet] = clean_names
            all_names.update(clean_names)

    # ======= Tạo bảng tổng hợp tên nhân viên xuất hiện theo từng sheet =======
    all_names = sorted(all_names)
    summary_data = []

    for name in all_names:
        row = {"Tên nhân viên": name}
        total = 0
        for sheet in sheet_presence:
            if name in sheet_presence[sheet]:
                row[sheet] = "✅"
                total += 1
            else:
                row[sheet] = ""
        row["Tổng cộng"] = total
        summary_data.append(row)

    df_summary = pd.DataFrame(summary_data)

    st.success(f"✅ Tổng cộng có {len(all_names)} nhân viên duy nhất sau chuẩn hóa.")
    st.dataframe(df_summary, use_container_width=True)

    # ======= Cho phép tải về Excel =======
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Summary")
        processed_data = output.getvalue()
        return processed_data

    st.download_button(
        label="📥 Tải bảng thống kê nhân viên",
        data=to_excel(df_summary),
        file_name="Thong_Ke_Nhan_Vien.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

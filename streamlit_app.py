import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="📋 Danh sách Nhân Viên", layout="wide")
st.title("📋 Danh sách Nhân Viên từ File Excel")

# ===== Hàm chuẩn hóa tên nhân viên =====
def normalize_name(name):
    if pd.isna(name) or not isinstance(name, str) or name.strip() == "":
        return None
    name = re.sub(r"\(.*?\)", "", name)  # Xóa (Event), (abc)
    name = re.sub(r"[^\w\sÀ-ỹ]", "", name)  # Xóa ký tự đặc biệt
    name = re.sub(r"\s+", " ", name).strip()
    name = name.title()
    if name.lower() in ["nan", "组员名字", "组员"]:
        return None
    return name

# ===== Tách tên nhân viên theo block merge 5 dòng =====
def extract_names_from_column(col_series):
    names = []
    prev_name = None
    empty_count = 0

    for value in col_series:
        name = normalize_name(value)
        if name:
            if name != prev_name:
                names.append(name)
                prev_name = name
            empty_count = 0
        else:
            empty_count += 1
            if empty_count >= 2:  # Gặp 2 dòng trống liên tiếp thì dừng
                break
    return set(names)

# ===== Load nhiều file Excel =====
uploaded_files = st.file_uploader("Kéo & thả nhiều file Excel vào đây", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_names = set()
    sheet_presence = {}  # {sheet_name: set(nhân viên)}

    for uploaded_file in uploaded_files:
        excel_data = pd.ExcelFile(uploaded_file)

        for sheet_name in excel_data.sheet_names:
            try:
                df = excel_data.parse(sheet_name, header=None)
            except:
                continue

            if df.shape[1] < 2:
                continue

            col_B = df.iloc[3:, 1]  # Bỏ B1:B3
            names = extract_names_from_column(col_B)
            sheet_presence[sheet_name] = names
            all_names.update(names)

    # ======= Tạo bảng thống kê =======
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

    # ======= Cho phép tải xuống =======
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Summary")
        return output.getvalue()

    st.download_button(
        label="📥 Tải bảng thống kê nhân viên",
        data=to_excel(df_summary),
        file_name="Thong_Ke_Nhan_Vien.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

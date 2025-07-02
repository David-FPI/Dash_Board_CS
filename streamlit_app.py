import streamlit as st
import pandas as pd
import math
from pathlib import Path
import plotly.express as px
import os
os.system("pip install openpyxl")

# Set the title and favicon that appear in the Browser's tab bar.
st.set_page_config(
    page_title='GDP dashboard',
    page_icon=':earth_americas:',
)

# === Upload file ===
uploaded_file = st.file_uploader("📥 Kéo file Excel vào đây", type=["xlsx"])
if uploaded_file:

    def extract_data_from_sheet(sheet_df, sheet_name):
        data = []
        current_nv = None
        rows = sheet_df.shape[0]

        i = 3  # Bắt đầu từ dòng 4 (index 3), bỏ qua header
        while i < rows:
            row = sheet_df.iloc[i]
            name_cell = str(row[1]).strip() if pd.notna(row[1]) else ""

            if name_cell and name_cell.lower() not in ["nan", "组员名字", "表格不要做任何调整，除前两列，其余全是公式"]:
                current_nv = name_cell

                for j in range(i, i + 6):
                    if j >= rows:
                        break
                    sub_row = sheet_df.iloc[j]
                    nguon = sub_row[2]
                    if pd.isna(nguon) or str(nguon).strip() in ["", "0"]:
                        break
                    data.append({
                        "Nhân viên": current_nv.strip(),
                        "Nguồn": str(nguon).strip(),
                        "Tương tác ≥10 câu": pd.to_numeric(sub_row[15], errors="coerce"),
                        "Group Zalo": pd.to_numeric(sub_row[18], errors="coerce"),
                        "Sheet": sheet_name
                    })
                i += 6
            else:
                i += 1
        return data

    def extract_all_data(file):
        xls = pd.ExcelFile(file)
        all_rows = []

        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                if df.shape[0] < 10 or df.shape[1] < 19:
                    continue
                extracted = extract_data_from_sheet(df, sheet_name)
                all_rows.extend(extracted)
            except Exception as e:
                st.warning(f"❌ Lỗi ở sheet '{sheet_name}': {e}")

        return pd.DataFrame(all_rows), xls

    # === Xử lý file upload
    df_all, xls = extract_all_data(uploaded_file)

    # === Chuẩn hóa tên nhân viên
    df_all["Nhân viên chuẩn"] = (
        df_all["Nhân viên"]
        .astype(str)
        .str.replace(r"\n.*", "", regex=True)
        .str.strip()
    )

    # === Tổng hợp KPI theo nhân viên
    df_summary = (
        df_all.groupby("Nhân viên chuẩn")
        .agg({
            "Tương tác ≥10 câu": "sum",
            "Group Zalo": "sum"
        })
        .rename(columns={
            "Tương tác ≥10 câu": "Tổng TT ≥10 câu",
            "Group Zalo": "Tổng Group Zalo"
        })
        .reset_index()
        .sort_values(by="Tổng TT ≥10 câu", ascending=False)
    )

    df_summary["Hiệu suất nhân viên (%)"] = (
        (df_summary["Tổng Group Zalo"] / df_summary["Tổng TT ≥10 câu"]) * 100
    ).round(2).fillna(0)

    st.subheader("📋 Bảng Tổng hợp Tương Tác & Group Zalo theo Nhân Viên")
    st.dataframe(df_summary, use_container_width=True)

    st.success(f"Tổng số nhân viên: {df_summary['Nhân viên chuẩn'].nunique()}")

    # === Tổng hợp theo từng sheet + nhân viên chuẩn
    df_by_sheet = (
        df_all.groupby(["Sheet", "Nhân viên chuẩn"])
        .agg({
            "Tương tác ≥10 câu": "sum",
            "Group Zalo": "sum"
        })
        .rename(columns={
            "Tương tác ≥10 câu": "TT ≥10 câu",
            "Group Zalo": "Group Zalo"
        })
        .reset_index()
        .sort_values(by=["Nhân viên chuẩn", "Sheet"])
    )

    st.subheader("📊 Bảng Chỉ Số Tương Tác & Group Zalo Theo Từng Sheet")
    st.dataframe(df_by_sheet, use_container_width=True)

    # === Vẽ biểu đồ KPI theo thời gian
    kpi_over_time = (
        df_all.groupby(["Sheet", "Nhân viên chuẩn"])
        .agg({
            "Tương tác ≥10 câu": "sum",
            "Group Zalo": "sum"
        })
        .reset_index()
        .rename(columns={
            "Tương tác ≥10 câu": "Tương tác",
            "Group Zalo": "Group"
        })
    )

    st.subheader(":bar_chart: Biểu đồ KPI theo thời gian")

    unique_employees = kpi_over_time["Nhân viên chuẩn"].unique().tolist()
    selected_employees = st.multiselect(
        "Chọn nhân viên cần xem:", unique_employees, default=unique_employees[:5]
    )

    kpi_option = st.selectbox(
        "Chọn KPI muốn theo dõi:",
        ["Tương tác", "Group"]
    )

    filtered_df = kpi_over_time[kpi_over_time["Nhân viên chuẩn"].isin(selected_employees)]

    if filtered_df.empty:
        st.warning("⚠️ Không có dữ liệu để hiển thị. Vui lòng chọn nhân viên có dữ liệu.")
    else:
        fig = px.line(
            filtered_df,
            x="Sheet",
            y=kpi_option,
            color="Nhân viên chuẩn",
            markers=True,
            title=f"Biểu đồ {kpi_option} qua các Sheet"
        )
        fig.update_layout(
            xaxis_title="Sheet",
            yaxis_title=kpi_option,
            legend_title="Nhân viên",
            hovermode="x unified",
            height=500
        )
        st.plotly_chart(fig, use_container_width=True)

    # === Tổng số kết bạn trong ngày theo nhân viên ===
    def extract_friend_adds(xls):
        all_data = []

        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                if df.shape[0] < 10 or df.shape[1] < 13:
                    continue

                i = 3
                current_nv = None

                while i < df.shape[0]:
                    row = df.iloc[i]
                    name = str(row[1]).strip() if pd.notna(row[1]) else ""

                    if name and name.lower() not in ["nan", "组员名字", "表格不要做任何调整，除前两列，其余全是公式"]:
                        current_nv = name
                        for j in range(i, i + 6):
                            if j >= df.shape[0]:
                                break
                            sub_row = df.iloc[j]
                            if pd.isna(sub_row[2]) or str(sub_row[2]).strip() == "":
                                break
                            friend_adds = pd.to_numeric(sub_row[9], errors="coerce")
                            all_data.append({
                                "Sheet": sheet,
                                "Nhân viên": current_nv,
                                "Kết bạn trong ngày": friend_adds
                            })
                        i += 6
                    else:
                        i += 1
            except Exception as e:
                continue

        return pd.DataFrame(all_data)

    df_friends = extract_friend_adds(xls)

    df_friends["Nhân viên chuẩn"] = df_friends["Nhân viên"].astype(str).str.replace(r"\n.*", "", regex=True).str.strip()

    friend_summary = (
        df_friends.groupby("Nhân viên chuẩn")["Kết bạn trong ngày"]
        .sum()
        .reset_index()
        .sort_values(by="Kết bạn trong ngày", ascending=False)
    )

    st.subheader("📋 Bảng Tổng hợp Kết Bạn Trong Ngày theo Nhân Viên")
    st.dataframe(friend_summary, use_container_width=True)

    # Merge để tạo bảng mới giống df_summary nhưng thêm cột Kết bạn
    merged_summary = pd.merge(df_summary.drop(columns=["Hiệu suất nhân viên (%)"]), friend_summary, on="Nhân viên chuẩn", how="left")

    st.subheader("📋 Bảng Tổng hợp Tương Tác & Group Zalo & Kết Bạn theo Nhân Viên")
    st.dataframe(merged_summary, use_container_width=True)

else:
    st.info("📎 Vui lòng tải lên file Excel báo cáo để bắt đầu.")

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
    page_icon=':earth_americas:', # This is an emoji shortcode. Could be a URL too.
)

# # -----------------------------------------------------------------------------
# # Declare some useful functions.

# @st.cache_data
# def get_gdp_data():
#     """Grab GDP data from a CSV file.

#     This uses caching to avoid having to read the file every time. If we were
#     reading from an HTTP endpoint instead of a file, it's a good idea to set
#     a maximum age to the cache with the TTL argument: @st.cache_data(ttl='1d')
#     """

#     # Instead of a CSV on disk, you could read from an HTTP endpoint here too.
#     DATA_FILENAME = Path(__file__).parent/'data/gdp_data.csv'
#     raw_gdp_df = pd.read_csv(DATA_FILENAME)

#     MIN_YEAR = 1960
#     MAX_YEAR = 2022

#     # The data above has columns like:
#     # - Country Name
#     # - Country Code
#     # - [Stuff I don't care about]
#     # - GDP for 1960
#     # - GDP for 1961
#     # - GDP for 1962
#     # - ...
#     # - GDP for 2022
#     #
#     # ...but I want this instead:
#     # - Country Name
#     # - Country Code
#     # - Year
#     # - GDP
#     #
#     # So let's pivot all those year-columns into two: Year and GDP
#     gdp_df = raw_gdp_df.melt(
#         ['Country Code'],
#         [str(x) for x in range(MIN_YEAR, MAX_YEAR + 1)],
#         'Year',
#         'GDP',
#     )

#     # Convert years from string to integers
#     gdp_df['Year'] = pd.to_numeric(gdp_df['Year'])

#     return gdp_df

# gdp_df = get_gdp_data()

# # -----------------------------------------------------------------------------
# # Draw the actual page

# # Set the title that appears at the top of the page.
# '''
# # :earth_americas: GDP dashboard

# Browse GDP data from the [World Bank Open Data](https://data.worldbank.org/) website. As you'll
# notice, the data only goes to 2022 right now, and datapoints for certain years are often missing.
# But it's otherwise a great (and did I mention _free_?) source of data.
# '''

# # Add some spacing
# ''
# ''

# min_value = gdp_df['Year'].min()
# max_value = gdp_df['Year'].max()

# from_year, to_year = st.slider(
#     'Which years are you interested in?',
#     min_value=min_value,
#     max_value=max_value,
#     value=[min_value, max_value])

# countries = gdp_df['Country Code'].unique()

# if not len(countries):
#     st.warning("Select at least one country")

# selected_countries = st.multiselect(
#     'Which countries would you like to view?',
#     countries,
#     ['DEU', 'FRA', 'GBR', 'BRA', 'MEX', 'JPN'])

# ''
# ''
# ''

# # Filter the data
# filtered_gdp_df = gdp_df[
#     (gdp_df['Country Code'].isin(selected_countries))
#     & (gdp_df['Year'] <= to_year)
#     & (from_year <= gdp_df['Year'])
# ]

# st.header('GDP over time', divider='gray')

# ''

# st.line_chart(
#     filtered_gdp_df,
#     x='Year',
#     y='GDP',
#     color='Country Code',
# )

# ''
# ''


# first_year = gdp_df[gdp_df['Year'] == from_year]
# last_year = gdp_df[gdp_df['Year'] == to_year]

# st.header(f'GDP in {to_year}', divider='gray')

# ''

# cols = st.columns(4)

# for i, country in enumerate(selected_countries):
#     col = cols[i % len(cols)]

#     with col:
#         first_gdp = first_year[first_year['Country Code'] == country]['GDP'].iat[0] / 1000000000
#         last_gdp = last_year[last_year['Country Code'] == country]['GDP'].iat[0] / 1000000000

#         if math.isnan(first_gdp):
#             growth = 'n/a'
#             delta_color = 'off'
#         else:
#             growth = f'{last_gdp / first_gdp:,.2f}x'
#             delta_color = 'normal'

#         st.metric(
#             label=f'{country} GDP',
#             value=f'{last_gdp:,.0f}B',
#             delta=growth,
#             delta_color=delta_color
#         )


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

            # Nếu có tên nhân viên hợp lệ
            if name_cell and name_cell.lower() not in ["nan", "组员名字", "表格不要做任何调整，除前两列，其余全是公式"]:
                current_nv = name_cell

                # Đọc 6 dòng nguồn kế tiếp
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

        return pd.DataFrame(all_rows)

    # === Xử lý file upload
    df_all = extract_all_data(uploaded_file)

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
# === Tính thêm cột Hiệu suất (Group Zalo / Tương tác ≥10 câu) * 100
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


# === Sau khi df_all đã được xử lý và có cột "Nhân viên chuẩn" ===
    
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
# Lấy dữ liệu cột "Tổng số kết bạn trong ngày" từ tất cả các sheet và group by nhân viên chuẩn
    
    def extract_friend_adds(xls):
        all_data = []
    
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
    
                if df.shape[0] < 10 or df.shape[1] < 13:
                    continue
    
                i = 3  # Bỏ qua 3 dòng đầu
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
                            name_in_loop = str(sub_row[1]).strip()
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
    
    # Chuẩn hóa tên nhân viên
    df_friends["Nhân viên chuẩn"] = df_friends["Nhân viên"].astype(str).str.replace(r"\n.*", "", regex=True).str.strip()
    
    # Tổng hợp
    friend_summary = (
        df_friends.groupby("Nhân viên chuẩn")["Kết bạn trong ngày"]
        .sum()
        .reset_index()
        .sort_values(by="Kết bạn trong ngày", ascending=False)
    )
    
    friend_summary.head(10)

else:
    st.info("📎 Vui lòng tải lên file Excel báo cáo để bắt đầu.")

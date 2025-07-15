import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Tổng hợp báo cáo nhân viên", layout="wide")
st.title("📊 Tổng hợp báo cáo nhân viên từ nhiều file Excel")

# ======================== Hàm chuẩn hóa & xử lý ========================
def normalize_staff_name(name):
    if not isinstance(name, str):
        return ""
    name = re.sub(r"\(.*?\)", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name

def extract_header(df):
    if len(df) < 3:
        raise ValueError("File thiếu dòng tiêu đề (ít hơn 3 dòng).")
    row1 = df.iloc[1].fillna("")
    row2 = df.iloc[2].fillna("")
    header = row1.astype(str) + " " + row2.astype(str)
    header = header.str.replace(r"\s+", " ", regex=True).str.strip()
    return header

def extract_data_block(df_raw):
    header = extract_header(df_raw)
    header = pd.Series(header)
    header = header.where(~header.duplicated(), header + "_" + header.groupby(header).cumcount().astype(str))
    df_data = df_raw.iloc[3:].copy()

    cutoff_idx = df_data[df_data.iloc[:, 0].astype(str).str.contains("统计|Tổng", case=False, na=False)].index
    if not cutoff_idx.empty:
        df_data = df_data.loc[:cutoff_idx[0] - 1]

    df_data.columns = header
    df_data.reset_index(drop=True, inplace=True)

    staff_col = next((c for c in df_data.columns if "nhân viên" in c.lower()), None)
    if not staff_col:
        raise ValueError("Không tìm thấy cột Nhân viên.")

    last_name = None
    for i in range(len(df_data)):
        name = df_data.at[i, staff_col]
        if pd.notna(name) and str(name).strip() != "":
            last_name = normalize_staff_name(name)
        elif last_name:
            df_data.at[i, staff_col] = last_name

    return df_data

# ======================== Xử lý nhiều sheet ========================
def process_all_sheets(file):
    xls = pd.ExcelFile(file)
    all_data = []
    log = []
    for sheet in xls.sheet_names:
        try:
            df_raw = xls.parse(sheet, header=None)
            cleaned = extract_data_block(df_raw)
            cleaned["__Sheet__"] = sheet
            all_data.append(cleaned)
            log.append({"Sheet": sheet, "Status": "✅ Đã xử lý", "Rows": len(cleaned)})
        except Exception as e:
            log.append({"Sheet": sheet, "Status": f"❌ Bỏ qua - {str(e)}", "Rows": 0})
    log_df = pd.DataFrame(log)
    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame(), log_df

# ======================== Chuẩn hóa tên cột ========================
def normalize_column_name(col):
    col = str(col)
    col = re.sub(r"\s+", " ", col)
    col = col.strip().lower()
    return col

# ======================== Giao diện ========================
uploaded_files = st.file_uploader("📁 Tải lên 1 hoặc nhiều file báo cáo Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    full_data = []
    for file in uploaded_files:
        try:
            st.success(f"✔️ Đang xử lý: {file.name}")
            df_all, sheet_log = process_all_sheets(file)
            df_all["__File__"] = file.name
            full_data.append(df_all)
        except Exception as e:
            st.error(f"❌ Lỗi khi xử lý file {file.name}: {e}")

    if full_data:
        df_final = pd.concat(full_data, ignore_index=True)
        # 🔍 In thử các sheet và số cột nhận được từ mỗi sheet
        st.markdown("### 📌 Check: Cột nhận được từ mỗi sheet")

        if "__Sheet__" in df_final.columns:
            sheet_col_map = df_final.groupby("__Sheet__").agg(lambda x: list(x.index)).reset_index()
            sheet_col_map["Số dòng"] = sheet_col_map["__Sheet__"].apply(lambda sheet: len(df_final[df_final["__Sheet__"] == sheet]))
            sheet_col_map["Số cột"] = sheet_col_map["__Sheet__"].apply(lambda sheet: df_final[df_final["__Sheet__"] == sheet].shape[1])
            st.dataframe(sheet_col_map[["__Sheet__", "Số dòng", "Số cột"]], use_container_width=True)

            # Optional: hiển thị 3 dòng đầu của mỗi sheet
            for sheet in df_final["__Sheet__"].unique():
                st.markdown(f"#### 🧾 Sheet: `{sheet}` - 3 dòng đầu")
                st.dataframe(df_final[df_final["__Sheet__"] == sheet].head(3), use_container_width=True)
        else:
            st.warning("⚠️ Không tìm thấy cột '__Sheet__'. Có thể hàm process_all_sheets() đang bị lỗi.")




        st.subheader("✅ Dữ liệu đã tổng hợp")
        st.dataframe(df_final.head(50), use_container_width=True)

        # —————— START: TÍNH KPI ——————
        normalized_cols = {c: normalize_column_name(c) for c in df_final.columns}

        kpi_ketban_keywords = ["tổng số kết bạn trong ngày", "当天加zalo总数"]
        kpi_tuongtac_keywords = ["tương tác ≥10 câu", "≥10"]
        kpi_groupzalo_keywords = ["tham gia group zalo", "lượng tham gia group zalo"]
        # --- CÁC KPI MỞ RỘNG MỚI ---
        kpi_1_1_keywords = ["trao đổi 1-1", "私信zalo数", "tổng trao đổi trong ngày"]
        kpi_duoi10_keywords = ["đối thoại (<10 câu)", "对话 (<10 句)", "đối thoại", "trao đổi <10 câu", "<10"]
        kpi_khong_phan_hoi_keywords = ["không phản hồi", "无回复", "无"]


        def find_cols_by_keywords(keywords):
            return [orig for orig, norm in normalized_cols.items() if any(kw in norm for kw in keywords)]

        def find_col_by_keywords(keywords):
            return next((orig for orig, norm in normalized_cols.items()
                        if any(kw in norm for kw in keywords)), None)

        cols_ketban = find_cols_by_keywords(kpi_ketban_keywords)
        cols_tuongtac = find_cols_by_keywords(kpi_tuongtac_keywords)
        cols_groupzalo = find_cols_by_keywords(kpi_groupzalo_keywords)
        cols_1_1 = find_cols_by_keywords(kpi_1_1_keywords)
        cols_duoi10 = find_cols_by_keywords(kpi_duoi10_keywords)
        cols_khong_phan_hoi = find_cols_by_keywords(kpi_khong_phan_hoi_keywords)

        if not (cols_ketban and cols_tuongtac and cols_groupzalo):
            st.warning("⚠️ Không tìm đủ 3 cột KPI (kết bạn, tương tác, group Zalo). Vui lòng kiểm tra lại tên cột.")
        else:
            df_final["kpi_ketban"] = df_final[cols_ketban].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
            df_final["kpi_tuongtac_tren_10"] = df_final[cols_tuongtac].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
            df_final["kpi_groupzalo"] = df_final[cols_groupzalo].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
            df_final["kpi_traodoi_1_1"] = df_final[cols_1_1].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
            df_final["kpi_doi_thoai_duoi_10"] = df_final[cols_duoi10].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
            df_final["kpi_khong_phan_hoi"] = df_final[cols_khong_phan_hoi].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

            # 🎯 Nâng cấp tìm cột Nhân viên và Nguồn
            staff_keywords = ["nhân viên", "人员", "成员"]
            source_keywords = ["nguồn", "渠道"]

            staff_col = find_col_by_keywords(staff_keywords)
            source_col = find_col_by_keywords(source_keywords)

            kpi_cols = [
                "kpi_ketban", "kpi_traodoi_1_1", "kpi_tuongtac_tren_10", "kpi_doi_thoai_duoi_10", "kpi_khong_phan_hoi", "kpi_groupzalo"]


            df_kpi = df_final.groupby([staff_col, source_col], as_index=False)[kpi_cols].sum()
            st.subheader("📈 KPI theo nhân viên và nguồn")
            st.dataframe(df_kpi, use_container_width=True)



            df_kpi_total = df_kpi.groupby(staff_col, as_index=False)[kpi_cols].sum()
            # ===== 🔧 KPI tùy biến (cộng trừ nhân chia giữa các cột) =====
            with st.expander("🧮 Thiết kế công thức KPI tuỳ biến", expanded=False):
                col_names = df_kpi_total.columns.tolist()
                selected_cols = st.multiselect("📌 Chọn cột muốn dùng trong công thức:", col_names, default=[])

                common_formulas = {
                    "Hiệu suất group / kết bạn (%)": "kpi_groupzalo / kpi_ketban * 100",
                    "Tỉ lệ không phản hồi": "kpi_khong_phan_hoi / kpi_traodoi_1_1 * 100",
                    "Tương tác >10 / kết bạn (%)": "kpi_tuongtac_tren_10 / kpi_ketban * 100"
                }
                selected_common = st.selectbox("📂 Chọn công thức mẫu:", [""] + list(common_formulas.keys()))
                if selected_common:
                    custom_formula = common_formulas[selected_common]
                    st.info(f"📌 Công thức đã chọn: `{custom_formula}`")
                else:
                    custom_formula = st.text_input("🧠 Nhập công thức KPI tuỳ chỉnh")

                custom_col_name = st.text_input("📝 Tên cột mới:", value="KPI tuỳ biến")

                if st.button("🚀 Tính KPI tuỳ biến"):
                    if selected_cols and custom_formula and custom_col_name:
                        try:
                            calc_df = df_kpi_total[selected_cols].copy()
                            result = eval(custom_formula, {}, calc_df.to_dict("series"))
                            df_kpi_total[custom_col_name] = pd.to_numeric(result, errors="coerce").round(2)
                            st.success(f"✅ Đã thêm cột: {custom_col_name}")
                        except Exception as e:
                            st.error(f"❌ Lỗi công thức: {e}")
                    else:
                        st.warning("⚠️ Cần chọn đủ cột, công thức và tên cột mới.")

            


            # df_kpi_total["Hiệu suất (%)"] = pd.to_numeric(df_kpi_total["Hiệu suất (%)"], errors="coerce").round(2)


            st.subheader("📊 KPI tổng hợp theo nhân viên")
            st.dataframe(df_kpi_total, use_container_width=True)

        # —————— END: TÍNH KPI ——————

        csv = df_final.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Tải dữ liệu tổng hợp CSV", csv, "tong_hop_bao_cao.csv", "text/csv")

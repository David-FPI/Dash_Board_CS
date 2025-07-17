import streamlit as st
import pandas as pd
import re
import io
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
                st.markdown(f"#### 🧾 Sheet: {sheet} - 3 dòng đầu")
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
        kpi_duoi10_keywords = ["<10"]
        kpi_khong_phan_hoi_keywords = ["không phản hồi", "无回复", "无"]

        # Gán tiêu đề gốc để gắn nhãn dễ hiểu
        kpi_label_map = {}
 

        def find_cols_by_keywords(keywords, kpi_name=None):
            found = []
            for orig, norm in normalized_cols.items():
                if any(kw in norm for kw in keywords):
                    found.append(orig)
                    if kpi_name:
                        kpi_label_map[kpi_name] = orig  # chỉ lưu 1 tiêu đề đầu tiên
            return found

        # def find_cols_by_keywords(keywords):
        #     return [orig for orig, norm in normalized_cols.items() if any(kw in norm for kw in keywords)]

        def find_col_by_keywords(keywords):
            return next((orig for orig, norm in normalized_cols.items()
                        if any(kw in norm for kw in keywords)), None)

        cols_ketban = find_cols_by_keywords(kpi_ketban_keywords)
        cols_tuongtac = find_cols_by_keywords(kpi_tuongtac_keywords)
        cols_groupzalo = find_cols_by_keywords(kpi_groupzalo_keywords)
        cols_1_1 = find_cols_by_keywords(kpi_1_1_keywords)
        cols_duoi10 = find_cols_by_keywords(kpi_duoi10_keywords)
        cols_khong_phan_hoi = find_cols_by_keywords(kpi_khong_phan_hoi_keywords)
        # === Bổ sung các nhóm KPI mới ===
        kpi_luong_data_kh_keywords = ["弹窗", "客户联系社交媒体", "khách hàng nhắn tin", "流量"]
        kpi_zalo_meta_moi_keywords = ["（新）"]
        kpi_zalo_meta_cu_keywords = ["（老）"]
        kpi_zalo_meta_keywords = ["社交媒体加zalo好友"]
        kpi_zalo_sdt_moi_keywords = ["sdt加zalo好友新"]
        kpi_zalo_sdt_cu_keywords = ["sdt加zalo好友老"]
        kpi_zalo_sdt_keywords = ["sdt加zalo好友"]

        # Tạo set để loại trừ trùng lặp
        used_cols = set(cols_ketban + cols_tuongtac + cols_groupzalo + cols_1_1 + cols_duoi10 + cols_khong_phan_hoi)

        def find_col_exclude_used(keywords):
            for orig, norm in normalized_cols.items():
                if orig not in used_cols and any(kw in norm for kw in keywords):
                    used_cols.add(orig)
                    return orig
            return None
        
        # Dò từng cột và gán vào df_final nếu tìm được
        kpi_extra_mapping = {
            "kpi_luong_data_kh": find_col_exclude_used(kpi_luong_data_kh_keywords),
            "kpi_zalo_meta_moi": find_col_exclude_used(kpi_zalo_meta_moi_keywords),
            "kpi_zalo_meta_cu": find_col_exclude_used(kpi_zalo_meta_cu_keywords),
            "kpi_zalo_meta": find_col_exclude_used(kpi_zalo_meta_keywords),
            "kpi_zalo_sdt_moi": find_col_exclude_used(kpi_zalo_sdt_moi_keywords),
            "kpi_zalo_sdt_cu": find_col_exclude_used(kpi_zalo_sdt_cu_keywords),
            "kpi_zalo_sdt": find_col_exclude_used(kpi_zalo_sdt_keywords)
        }


        for kpi_name, col_name in kpi_extra_mapping.items():
            if col_name:
                df_final[kpi_name] = pd.to_numeric(df_final[col_name], errors="coerce").fillna(0)



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
                "kpi_luong_data_kh", "kpi_zalo_meta_moi", "kpi_zalo_meta_cu", "kpi_zalo_meta",
                "kpi_zalo_sdt_moi", "kpi_zalo_sdt_cu", "kpi_zalo_sdt",
                "kpi_ketban", "kpi_traodoi_1_1", "kpi_doi_thoai_duoi_10", "kpi_tuongtac_tren_10",   
                "kpi_khong_phan_hoi",           "kpi_groupzalo"

            ]



            # Dựa vào logic dò cột trong code của bạn, đây là phần nên thêm để dò thêm 3 cột:
            #   - "AI1" => kpi_ai1
            #   - "Block Chain1" => kpi_block_chain1
            #   - "Web31" => kpi_web31

            # Bổ sung từ khóa cho các KPI này:
            kpi_ai1_keywords = ["ai1"]
            kpi_block_chain1_keywords = ["block chain1"]
            kpi_web31_keywords = ["web31"]

            # Thêm vào sau các phần dò cột mở rộng khác:
            kpi_extra_mapping.update({
                "kpi_ai1": find_col_exclude_used(kpi_ai1_keywords),
                "kpi_block_chain1": find_col_exclude_used(kpi_block_chain1_keywords),
                "kpi_web31": find_col_exclude_used(kpi_web31_keywords)
            })

            # Gán giá trị vào df_final nếu cột tồn tại:
            for kpi_name in ["kpi_ai1", "kpi_block_chain1", "kpi_web31"]:
                col_name = kpi_extra_mapping.get(kpi_name)
                if col_name:
                    df_final[kpi_name] = pd.to_numeric(df_final[col_name], errors="coerce").fillna(0)
                    kpi_cols.append(kpi_name)



            df_kpi = df_final.groupby([staff_col, source_col], as_index=False)[kpi_cols].sum()
            st.subheader("📈 KPI theo nhân viên và nguồn")
            st.dataframe(df_kpi, use_container_width=True)



            df_kpi_total = df_kpi.groupby(staff_col, as_index=False)[kpi_cols].sum()
            df_kpi_total.insert( df_kpi_total.columns.get_loc("kpi_khong_phan_hoi") + 1, "kpi_ty_le_phan_hoi (%)",   ((df_kpi_total["kpi_traodoi_1_1"] - df_kpi_total["kpi_khong_phan_hoi"]) / df_kpi_total["kpi_traodoi_1_1"] * 100).round(2))
            # ➕ Thêm dòng Tổng cộng
            total_row = df_kpi_total[kpi_cols].sum(numeric_only=True)
            total_row[staff_col] = "Tổng cộng"


            # Tính lại kpi_ty_le_phan_hoi (%) cho dòng Tổng cộng
            # Tính lại kpi_ty_le_phan_hoi (%) cho dòng Tổng cộng
            try:
                tong_traodoi = total_row.get("kpi_traodoi_1_1", 0)
                tong_khong_phan_hoi = total_row.get("kpi_khong_phan_hoi", 0)
                if tong_traodoi != 0:
                    ty_le = round((tong_traodoi - tong_khong_phan_hoi) / tong_traodoi * 100, 2)
                    total_row["kpi_ty_le_phan_hoi (%)"] = f"{ty_le:.2f}%"
                else:
                    total_row["kpi_ty_le_phan_hoi (%)"] = ""
            except:
                total_row["kpi_ty_le_phan_hoi (%)"] = ""

            df_kpi_total = pd.concat([df_kpi_total, pd.DataFrame([total_row])], ignore_index=True)
            # 🔍 Kiểm tra tổng chi tiết có khớp với kpi_traodoi_1_1 không
            tong_chi_tiet = (
                df_kpi_total["kpi_doi_thoai_duoi_10"] +
                df_kpi_total["kpi_tuongtac_tren_10"] +
                df_kpi_total["kpi_khong_phan_hoi"]
            )
            
            chenh_lech = df_kpi_total["kpi_traodoi_1_1"] - tong_chi_tiet
            
            # Ghi chú: Nếu đúng thì 'Yes', nếu sai thì 'No (+x)'
            df_kpi_total["kpi_check_1_1"] = chenh_lech.apply(lambda x: "Yes" if x == 0 else f"No ({x:+.0f})")

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
                    st.info(f"📌 Công thức đã chọn: {custom_formula}")
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
            rename_display = {
                kpi: f"{kpi} ({kpi_label_map.get(kpi, '')})"
                for kpi in kpi_cols if kpi in kpi_label_map
            }




        st.subheader("📊 KPI tổng hợp theo nhân viên")
        st.dataframe(df_kpi_total, use_container_width=True)


        # —————— END: TÍNH KPI ——————

        # Tạo file Excel trong bộ nhớ
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_kpi_total.to_excel(writer, index=False, sheet_name='Tổng hợp')
        output.seek(0)
        processed_data = output.getvalue()
        # Nút tải về file Excel
        st.download_button(
            label="📥 Tải dữ liệu tổng hợp Excel",
            data=processed_data,
            file_name="tong_hop_bao_cao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)  
        

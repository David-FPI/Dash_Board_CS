import streamlit as st
import pandas as pd
import re
import os
from unidecode import unidecode

st.set_page_config(page_title="\ud83d\udcc5 \u0110\u1ecdc T\u00ean Nh\u00e2n Vi\u00ean & T\u00ednh KPI", page_icon="\ud83d\udcbc")

# =====================
# \ud83d\udd27 T\u1ef1 \u0111\u1ed9ng c\u00e0i package (n\u1ebfu ch\u01b0a c\u00f3)
os.system("pip install openpyxl unidecode")

# =====================
# \ud83d\udd27 Chu\u1ea9n h\u00f3a text \u0111\u1ec3 so s\u00e1nh

def normalize_text(text):
    text = str(text).lower()
    text = re.sub(r"[\n\r]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = unidecode(text.strip())
    return text

# =====================
# \ud83d\udd8a\ufe0f T\u1eeb \u0111i\u1ec3n keyword \u0111\u1ec3 mapping c\u1ed9t
COLUMN_MAPPING_KEYWORDS = {
    "T\u01b0\u01a1ng t\u00e1c \u226510 c\u00e2u": ["10 cau", ">=10", "tuong tac", "so cau tuong tac"],
    "L\u01b0\u1ee3ng tham gia group Zalo": ["group zalo", "tham gia zalo", "nhom zalo", "zalo group", "join group", "zalo"],
    "T\u1ed5ng s\u1ed1 k\u1ebft b\u1ea1n trong ng\u00e0y": ["ket ban", "tong so ket ban", "ket ban trong ngay", "add zalo"]
}

# =====================
# \ud83d\udcc2 Tr\u00edch xu\u1ea5t d\u1eef li\u1ec7u t\u1eeb sheet

def extract_data_from_sheet(df, sheet_name):
    data = []
    rows = df.shape[0]

    # Lo\u1ea1i b\u1ecf 2 d\u00f2ng \u0111\u1ea7u \u0111\u1ec3 l\u1ea5y header d\u00f2ng 3
    df = df.iloc[2:].reset_index(drop=True)
    df.columns = [normalize_text(col) for col in df.iloc[0]]
    df = df[1:].reset_index(drop=True)

    # Mapping column names
    col_mapping = {}
    for standard_name, keywords in COLUMN_MAPPING_KEYWORDS.items():
        for col in df.columns:
            for keyword in keywords:
                if keyword in col:
                    col_mapping[standard_name] = col
                    break
            if standard_name in col_mapping:
                break

    found_cols = list(col_mapping.keys())
    if len(found_cols) < 3:
        st.warning(f"\u26a0\ufe0f Sheet {sheet_name} kh\u00f4ng \u0111\u1ee7 c\u1ed9t KPI \u2014 d\u00f2 \u0111\u01b0\u1ee3c {found_cols}")
        return []

    # Fill tên nhân viên từ cột B
    if 1 in df.columns:
        df[1] = df[1].fillna(method='ffill')

    current_nv = None
    empty_count = 0
    for _, row in df.iterrows():
        if pd.notna(row[1]):
            name_cell = str(row[1]).strip()
            if name_cell.lower() in ["\u7ec4\u5458\u540d\u5b57", "\u7edf\u8ba1", "b\u1ea3ng t\u1ed5ng", "t\u1ed5ng"]:
                continue
            current_nv = re.sub(r"\\(.*?\\)", "", name_cell).strip()

        if not current_nv:
            continue

        nguon = str(row[2]).strip() if pd.notna(row[2]) else ""
        if nguon == "" or nguon.lower() == "nan":
            empty_count += 1
            if empty_count >= 2:
                break
            continue
        else:
            empty_count = 0

        data.append({
            "Nh\u00e2n vi\u00ean": current_nv,
            "Ngu\u1ed3n": nguon,
            "Sheet": sheet_name,
            **{k: pd.to_numeric(row[col_mapping[k]], errors="coerce") for k in col_mapping}
        })

    return data

# =====================
# \ud83d\udcc3 X\u1eed l\u00fd to\u00e0n b\u1ed9 file

def extract_all_data(file):
    xls = pd.ExcelFile(file)
    all_rows = []

    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            if df.shape[0] < 10:
                continue
            extracted = extract_data_from_sheet(df, sheet_name)
            all_rows.extend(extracted)
        except Exception as e:
            st.warning(f"\u274c L\u1ed7i sheet {sheet_name}: {e}")

    return pd.DataFrame(all_rows)

# =====================
# \ud83d\udcc5 Giao di\u1ec7n Streamlit

st.title("\ud83d\udcc5 \u0110\u1ecdc T\u00ean Nh\u00e2n Vi\u00ean & T\u00ednh KPI T\u1eeb File Excel B\u00e1o C\u00e1o")

uploaded_files = st.file_uploader("K\u00e9o & th\u1ea3 nhi\u1ec1u file Excel v\u00e0o \u0111\u00e2y", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        st.write(f"\ud83d\udcc2 \u0110ang x\u1eed l\u00fd: `{file.name}`")
        df = extract_all_data(file)
        all_data.append(df)

    df_all = pd.concat(all_data, ignore_index=True)

    if not df_all.empty:
        df_all["Nh\u00e2n vi\u00ean chuẩn"] = df_all["Nh\u00e2n vi\u00ean"].apply(lambda x: re.sub(r"\\(.*?\\)", "", str(x)).strip().title())

        st.subheader("\u2705 Danh s\u00e1ch Nh\u00e2n vi\u00ean \u0111\u00e3 chu\u1ea9n h\u00f3a")
        st.dataframe(df_all[["Nh\u00e2n vi\u00ean", "Nh\u00e2n vi\u00ean chuẩn", "Sheet"]].drop_duplicates(), use_container_width=True)

        st.success(f"\u2705 T\u1ed5ng s\u1ed1 d\u00f2ng d\u1eef li\u1ec7u: {len(df_all)} \u2014 \ud83d\udc69\u200d\ud83d\udcbc Nh\u00e2n vi\u00ean duy nh\u1ea5t: {df_all['Nh\u00e2n vi\u00ean chuẩn'].nunique()}")

        # ========== KPI Tu\u1ef3 Bi\u1ebfn ==========
        st.header("\ud83d\udcca KPI Dashboard - T\u00ednh KPI Tu\u1ef3 Bi\u1ebfn")
        st.subheader("\ud83d\udd22 D\u1eef li\u1ec7u t\u1ed5ng h\u1ee3p ban \u0111\u1ea7u")
        st.dataframe(df_all.head(), use_container_width=True)

        st.subheader("\u2699\ufe0f C\u1ea5u h\u00ecnh KPI Tu\u1ef3 Bi\u1ebfn")
        kpi_col1 = st.selectbox("Ch\u1ecdn c\u1ed9t A", df_all.columns)
        operator = st.selectbox("Ph\u00e9p to\u00e1n", ["/", "*", "+", "-"])
        kpi_col2 = st.selectbox("Ch\u1ecdn c\u1ed9t B", df_all.columns)
        kpi_name = st.text_input("T\u00ean ch\u1ec9 s\u1ed1 KPI m\u1edbi", "Hi\u1ec7u su\u1ea5t (%)")

        if st.button("\u2705 T\u00ednh KPI"):
            try:
                if operator == "/":
                    df_all[kpi_name] = df_all[kpi_col1] / df_all[kpi_col2] * 100
                elif operator == "*":
                    df_all[kpi_name] = df_all[kpi_col1] * df_all[kpi_col2]
                elif operator == "+":
                    df_all[kpi_name] = df_all[kpi_col1] + df_all[kpi_col2]
                elif operator == "-":
                    df_all[kpi_name] = df_all[kpi_col1] - df_all[kpi_col2]
                st.success(f"\u2705 KPI m\u1edbi \u0111\u00e3 \u0111\u01b0\u1ee3c t\u00ednh: {kpi_name}")
                st.dataframe(df_all[["Nh\u00e2n vi\u00ean chuẩn", kpi_name, "Sheet"]], use_container_width=True)
            except Exception as e:
                st.error(f"\u26a0\ufe0f L\u1ed7i khi t\u00ednh KPI: {e}")
    else:
        st.warning("\u2757\ufe0f Kh\u00f4ng c\u00f3 d\u1eef li\u1ec7u n\u00e0o \u0111\u01b0\u1ee3c tr\u00edch xu\u1ea5t. Vui l\u00f2ng ki\u1ec3m tra file.")
else:
    st.info("\ud83d\udcc1 Vui l\u00f2ng upload file Excel \u0111\u1ec3 b\u1eaft \u0111\u1ea7u.")

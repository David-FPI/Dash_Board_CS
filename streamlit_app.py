import streamlit as st
import pandas as pd
import unicodedata
import re
import os

st.set_page_config(page_title="📅 Đọc Tên Nhân Viên & Tính KPI", page_icon="💼")

# =====================
# 🔧 Tự động cài package nếu thiếu
os.system("pip install openpyxl")

# ✅ Hàm chuẩn hóa văn bản tiêu đề
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = text.strip().lower()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = re.sub(r'\s+', ' ', text)
    return text

# ✅ Danh sách keyword cho các chỉ số KPI
KEYWORDS_KET_BAN = [
    "kết bạn", "tổng số kết bạn", "tổng kết bạn", "số kết bạn trong ngày",
    "当天加zalo总数", "当天加好友", "当天加 zalo", "加好友", "加好友人数",
    "当天加好友数", "总加好友", "add friend", "total add friend",
    "friend request", "friends added", "用户邀请加好友", "邀请加好友",
    "zalo add", "加zalo", "zalo số lượng kết bạn", "số bạn zalo", "邀请进群zalo"
]

KEYWORDS_TUONG_TAC = [
    "≥10", ">=10", "10 câu", "tuong tac", "số lượng tương tác",
    "tương tác 10 câu", "tương tác", "互动", "số câu hỏi",
    "tương tác với khách", "≥10句", "互动次数"
]

KEYWORDS_GROUP_ZALO = [
    "group zalo", "zalo group", "tham gia group", "tham gia zalo",
    "nhóm zalo", "zalo nhóm", "zalo tham gia", "加zalo群",
    "加入zalo群数量", "vào group zalo", "vào nhóm zalo"
]

# ✅ Hàm dò keyword cho từng chỉ số KPI
def is_ket_ban_column(col):
    normalized = normalize_text(col)
    return any(keyword in normalized for keyword in KEYWORDS_KET_BAN)

def is_tuong_tac_column(col):
    normalized = normalize_text(col)
    return any(keyword in normalized for keyword in KEYWORDS_TUONG_TAC)

def is_group_zalo_column(col):
    normalized = normalize_text(col)
    return any(keyword in normalized for keyword in KEYWORDS_GROUP_ZALO)

# ✅ Hàm dò toàn bộ mapping KPI từ list cột
def detect_kpi_columns(columns):
    result = {}
    for col in columns:
        if is_ket_ban_column(col):
            result["Tổng số kết bạn trong ngày"] = col
        elif is_tuong_tac_column(col):
            result["Tương tác ≥10 câu"] = col
        elif is_group_zalo_column(col):
            result["Lượng tham gia group Zalo"] = col
    return result

# ✅ Tải file và demo kết quả dò cột
st.title("📊 Dò Cột KPI Theo Từ Khóa")
file = st.file_uploader("📤 Upload 1 file Excel", type=["xlsx"])
if file:
    xls = pd.ExcelFile(file)
    all_data = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df.shape[0] < 3:
            continue
        header_row = df.iloc[2].fillna("").astype(str)
        st.write(f"📝 Sheet: {sheet_name}")
        st.write("🎯 Tiêu đề dòng 3:", list(header_row))
        kpi_mapping = detect_kpi_columns(header_row)
        st.write("✅ Mapping cột KPI:", kpi_mapping)
else:
    st.info("📎 Vui lòng upload file Excel.")

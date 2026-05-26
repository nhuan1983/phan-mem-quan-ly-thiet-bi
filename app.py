import streamlit as st
import pandas as pd
import datetime
import json
from docx import Document
from io import BytesIO
import plotly.express as px

# 1. CẤU HÌNH TRANG VÀ ẨN GIAO DIỆN THỪA (CSS)
st.set_page_config(page_title="Hệ thống Quản lý KHTN", layout="wide")
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# 2. CẤU HÌNH KẾT NỐI GOOGLE SHEETS
USE_CLOUD_DB = False
sh = None
if "gspread_creds" in st.secrets and "spreadsheet_key" in st.secrets:
    try:
        import gspread
        creds_dict = dict(st.secrets["gspread_creds"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        gc = gspread.service_account_from_dict(creds_dict)
        sh = gc.open_by_key(st.secrets["spreadsheet_key"])
        USE_CLOUD_DB = True
    except:
        st.error("Lỗi kết nối CSDL đám mây")

def load_data(sheet_name, default_df):
    if USE_CLOUD_DB:
        try:
            df = pd.DataFrame(sh.worksheet(sheet_name).get_all_records())
            if sheet_name == 'users' and 'Vai trò' in df.columns:
                df['Vai trò'] = df['Vai trò'].apply(lambda x: [r.strip() for r in str(x).split(',')])
            return df if not df.empty else default_df
        except: return default_df
    return default_df

def save_data(sheet_name, df_to_save):
    if USE_CLOUD_DB:
        try:
            worksheet = sh.worksheet(sheet_name)
            worksheet.clear()
            df_copy = df_to_save.copy()
            if sheet_name == 'users' and 'Vai trò' in df_copy.columns:
                df_copy['Vai trò'] = df_copy['Vai trò'].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)
            for col in df_copy.columns:
                if df_copy[col].dtype == 'object': df_copy[col] = df_copy[col].astype(str)
            worksheet.update([df_copy.columns.values.tolist()] + df_copy.values.tolist())
        except: pass

# 3. KHỞI TẠO DỮ LIỆU
if 'school_info' not in st.session_state:
    st.session_state.school_info = load_data('school_info', pd.DataFrame([{'ten_truong': 'Nam Thượng', 'don_vi_chu_quan': 'UBND xã', 'nam_hoc': '2025-2026'}])).iloc[0].to_dict()
if 'users' not in st.session_state:
    st.session_state.users = load_data('users', pd.DataFrame({'Tài khoản': ['admin'], 'Mật khẩu': ['123'], 'Họ tên': ['Admin'], 'Vai trò': [['Quản trị viên']]}))
if 'chemicals' not in st.session_state:
    st.session_state.chemicals = load_data('chemicals', pd.DataFrame(columns=['Mã vật tư', 'Tên vật tư', 'Phân môn', 'Số lượng', 'Hạn sử dụng', 'Tình trạng']))
if 'bookings' not in st.session_state:
    st.session_state.bookings = load_data('bookings', pd.DataFrame(columns=['Người đăng ký', 'Ngày', 'Buổi', 'Tiết', 'Lớp', 'Môn', 'Thiết bị']))
if 'evaluations' not in st.session_state:
    df_eval = load_data('evaluations', pd.DataFrame())
    st.session_state.evaluations = df_eval.to_dict('records') if not df_eval.empty else []

if 'logged_in' not in st.session_state: st.session_state.logged_in = False

# 4. GIAO DIỆN ĐĂNG NHẬP
if not st.session_state.logged_in:
    st.title("ĐĂNG NHẬP HỆ THỐNG")
    with st.form("login"):
        user = st.text_input("Tài khoản")
        pwd = st.text_input("Mật khẩu", type="password")
        if st.form_submit_button("Đăng nhập"):
            user_match = st.session_state.users[(st.session_state.users['Tài khoản'].astype(str) == str(user)) & (st.session_state.users['Mật khẩu'].astype(str) == str(pwd))]
            if not user_match.empty:
                st.session_state.logged_in = True
                st.session_state.current_user = user_match.iloc[0].to_dict()
                st.rerun()
            else: st.error("Sai thông tin!")
    st.stop()

# 5. SIDEBAR VÀ MENU
current_user = st.session_state.current_user
active_role = st.sidebar.selectbox("Chức vụ:", current_user['Vai trò'])
menu = st.sidebar.radio("📌 Chức năng:", ["Trang chủ & Cảnh báo", "Quản lý Kho (Vật tư)", "Đăng ký thiết bị", "Đánh giá chuyên môn", "Xuất báo cáo (.docx)", "Đổi mật khẩu", "Quản lý Hệ thống (Admin)"])

if st.sidebar.button("Đăng xuất"):
    st.session_state.logged_in = False
    st.rerun()

# 6. ĐIỀU HƯỚNG CÁC MODULE (Cấu trúc if-elif chuẩn)
if menu == "Trang chủ & Cảnh báo":
    st.header("📊 Dashboard")
    # ... (Chèn code Dashboard đã làm ở bước trước) ...
    
elif menu == "Quản lý Kho (Vật tư)":
    st.header("📦 Kho vật tư")
    # ... (Chèn code Quản lý kho) ...

elif menu == "Đăng ký thiết bị":
    st.header("📝 Đăng ký phòng")
    # ... (Chèn code Đăng ký) ...

elif menu == "Đánh giá chuyên môn":
    if active_role not in ["Quản trị viên", "Hiệu trưởng", "Phó Hiệu trưởng", "Tổ trưởng chuyên môn"]:
        st.error("Bạn không có quyền!")
        st.stop()
    st.header("📋 Đánh giá")
    # ... (Chèn code Đánh giá) ...

elif menu == "Xuất báo cáo (.docx)":
    if active_role not in ["Quản trị viên", "Hiệu trưởng", "Phó Hiệu trưởng", "Tổ trưởng chuyên môn"]:
        st.error("Bạn không có quyền!")
        st.stop()
    # ... (Chèn code Xuất báo cáo) ...

elif menu == "Đổi mật khẩu":
    st.header("🔑 Đổi mật khẩu")
    with st.form("pass"):
        o, n, r = st.text_input("Cũ", type="password"), st.text_input("Mới", type="password"), st.text_input("Nhập lại", type="password")
        if st.form_submit_button("Xác nhận"):
            if o == current_user['Mật khẩu'] and n == r:
                idx = st.session_state.users[st.session_state.users['Tài khoản'] == current_user['Tài khoản']].index[0]
                st.session_state.users.at[idx, 'Mật khẩu'] = n
                save_data('users', st.session_state.users)
                st.success("Đổi thành công!")
            else: st.error("Lỗi thông tin!")

elif menu == "Quản lý Hệ thống (Admin)":
    if active_role != "Quản trị viên":
        st.error("Chỉ dành cho Admin!")
        st.stop()
    st.header("⚙️ Quản trị hệ thống")
    # ... (Chèn code Admin) ...

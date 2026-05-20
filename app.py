import streamlit as st
import pandas as pd
import datetime
from docx import Document
from io import BytesIO

# ==========================================
# CẤU HÌNH TRANG & KHỞI TẠO CƠ SỞ DỮ LIỆU
# ==========================================
# Khởi tạo thông tin trường học (Mặc định)
if 'school_info' not in st.session_state:
    st.session_state.school_info = {
        'ten_truong': 'TH&THCS Nam Thượng',
        'phong_gd': 'Phòng GD&ĐT Huyện...',
        'nam_hoc': '2025-2026'
    }

st.set_page_config(page_title=f"Quản lý KHTN - {st.session_state.school_info['ten_truong']}", layout="wide")

if 'users' not in st.session_state:
    st.session_state.users = pd.DataFrame({
        'Tài khoản': ['admin', 'ht', 'totruong', 'gv01'],
        'Mật khẩu': ['123', '123', '123', '123'],
        'Họ tên': ['Quản trị viên (PHT)', 'Nguyễn Văn A (Hiệu trưởng)', 'Trần Thị B (Tổ trưởng)', 'Lê Văn C (Giáo viên)'],
        'Vai trò': [
            ['Quản trị viên', 'Phó Hiệu trưởng', 'Giáo viên bộ môn'],
            ['Hiệu trưởng', 'Giáo viên bộ môn'],
            ['Tổ trưởng chuyên môn', 'Giáo viên bộ môn'],
            ['Giáo viên bộ môn']
        ]
    })

if 'chemicals' not in st.session_state:
    st.session_state.chemicals = pd.DataFrame({
        'Mã vật tư': ['HC01', 'HC02', 'VL01', 'SH01'],
        'Tên vật tư': ['Axit Sunfuric (H2SO4)', 'Natri Hidroxit (NaOH)', 'Bộ Khúc xạ ánh sáng', 'Tiêu bản tế bào'],
        'Phân môn': ['Hóa học', 'Hóa học', 'Vật lý', 'Sinh học'],
        'Hạn sử dụng': [datetime.date(2026, 6, 15), datetime.date(2026, 4, 10), None, datetime.date(2027, 1, 1)],
        'Tình trạng': ['Tốt', 'Sắp hết hạn', 'Tốt', 'Tốt']
    })

if 'bookings' not in st.session_state:
    st.session_state.bookings = pd.DataFrame(columns=['Người đăng ký', 'Ngày', 'Tiết', 'Lớp', 'Môn', 'Thiết bị'])

if 'evaluations' not in st.session_state:
    st.session_state.evaluations = []

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'current_user' not in st.session_state:
    st.session_state.current_user = None

# ==========================================
# GIAO DIỆN ĐĂNG NHẬP
# ==========================================
if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center; color: #1E88E5;'>HỆ THỐNG QUẢN LÝ KHTN</h1>", unsafe_allow_html=True)
    st.markdown(f"<h3 style='text-align: center;'>Trường {st.session_state.school_info['ten_truong']}</h3>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        with st.form("login_form"):
            st.write("Vui lòng đăng nhập để tiếp tục")
            username = st.text_input("Tài khoản")
            password = st.text_input("Mật khẩu", type="password")
            submit_login = st.form_submit_button("Đăng nhập", use_container_width=True)
            
            if submit_login:
                user_match = st.session_state.users[(st.session_state.users['Tài khoản'] == username) & (st.session_state.users['Mật khẩu'] == password)]
                if not user_match.empty:
                    st.session_state.logged_in = True
                    st.session_state.current_user = user_match.iloc[0].to_dict()
                    st.rerun()
                else:
                    st.error("Sai tài khoản hoặc mật khẩu!")
    st.stop()

# ==========================================
# SIDEBAR: THANH ĐIỀU HƯỚNG & CHUYỂN ĐỔI VAI TRÒ
# ==========================================
current_user = st.session_state.current_user
user_roles = current_user['Vai trò']

st.sidebar.title(st.session_state.school_info['ten_truong'])
st.sidebar.success(f"👤 Chào, {current_user['Họ tên']}")

st.sidebar.markdown("---")
active_role = st.sidebar.selectbox("🔄 Bạn đang làm việc với tư cách là:", user_roles)
st.sidebar.markdown("---")

menu_options = ["Trang chủ & Cảnh báo", "Quản lý Kho (Vật tư)", "Đăng ký thiết bị"]

if active_role in ["Quản trị viên", "Hiệu trưởng", "Phó Hiệu trưởng", "Tổ trưởng chuyên môn"]:
    menu_options.append("Đánh giá chuyên môn")
    menu_options.append("Xuất báo cáo (.docx)")

if active_role == "Quản trị viên":
    menu_options.insert(0, "Quản lý Hệ thống (Admin)")

menu = st.sidebar.radio("📌 Chọn chức năng:", menu_options)

if st.sidebar.button("Đăng xuất"):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.rerun()

# ==========================================
# MODULE: QUẢN LÝ HỆ THỐNG (Chỉ Admin)
# ==========================================
if menu == "Quản lý Hệ thống (Admin)":
    st.header("⚙️ Quản lý Hệ thống & Cấu hình Đơn vị")
    
    # --- TÍNH NĂNG MỚI: CẤU HÌNH TRƯỜNG HỌC ---
    st.subheader("1. 🏫 Cấu hình thông tin Trường học (White-label)")
    st.info("Chỉnh sửa thông tin tại đây sẽ thay đổi toàn bộ tên trường trên giao diện và trong các mẫu báo cáo xuất ra.")
    with st.form("school_config_form"):
        sc_col1, sc_col2, sc_col3 = st.columns(3)
        edit_ten_truong = sc_col1.text_input("Tên trường", value=st.session_state.school_info['ten_truong'])
        edit_phong_gd = sc_col2.text_input("Đơn vị chủ quản (Phòng GD)", value=st.session_state.school_info['phong_gd'])
        edit_nam_hoc = sc_col3.text_input("Năm học", value=st.session_state.school_info['nam_hoc'])
        
        if st.form_submit_button("💾 Lưu cấu hình đơn vị"):
            st.session_state.school_info['ten_truong'] = edit_ten_truong
            st.session_state.school_info['phong_gd'] = edit_phong_gd
            st.session_state.school_info['nam_hoc'] = edit_nam_hoc
            st.success("Đã cập nhật thông tin đơn vị thành công!")
            st.rerun()

    st.markdown("---")
    st.subheader("2. Danh sách tài khoản hiện tại")
    df_display = st.session_state.users.copy()
    df_display['Vai trò'] = df_display['Vai trò'].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)
    st.dataframe(df_display, use_

import streamlit as st
import pandas as pd
import datetime
from docx import Document
from io import BytesIO

# ==========================================
# CẤU HÌNH TRANG & KHỞI TẠO CƠ SỞ DỮ LIỆU
# ==========================================
st.set_page_config(page_title="Quản lý KHTN - TH&THCS Nam Thượng", layout="wide")

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
    st.markdown("<h3 style='text-align: center;'>Trường TH&THCS Nam Thượng</h3>", unsafe_allow_html=True)
    
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

st.sidebar.title("TH&THCS Nam Thượng")
st.sidebar.success(f"👤 Chào, {current_user['Họ tên']}")

st.sidebar.markdown("---")
active_role = st.sidebar.selectbox("🔄 Bạn đang làm việc với tư cách là:", user_roles)
st.sidebar.markdown("---")

# Phân quyền hiển thị Menu
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
    st.header("⚙️ Quản lý Hệ thống & Cấp tài khoản")
    
    st.subheader("1. Danh sách tài khoản hiện tại")
    df_display = st.session_state.users.copy()
    df_display['Vai trò'] = df_display['Vai trò'].apply(lambda x: ", ".join(x))
    st.dataframe(df_display, use_container_width=True)
    
    st.subheader("2. Cấp tài khoản mới")
    with st.form("add_user_form"):
        col1, col2 = st.columns(2)
        with col1:
            new_acc = st.text_input("Tên đăng nhập")
            new_pwd = st.text_input("Mật khẩu")
        with col2:
            new_name = st.text_input("Họ và tên người dùng")
            new_roles = st.multiselect("Chọn các vai trò (Có thể chọn nhiều)", 
                                       ["Giáo viên bộ môn", "Tổ trưởng chuyên môn", "Phó Hiệu trưởng", "Hiệu trưởng", "Quản trị viên"])
        
        if st.form_submit_button("Tạo tài khoản"):
            if new_acc and new_pwd and new_name and len(new_roles) > 0:
                new_row = pd.DataFrame([{'Tài khoản': new_acc, 'Mật khẩu': new_pwd, 'Họ tên': new_name, 'Vai trò': new_roles}])
                st.session_state.users = pd.concat([st.session_state.users, new_row], ignore_index=True)
                st.success(f"Đã cấp tài khoản thành công cho: {new_name}")
                st.rerun()
            else:
                st.warning("Vui lòng điền đủ thông tin và chọn ít nhất 1 vai trò!")

# ==========================================
# MODULE: TRANG CHỦ & CẢNH BÁO
# ==========================================
elif menu == "Trang chủ & Cảnh báo":
    st.header("📊 Bảng điều khiển (Dashboard)")
    
    col1, col2 = st.columns(2)
    col1.metric("Tổng số thiết bị/Hóa chất trong kho", len(st.session_state.chemicals))
    col2.metric("Số lượt đăng ký mượn phòng", len(st.session_state.bookings))
    
    today = datetime.date.today()
    df_exp = st.session_state.chemicals.dropna(subset=['Hạn sử dụng'])
    df_warning = df_exp[(df_exp['Hạn sử dụng'] - today).dt.days <= 30]
    
    st.subheader("⚠️ Cảnh báo an toàn (Hóa chất/Mẫu vật)")
    if not df_warning.empty:
        st.error("Phát hiện vật tư sắp hết hạn hoặc đã hết hạn!")
        st.dataframe(df_warning, use_container_width=True)
    else:
        st.success("Tất cả hóa chất và tiêu bản đều trong thời hạn sử dụng an toàn.")

# ==========================================
# MODULE: QUẢN LÝ KHO (THÊM THIẾT BỊ)
# ==========================================
elif menu == "Quản lý Kho (Vật tư)":
    st.header("📦 Quản lý Kho Thiết bị & Hóa chất")
    
    st.subheader("Danh mục vật tư hiện có")
    st.dataframe(st.session_state.chemicals, use_container_width=True)
    
    # Form thêm vật tư: Mở cho Quản trị viên, Tổ trưởng (và Ban giám hiệu)
    if active_role in ["Quản trị viên", "Tổ trưởng chuyên môn", "Phó Hiệu trưởng", "Hiệu trưởng"]:
        st.markdown("---")
        st.subheader("➕ Bổ sung vật tư/thiết bị mới")
        with st.form("add_chem_form"):
            c1, c2, c3 = st.columns(3)
            ma_vt = c1.text_input("Mã vật tư (VD: VL05)")
            ten_vt = c2.text_input("Tên vật tư")
            phan_mon = c3.selectbox("Phân môn", ["Vật lý", "Hóa học", "Sinh học"])
            
            c4, c5 = st.columns(2)
            han_su_dung = c4.date_input("Hạn sử dụng (Nếu không có, để nguyên)", value=None)
            tinh_trang = c5.selectbox("Tình trạng", ["Tốt", "Cần sửa chữa", "Đang đặt mua"])
            
            if st.form_submit_button("Lưu vào kho"):
                if ma_vt and ten_vt:
                    new_item = pd.DataFrame([{'Mã vật tư': ma_vt, 'Tên vật tư': ten_vt, 'Phân môn': phan_mon, 'Hạn sử dụng': han_su_dung, 'Tình trạng': tinh_trang}])
                    st.session_state.chemicals = pd.concat([st.session_state.chemicals, new_item], ignore_index=True)
                    st.success(f"Đã bổ sung thành công [{ten_vt}] vào cơ sở dữ liệu!")
                    st.rerun()
                else:
                    st.warning("Vui lòng nhập Mã vật tư và Tên vật tư!")
    else:
        st.info("Chỉ Quản trị viên và Tổ chuyên môn mới có quyền bổ sung thiết bị mới.")

# ==========================================
# MODULE: ĐĂNG KÝ THIẾT BỊ
# ==========================================
elif menu == "Đăng ký thiết bị":
    st.header("📝 Đăng ký sử dụng phòng bộ môn")
    
    st.subheader("Lịch đăng ký của toàn trường")
    if not st.session_state.bookings.empty:
        st.dataframe(st.session_state.bookings, use_container_width=True)
    else:
        st.info("Chưa có lịch đăng ký nào.")
        
    st.subheader("Tạo đăng ký mới")
    with st.form("booking_form"):
        col1, col2 = st.columns(2)
        with col1:
            date = st.date_input("Ngày dạy")
            period = st.selectbox("Tiết học", [1, 2, 3, 4, 5])
            lop = st.text_input("Lớp (VD: 9A)")
            subject = st.selectbox("Phân môn", ["Vật lý", "Hóa học", "Sinh học"])
        with col2:
            equipment = st.multiselect("Chọn thiết bị/Hóa chất cần mượn", st.session_state.chemicals['Tên vật tư'].tolist())
            
        if st.form_submit_button("Xác nhận đăng ký"):
            if lop:
                new_book = pd.DataFrame([{'Người đăng ký': current_user['Họ tên'], 'Ngày': date, 'Tiết': period, 'Lớp': lop, 'Môn': subject, 'Thiết bị': ", ".join(equipment)}])
                st.session_state.bookings = pd.concat([st.session_state.bookings, new_book], ignore_index=True)
                st.success("Đã ghi nhận lịch đăng ký!")
                st.rerun()
            else:
                st.warning("Vui lòng nhập tên Lớp!")

# ==========================================
# MODULE: ĐÁNH GIÁ CHUYÊN MÔN
# ==========================================
elif menu == "Đánh giá chuyên môn":
    st.header("📋 Đánh giá năng lực tổ chức thực hành")
    
    list_gv = st.session_state.users['Họ tên'].tolist()
    target_gv = st.selectbox("1. Chọn Giáo viên để đánh giá:", ["-- Chọn người được đánh giá --"] + list_gv)
    
    if target_gv != "-- Chọn người được đánh giá --":
        gv_bookings = st.session_state.bookings[st.session_state.bookings['Người đăng ký'] == target_gv]
        
        if gv_bookings.empty:
            st.warning(f"{target_gv} chưa đăng ký tiết dạy thực hành nào trên hệ thống.")
        else:
            booking_options = []
            for idx, row in gv_bookings.iterrows():
                booking_options.append(f"Ngày {row['Ngày']} - Tiết {row['Tiết']} - Lớp {row['Lớp']} - Môn {row['Môn']}")
            
            target_tiet = st.selectbox("2. Chọn Tiết dạy để đánh giá:", booking_options)
            
            st.markdown("### 3. Chấm điểm Rubric")
            c1 = st.slider("1. Công tác chuẩn bị thiết bị, vật tư (Tối đa 20đ)", 0, 20, 15)
            c2 = st.slider("2. Đảm bảo quy tắc an toàn PTN (Tối đa 30đ)", 0, 30, 25)
            c3 = st.slider("3. Tổ chức và hướng dẫn HS thao tác (Tối đa 30đ)", 0, 30, 25)
            c4 = st.slider("4. Đánh giá kết quả & Liên hệ PISA (Tối đa 20đ)", 0, 20, 15)
            
            total = c1 + c2 + c3 + c4
            st.markdown(f"**Tổng điểm:** {total}/100")
            
            if total >= 85: rank = "Tốt"
            elif total >= 65: rank = "Khá"
            elif total >= 50: rank = "Đạt"
            else: rank = "Chưa đạt"
            
            comment = st.text_area("Nhận xét ưu điểm & Tồn tại:")
            
            if st.button("Lưu đánh giá"):
                record = {
                    "Người được đánh giá": target_gv,
                    "Tiết dạy": target_tiet,
                    "Người đánh giá": current_user['Họ tên'],
                    "Chức vụ người đánh giá": active_role,
                    "Tổng điểm": total,
                    "Xếp loại": rank,
                    "Nhận xét": comment,
                    "Ngày chấm": datetime.date.today().strftime("%d/%m/%Y")
                }
                st.session_state.evaluations.append(record)
                st.success("Đã lưu hồ sơ đánh giá vào hệ thống!")

# ==========================================
# MODULE: XUẤT BÁO CÁO TỰ ĐỘNG (.DOCX)
# ==========================================
elif menu == "Xuất báo cáo (.docx)":
    st.header("🖨️ Kết xuất hồ sơ minh chứng")
    
    if len(st.session_state.evaluations) == 0:
        st.info("Chưa có hồ sơ đánh giá nào trong hệ thống.")
    else:
        df_evals = pd.DataFrame(st.session_state.evaluations)
        st.dataframe(df_evals, use_container_width=True)
        
        selected_idx = st.selectbox("Chọn hồ sơ cần xuất báo cáo", range(len(st.session_state.evaluations)), 
                                    format_func=lambda x: f"Đánh giá: {st.session_state.evaluations[x]['Người được đánh giá']} ({st.session_state.evaluations[x]['Tiết dạy']})")
        
        target_record = st.session_state.evaluations[selected_idx]
        
        def create_docx(data):
            doc = Document()
            doc.add_heading('PHIẾU ĐÁNH GIÁ NĂNG LỰC THỰC HÀNH', 0)
            doc.add_paragraph(f"Trường: TH&THCS Nam Thượng")
            doc.add_paragraph(f"Ngày đánh giá: {data['Ngày chấm']}")
            doc.add_paragraph(f"Người đánh giá: {data['Người đánh giá']} ({data['Chức vụ người đánh giá']})")
            
            doc.add_heading('Thông tin tiết dạy', level=1)
            doc.add_paragraph(f"- Giáo viên dạy: {data['Người được đánh giá']}")
            doc.add_paragraph(f"- Thông tin tiết: {data['Tiết dạy']}")
            
            doc.add_heading('Kết quả', level=1)
            doc.add_paragraph(f"- Tổng điểm: {data['Tổng điểm']}/100")
            doc.add_paragraph(f"- Xếp loại: {data['Xếp loại']}")
            
            doc.add_heading('Nhận xét chuyên môn', level=1)
            doc.add_paragraph(data['Nhận xét'])
            
            bio = BytesIO()
            doc.save(bio)
            return bio.getvalue()

        docx_file = create_docx(target_record)
        st.download_button(
            label="📄 Tải xuống báo cáo file Word",
            data=docx_file,
            file_name=f"DanhGia_{target_record['Người được đánh giá']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

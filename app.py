import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ==========================================
# CẤU HÌNH KẾT NỐI GOOGLE SHEETS
# ==========================================
USE_CLOUD_DB = False
conn = None
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
    except Exception as e:
        st.sidebar.error(f"⚠️ Lỗi kết nối Google Sheets: {e}")

# --- CÁC HÀM ĐỌC/GHI DỮ LIỆU ---
def load_data(sheet_name, default_df):
    if USE_CLOUD_DB:
        try:
            worksheet = sh.worksheet(sheet_name)
            records = worksheet.get_all_records()
            if not records:
                return default_df
            df = pd.DataFrame(records)
            if sheet_name == 'users' and 'Vai trò' in df.columns:
                df['Vai trò'] = df['Vai trò'].apply(lambda x: [r.strip() for r in str(x).split(',')])
            return df
        except Exception:
            return default_df
    return default_df

def save_data(sheet_name, df_to_save):
    if USE_CLOUD_DB:
        try:
            try:
                worksheet = sh.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = sh.add_worksheet(title=sheet_name, rows="100", cols="20")
            
            worksheet.clear()
            df_copy = df_to_save.copy()
            if sheet_name == 'users' and 'Vai trò' in df_copy.columns:
                df_copy['Vai trò'] = df_copy['Vai trò'].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)
            
            for col in df_copy.columns:
                if df_copy[col].dtype == 'object':
                    df_copy[col] = df_copy[col].astype(str)
                    
            # Cập nhật an toàn với parameters chuẩn
            worksheet.update(values=[df_copy.columns.values.tolist()] + df_copy.values.tolist(), range_name="A1")
        except Exception as e:
            st.error(f"Không thể đồng bộ lên Cloud: {e}")

# ==========================================
# KHỞI TẠO HOẶC TẢI CƠ SỞ DỮ LIỆU
# ==========================================
default_school = pd.DataFrame([{'ten_truong': 'TH&THCS Nam Thượng', 'don_vi_chu_quan': 'Ủy ban nhân dân huyện', 'nam_hoc': '2025-2026'}])
df_school_db = load_data('school_info', default_school)
if 'school_info' not in st.session_state:
    st.session_state.school_info = df_school_db.iloc[0].to_dict()

default_users = pd.DataFrame({
    'Tài khoản': ['admin', 'ht', 'totruong', 'gv01'],
    'Mật khẩu': ['123', '123', '123', '123'],
    'Họ tên': ['Quản trị viên (PHT)', 'Hiệu trưởng', 'Tổ trưởng KHTN', 'Giáo viên Vật Lý'],
    'Vai trò': [['Quản trị viên', 'Phó Hiệu trưởng', 'Giáo viên bộ môn'], ['Hiệu trưởng', 'Giáo viên bộ môn'], ['Tổ trưởng chuyên môn', 'Giáo viên bộ môn'], ['Giáo viên bộ môn']]
})
if 'users' not in st.session_state:
    st.session_state.users = load_data('users', default_users)

# ĐÃ BỔ SUNG CỘT ĐƠN VỊ VÀO KHỞI TẠO
default_chem = pd.DataFrame({
    'Mã vật tư': ['HC01', 'HC02', 'VL01', 'SH01', 'DC01'],
    'Tên vật tư': ['Axit Sunfuric (H2SO4)', 'Natri Hidroxit (NaOH)', 'Bộ Khúc xạ ánh sáng', 'Tiêu bản tế bào', 'Kính hiển vi quang học'],
    'Phân môn': ['Hóa học', 'Hóa học', 'Vật lý', 'Sinh học', 'Dùng chung'],
    'Số lượng': [5, 10, 3, 20, 5],
    'Đơn vị': ['ml', 'gam', 'bộ', 'cái', 'cái'], 
    'Hạn sử dụng': ['15/06/2026', '10/04/2026', '', '01/01/2027', ''],
    'Tình trạng': ['Tốt', 'Sắp hết hạn', 'Tốt', 'Tốt', 'Tốt']
})
if 'chemicals' not in st.session_state:
    st.session_state.chemicals = load_data('chemicals', default_chem)

if 'bookings' not in st.session_state:
    st.session_state.bookings = load_data('bookings', pd.DataFrame(columns=['Người đăng ký', 'Ngày', 'Buổi', 'Tiết', 'Lớp', 'Môn', 'Thiết bị']))

if 'evaluations' not in st.session_state:
    df_eval_loaded = load_data('evaluations', pd.DataFrame())
    st.session_state.evaluations = df_eval_loaded.to_dict('records') if not df_eval_loaded.empty else []

if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_user' not in st.session_state: st.session_state.current_user = None

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
st.set_page_config(page_title="Phần mềm Quản lý Thiết bị", layout="wide")
st.title(f"HỆ THỐNG QUẢN LÝ THIẾT BỊ DẠY HỌC - TRƯỜNG {st.session_state.school_info['ten_truong'].upper()}")

# --- ĐĂNG NHẬP ---
if not st.session_state.logged_in:
    st.subheader("Đăng nhập hệ thống")
    tk = st.text_input("Tài khoản")
    mk = st.text_input("Mật khẩu", type="password")
    if st.button("Đăng nhập"):
        user_match = st.session_state.users[(st.session_state.users['Tài khoản'].astype(str) == str(tk)) & (st.session_state.users['Mật khẩu'].astype(str) == str(mk))]
        if not user_match.empty:
            st.session_state.logged_in = True
            st.session_state.current_user = user_match.iloc[0].to_dict()
            st.success("Đăng nhập thành công!")
            st.rerun()
        else:
            st.error("Sai tài khoản hoặc mật khẩu!")
    st.stop()

# --- MENU ĐIỀU HƯỚNG ---
st.sidebar.header(f"Xin chào, {st.session_state.current_user['Họ tên']}")
if st.sidebar.button("Đăng xuất"):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.rerun()

active_role = st.session_state.current_user['Vai trò'][0] if isinstance(st.session_state.current_user['Vai trò'], list) else st.session_state.current_user['Vai trò']
menu = st.sidebar.radio("Chọn chức năng:", ["Quản lý Kho Thiết bị", "Quản lý Tài khoản"])

# --- CHỨC NĂNG 1: QUẢN LÝ KHO THIẾT BỊ ---
if menu == "Quản lý Kho Thiết bị":
    st.header("Danh mục Thiết bị / Vật tư")
    
    # Hiển thị bảng dữ liệu
    st.dataframe(st.session_state.chemicals, use_container_width=True)
    
    # Nút xuất Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.chemicals.to_excel(writer, index=False)
    st.download_button(
        label="Tải danh sách về máy (.xlsx)",
        data=output.getvalue(),
        file_name="Danh_muc_thiet_bi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.markdown("---")
    
    # Form thêm thiết bị mới
    st.subheader("Thêm thiết bị mới vào kho")
    with st.form("add_item_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            ma_vt = st.text_input("Mã vật tư")
            ten_vt = st.text_input("Tên vật tư")
        with col2:
            phan_mon = st.selectbox("Phân môn", ["Vật lý", "Hóa học", "Sinh học", "Dùng chung"])
            so_luong = st.number_input("Số lượng", min_value=1, step=1)
        with col3:
            don_vi = st.text_input("Đơn vị (cái, bộ, ml, gam...)") # Ô nhập đơn vị
            han_su_dung = st.date_input("Hạn sử dụng", value=None)
        
        tinh_trang = st.selectbox("Tình trạng", ["Tốt", "Sắp hết hạn", "Hỏng", "Cần thanh lý"])
        
        if st.form_submit_button("Lưu vào kho"):
            if ma_vt and ten_vt and don_vi:
                han_str = han_su_dung.strftime('%d/%m/%Y') if han_su_dung else ""
                # BỔ SUNG CỘT ĐƠN VỊ VÀO CHUỖI LƯU THỦ CÔNG
                new_item = pd.DataFrame([{'Mã vật tư': ma_vt, 'Tên vật tư': ten_vt, 'Phân môn': phan_mon, 'Số lượng': int(so_luong), 'Đơn vị': don_vi, 'Hạn sử dụng': han_str, 'Tình trạng': tinh_trang}])
                st.session_state.chemicals = pd.concat([st.session_state.chemicals, new_item], ignore_index=True)
                
                # Lưu đồng bộ lên Google Sheets
                save_data('chemicals', st.session_state.chemicals)
                st.success("Đã bổ sung thiết bị vào kho!")
                st.rerun()
            else:
                st.warning("Vui lòng nhập đủ Mã vật tư, Tên vật tư và Đơn vị!")
    
    st.markdown("---")
    
    # Form tải lên hàng loạt từ Excel
    st.subheader("Nhập hàng loạt từ file Excel")
    uploaded_chem = st.file_uploader("Chọn file Excel (đảm bảo cấu trúc 7 cột gồm cả Đơn vị)", type=['xlsx', 'xls'])
    if uploaded_chem is not None and st.button("Tiến hành nhập dữ liệu thiết bị"):
        try:
            df_new = pd.read_excel(uploaded_chem)
            st.session_state.chemicals = pd.concat([st.session_state.chemicals, df_new], ignore_index=True)
            
            # ĐÃ CHÈN LỆNH LƯU LÊN CLOUD
            save_data('chemicals', st.session_state.chemicals)
            st.success("Nhập thành công danh mục thiết bị hàng loạt!")
            st.rerun()
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")

# --- CHỨC NĂNG 2: QUẢN LÝ TÀI KHOẢN ---
elif menu == "Quản lý Tài khoản":
    st.header("Quản lý Tài khoản Hệ thống")
    if active_role not in ["Quản trị viên", "Hiệu trưởng", "Phó Hiệu trưởng"]:
        st.warning("Chỉ Ban Giám Hiệu mới có quyền truy cập chức năng này!")
    else:
        st.dataframe(st.session_state.users, use_container_width=True)
        
        st.subheader("Thêm tài khoản từ file Excel")
        uploaded_users = st.file_uploader("Chọn file Excel tài khoản đã điền", type=['xlsx', 'xls'])
        if uploaded_users is not None and st.button("Tiến hành nhập dữ liệu tài khoản"):
            try:
                df_new = pd.read_excel(uploaded_users)
                df_new['Vai trò'] = df_new['Vai trò'].astype(str).apply(lambda x: [r.strip() for r in x.split(',')])
                df_new['Tài khoản'] = df_new['Tài khoản'].astype(str)
                df_new['Mật khẩu'] = df_new['Mật khẩu'].astype(str)
                st.session_state.users = pd.concat([st.session_state.users, df_new], ignore_index=True)
                
                # Lưu đồng bộ lên Google Sheets
                save_data('users', st.session_state.users)
                st.success("Nhập thành công tài khoản hàng loạt!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi: {e}")

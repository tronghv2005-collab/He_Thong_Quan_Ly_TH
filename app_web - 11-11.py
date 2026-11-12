import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import calendar
import os
import hashlib
import plotly.express as px
import plotly.graph_objects as go

# Import các module báo cáo
try:
    import BC_th as bc
except ImportError:
    bc = None

try:
    import BC_NHNN
except ImportError:
    BC_NHNN = None

try:
    import BC_ATM
except ImportError:
    BC_ATM = None

try:
    import giao_dich_lon
except ImportError:
    giao_dich_lon = None

try:
    import bc_xnk
except ImportError:
    bc_xnk = None

try:
    import nhapdknoplinh
except ImportError:
    nhapdknoplinh = None

try:
    import quanlyloi
except ImportError:
    quanlyloi = None

try:
    from sla_report import show_sla_report_page
except ImportError:
    show_sla_report_page = None

try:
    import kpi_module
except ImportError:
    kpi_module = None

# ==================== ĐỊNH NGHĨA THƯ MỤC DÙNG CHUNG ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SHARED_DATA_DIR = os.path.join(BASE_DIR, "data_source")
OUTPUT_DIR = os.path.join(BASE_DIR, "templates")

os.makedirs(SHARED_DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==================== HỆ THỐNG XÁC THỰC & BẢO MẬT ĐĂNG NHẬP (CHUẨN MaNV) ====================
def hash_password(password):
    """Hàm băm mật khẩu bằng SHA-256"""
    return hashlib.sha256(str(password).encode('utf-8')).hexdigest()

def check_login_system(base_dir):
    param_dir = os.path.join(base_dir, "parameters")
    path_canbo = os.path.join(param_dir, "Thamso_canbo.xlsx")
    os.makedirs(param_dir, exist_ok=True)

    # Khởi tạo file mẫu cán bộ nếu chưa có
    if not os.path.exists(path_canbo):
        try:
            df_init_cb = pd.DataFrame([
                {'MaNV': '12345', 'MatKhau': hash_password('00'), 'TrangThai': True, 'TenNV': 'Nguyễn Văn A', 'NhomQuyen': 'KSV', 'CoSo': 'Cơ sở 1'},
                {'MaNV': '67890', 'MatKhau': hash_password('00'), 'TrangThai': True, 'TenNV': 'Trần Thị B', 'NhomQuyen': 'NV', 'CoSo': 'Cơ sở 2'}
            ])
            df_init_cb['MaNV'] = df_init_cb['MaNV'].astype(str)
            df_init_cb['MatKhau'] = df_init_cb['MatKhau'].astype(str)
            df_init_cb.to_excel(path_canbo, index=False)
        except:
            pass

    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'current_user' not in st.session_state:
        st.session_state.current_user = None
    if 'current_manv' not in st.session_state:
        st.session_state.current_manv = None
    if 'force_change_pass' not in st.session_state:
        st.session_state.force_change_pass = False

    if not st.session_state.logged_in:
        st.set_page_config(
            page_title="Đăng Nhập - Hệ Thống Báo Cáo Trung Tâm",
            page_icon="🔐",
            layout="centered"
        )
        st.markdown("### 🔐 ĐĂNG NHẬP HỆ THỐNG TRUNG TÂM TIẾP QUỸ & NHNN")
        
        with st.form("login_form_main"):
            username = st.text_input("Mã cán bộ (MaNV)")
            password = st.text_input("Mật khẩu", type="password")
            submit_login = st.form_submit_button("Đăng nhập", type="primary")

            if submit_login:
                try:
                    df_cb = pd.read_excel(path_canbo, dtype={'MaNV': str, 'MatKhau': str})
                    df_cb['MaNV'] = df_cb['MaNV'].str.strip()
                    df_cb['MatKhau'] = df_cb['MatKhau'].str.strip()
                    
                    user_row = df_cb[df_cb['MaNV'] == username.strip()]
                    
                    if not user_row.empty:
                        db_pass = str(user_row.iloc[0]['MatKhau']).strip()
                        
                        raw_trang_thai = user_row.iloc[0]['TrangThai']
                        if isinstance(raw_trang_thai, str):
                            is_first_login = raw_trang_thai.strip().lower() in ['true', '1', 'yes']
                        else:
                            is_first_login = bool(raw_trang_thai)
                        
                        input_pass_clean = password.strip()
                        input_check = input_pass_clean if (is_first_login and (db_pass == input_pass_clean or db_pass == hash_password(input_pass_clean))) else hash_password(input_pass_clean)
                        
                        if db_pass == input_check or (is_first_login and db_pass == input_pass_clean):
                            col_ten = next((c for c in df_cb.columns if 'ten' in str(c).lower() or 'tên' in str(c).lower()), None)
                            ten_nv = str(user_row.iloc[0][col_ten]).strip() if (col_ten and pd.notnull(user_row.iloc[0][col_ten])) else "Cán bộ"
                            
                            st.session_state.current_manv = username.strip()
                            st.session_state.current_user = f"{username.strip()} - {ten_nv}"
                            st.session_state.logged_in = True
                            st.session_state.force_change_pass = is_first_login
                            st.rerun()
                        else:
                            st.error("❌ Mật khẩu không chính xác!")
                    else:
                        st.error("❌ Không tìm thấy mã cán bộ trong hệ thống!")
                except Exception as e:
                    st.error(f"❌ Lỗi đọc file danh sách cán bộ: {e}")
        return False

    # Bắt buộc đổi mật khẩu trong lần đầu đăng nhập (Không giới hạn độ dài/độ mạnh)
    if st.session_state.force_change_pass:
        st.set_page_config(
            page_title="Đổi Mật Khẩu Lần Đầu",
            page_icon="⚠️",
            layout="centered"
        )
        st.warning("⚠️ **Đăng nhập lần đầu: Vui lòng nhập mật khẩu mới để tiếp tục!**")
        
        with st.form("change_pass_form_main"):
            new_pass = st.text_input("Nhập mật khẩu mới", type="password", key="np1")
            confirm_pass = st.text_input("Xác nhận lại mật khẩu mới", type="password", key="np2")
            submit_change = st.form_submit_button("Lưu mật khẩu và Vào hệ thống", type="primary")

            if submit_change:
                if new_pass.strip() == "":
                    st.warning("⚠️ Mật khẩu không được để trống!")
                elif new_pass.strip() != confirm_pass.strip():
                    st.warning("⚠️ Mật khẩu xác nhận không khớp!")
                else:
                    try:
                        df_cb = pd.read_excel(path_canbo, dtype={'MaNV': str, 'MatKhau': str})
                        df_cb['MaNV'] = df_cb['MaNV'].str.strip()
                        df_cb['MatKhau'] = df_cb['MatKhau'].str.strip()
                        
                        target_ma = str(st.session_state.get('current_manv', '')).strip()
                        
                        idx = df_cb[df_cb['MaNV'] == target_ma].index
                        if not idx.empty:
                            df_cb.loc[idx, 'MatKhau'] = str(hash_password(new_pass.strip()))
                            df_cb.loc[idx, 'TrangThai'] = False
                            df_cb.to_excel(path_canbo, index=False)
                            
                            st.session_state.force_change_pass = False
                            st.success("🎉 Đổi mật khẩu thành công! Đang chuyển vào hệ thống...")
                            st.rerun()
                        else:
                            st.error(f"❌ Không tìm thấy dòng dữ liệu của Mã NV '{target_ma}' trong file Thamso_canbo.xlsx.")
                    except Exception as e:
                        st.error(f"❌ Lỗi cập nhật mật khẩu: {e}")
        return False

    return True

# --- THỰC THI KIỂM TRA ĐĂNG NHẬP TRƯỚC KHI HIỂN THỊ HỆ THỐNG ---
if not check_login_system(BASE_DIR):
    st.stop()

# Cấu hình giao diện chính
st.set_page_config(
    page_title="Hệ Thống Báo Cáo Trung Tâm Tiếp Quỹ & NHNN",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <style>
    .main { background-color: #FAFAFC; }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #FFF9FB 0%, #F5EFF2 100%);
        border: 1px solid #EEDCE3;
        padding: 16px;
        border-radius: 14px;
        box-shadow: 0 4px 15px rgba(210, 180, 195, 0.12);
    }
    div[data-testid="stMetric"] label { color: #7D6B75 !important; font-weight: 500; }
    h3, h4 { color: #4A3B45; font-family: 'Helvetica Neue', sans-serif; }
    .stButton > button {
        border-radius: 10px;
        border: none;
        background: linear-gradient(135deg, #D4849D 0%, #B86B83 100%);
        color: white;
        font-weight: 500;
        box-shadow: 0 3px 10px rgba(184, 107, 131, 0.25);
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #C2728B 0%, #A35870 100%);
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 HỆ THỐNG XỬ LÝ VÀ XUẤT BÁO CÁO TỰ ĐỘNG")
st.markdown("---")

today = date.today()
default_kc_start = date(today.year, today.month, 1)
default_kc_end = today

if today.month == 1:
    prev_month_year = today.year - 1
    prev_month = 12
else:
    prev_month_year = today.year
    prev_month = today.month - 1

last_day_of_prev_month = calendar.monthrange(prev_month_year, prev_month)[1]
default_kp_start = date(prev_month_year, prev_month, 1)
default_kp_end = date(prev_month_year, prev_month, last_day_of_prev_month)

# ==================== THANH ĐIỀU HƯỚNG (SIDEBAR) ====================
st.sidebar.markdown(f"👤 **Tài khoản:** {st.session_state.current_user}")
if st.sidebar.button("🚪 Đăng xuất", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.current_manv = None
    st.session_state.force_change_pass = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("⚙️ TÙY CHỌN CHỨC NĂNG")
report_type = st.sidebar.selectbox(
    "Chọn loại báo cáo cần thực hiện:", 
    (
        "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)", 
        "Báo Cáo Nộp Lĩnh NHNN (Ds_NopLinhNHNN)",
        "Báo Cáo Hoạt Động ATM (BC_atm)",
        "Giao Dịch Phát Sinh Lớn (Ds_TiepQuyThuGom)",
        "Báo Cáo XNK Theo Loại Tiền & Đối Tác",  
        "Nhập Đăng Ký Nộp Lĩnh",  
        "Quản Lý Lỗi / Vi Phạm",  
        "Báo Cáo SLA Kho Quỹ",
        "Báo Cáo KPI Cán Bộ" 
    )
)

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Thiết lập khoảng thời gian")
start_prev = st.sidebar.date_input("Từ ngày (Kỳ trước)", value=default_kp_start)
end_prev = st.sidebar.date_input("Đến ngày (Kỳ trước)", value=default_kp_end)
start_curr = st.sidebar.date_input("Từ ngày (Kỳ này)", value=default_kc_start)
end_curr = st.sidebar.date_input("Đến ngày (Kỳ này)", value=default_kc_end)

# ==================== HÀM RENDER DASHBOARD SLA ====================
def render_sla_violation_dashboard(df_master, shared_data_dir, param_dir, start_date, end_date):
    st.markdown("---")
    st.subheader("🌸 DASHBOARD TRỰC QUAN & GIÁM SÁT VI PHẠM SLA")

    @st.cache_data
    def load_mock_or_real_sla_data(df_m):
        if df_m.empty: return pd.DataFrame()
        np.random.seed(42)
        nghiep_vu_list = ['Tiếp quỹ định lượng', 'Tiếp quỹ ATM', 'Nộp lĩnh khác', 'Kiểm đếm tiền', 'Xử lý tiền mặt']
        mock_records = []
        for _, row in df_m.iterrows():
            num_tx = np.random.randint(5, 25)
            for _ in range(num_tx):
                nv = np.random.choice(nghiep_vu_list)
                is_violation = np.random.choice([0, 1], p=[0.93, 0.07])
                mock_records.append({
                    'MaSo': row.get('MaSo', 'N/A'),
                    'HoTen': row.get('HoTen', 'Nhân sự'),
                    'CoSo': row.get('CoSo', 'Cơ sở 1'),
                    'ViTri': row.get('ViTri', 'Nhân viên'),
                    'LoaiNghiepVu': nv,
                    'ViPham': is_violation 
                })
        return pd.DataFrame(mock_records)

    df_sla = load_mock_or_real_sla_data(df_master)
    if df_sla.empty: return

    col_f1, col_f2 = st.columns(2)
    with col_f1: selected_coso_sla = st.selectbox("🔍 Lọc theo Cơ sở:", ["Tất cả"] + list(df_sla['CoSo'].unique()), key="sla_filter_coso")
    with col_f2: selected_nv_sla = st.selectbox("🔍 Lọc theo Nghiệp vụ:", ["Tất cả"] + list(df_sla['LoaiNghiepVu'].unique()), key="sla_filter_nv")

    df_filtered = df_sla.copy()
    if selected_coso_sla != "Tất cả": df_filtered = df_filtered[df_filtered['CoSo'] == selected_coso_sla]
    if selected_nv_sla != "Tất cả": df_filtered = df_filtered[df_filtered['LoaiNghiepVu'] == selected_nv_sla]

    total_tx = len(df_filtered)
    total_violation = df_filtered['ViPham'].sum()
    rate_violation = (total_violation / total_tx * 100) if total_tx > 0 else 0.0

    mcol1, mcol2, mcol3 = st.columns(3)
    mcol1.metric("📌 Tổng số Giao dịch giám sát", f"{total_tx:,}")
    mcol2.metric("⚠️ Tổng số vụ Vi phạm SLA", f"{total_violation:,}", delta=f"-{rate_violation:.2f}%", delta_color="inverse")
    mcol3.metric("📈 Tỷ lệ vi phạm chung", f"{rate_violation:.2f}%")


# ==================== ĐIỀU HƯỚNG HIỂN THỊ CÁC MODULE ====================

if report_type == "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)":
    st.subheader("📋 MODULE BÁO CÁO TỔNG HỢP TIẾP QUỸ & THU GOM")
    PATH_FILE1 = os.path.join(SHARED_DATA_DIR, "Ds_TiepQuyThuGom.xlsx")
    PATH_FILE2 = os.path.join(SHARED_DATA_DIR, "Ds_Noplinhkhac.xlsx")
    PATH_PARAM_KM = os.path.join(BASE_DIR, "parameters", "Thamso_KM.xlsx")
    PATH_TEMPLATE = os.path.join(BASE_DIR, "templates", "Template_BaoCao1.xlsx")

    if st.button("⚡ Chạy và Xem Báo Cáo Tổng Hợp", type="primary"):
        if bc is None:
            st.error("❌ Không tìm thấy module `BC_th.py`.")
        else:
            with st.spinner("Đang xử lý báo cáo tổng hợp..."):
                if os.path.exists(PATH_FILE1) or os.path.exists(PATH_FILE2):
                    df1 = bc.clean_and_prepare_file1(PATH_FILE1) if os.path.exists(PATH_FILE1) else pd.DataFrame()
                    df2 = bc.clean_and_prepare_file2(PATH_FILE2) if os.path.exists(PATH_FILE2) else pd.DataFrame()
                    df_all = pd.concat([df1, df2], ignore_index=True)
                    df_param_km = pd.read_excel(PATH_PARAM_KM) if os.path.exists(PATH_PARAM_KM) else pd.DataFrame()
                    
                    all_coso = ["Cơ sở 1", "Cơ sở 2", "Hub NVC", "CS HĐ", "CS HP", "Toàn trung tâm"]
                    all_nghiep_vu = [nv for nv in df_all['LoaiNghiepVu'].dropna().unique() if nv != 'Bỏ qua'] + ['Tổng Thu', 'Tổng Chi']
                    
                    df_prev_res = bc.aggregate_period(df_all, start_prev, end_prev, df_param_km)
                    df_curr_res = bc.aggregate_period(df_all, start_curr, end_curr, df_param_km)
                    final_table = bc.format_summary_dataframe(df_prev_res, df_curr_res, all_nghiep_vu, all_coso)
                    
                    st.dataframe(final_table, use_container_width=True)

elif report_type == "Báo Cáo Nộp Lĩnh NHNN (Ds_NopLinhNHNN)":
    st.subheader("🏦 MODULE: DANH SÁCH NỘP LĨNH NHNN")
    thang_nam_nhnn = start_curr.strftime("%m%Y")
    path_output_nhnn = os.path.join(OUTPUT_DIR, f"Bao_Cao_DS_NopLinh_NHNN_{thang_nam_nhnn}.xlsx")
    
    if st.button("🚀 Chạy và Xem Báo Cáo NHNN", type="primary"):
        if BC_NHNN is None:
            st.error("❌ Không tìm thấy module `BC_NHNN.py`.")
        else:
            success = BC_NHNN.process_ds_nop_linh_nhnn_pipeline(SHARED_DATA_DIR, start_curr, end_curr, start_prev, end_prev, path_output_nhnn)
            if success and os.path.exists(path_output_nhnn):
                st.success("🎉 Xử lý báo cáo NHNN thành công!")
                st.dataframe(pd.read_excel(path_output_nhnn), use_container_width=True)

elif report_type == "Báo Cáo Hoạt Động ATM (BC_atm)":
    st.subheader("🏧 MODULE BÁO CÁO HOẠT ĐỘNG ATM & CRM")
    PATH_NHATKY = os.path.join(SHARED_DATA_DIR, "BC_nhatkygiaodich.xlsx")
    PATH_PARAM_ATM = os.path.join(BASE_DIR, "parameters", "Thamso_atm.xlsx")
    path_output_atm = os.path.join(OUTPUT_DIR, f"BC_atm_{start_curr.strftime('%m%Y')}.xlsx")
    
    if st.button("🚀 Chạy và Xuất Báo Cáo ATM", type="primary"):
        if BC_ATM is None:
            st.error("❌ Không tìm thấy module `BC_ATM.py`.")
        else:
            df_res = BC_ATM.process_atm_summary_report(PATH_NHATKY, PATH_PARAM_ATM, start_curr, end_curr, start_prev, end_prev, path_output_atm)
            if df_res is not None: st.dataframe(df_res, use_container_width=True)

elif report_type == "Giao Dịch Phát Sinh Lớn (Ds_TiepQuyThuGom)":
    st.subheader("🔥 MODULE LỌC GIAO DỊCH PHÁT SINH LỚN")
    top_n = st.slider("Top N giao dịch:", 1, 50, 10)
    if st.button("🚀 Chạy Lọc", type="primary"):
        if giao_dich_lon:
            (vnd_c, usd_c), (vnd_p, usd_p) = giao_dich_lon.process_giao_dich_lon_report(SHARED_DATA_DIR, start_curr, end_curr, start_prev, end_prev, top_n)
            st.dataframe(vnd_c, use_container_width=True)

elif report_type == "Báo Cáo XNK Theo Loại Tiền & Đối Tác":
    st.subheader("💱 MODULE BÁO CÁO XUẤT NHẬP KHẨU")
    if st.button("🚀 Chạy XNK", type="primary"):
        if bc_xnk:
            curr, prev, diff, out = bc_xnk.xu_ly_bao_cao_xnk_pipeline(SHARED_DATA_DIR, OUTPUT_DIR, start_curr, end_curr, start_prev, end_prev)
            st.dataframe(curr, use_container_width=True)

elif report_type == "Nhập Đăng Ký Nộp Lĩnh":
    if nhapdknoplinh:
        nhapdknoplinh.show_dang_ky_nop_linh_page(BASE_DIR, SHARED_DATA_DIR, st.session_state.get('current_user', 'System'))

elif report_type == "Quản Lý Lỗi / Vi Phạm":
    if quanlyloi:
        quanlyloi.show_quan_ly_loi_page(BASE_DIR, SHARED_DATA_DIR, current_user=st.session_state.get('current_user', 'System'))

elif report_type == "Báo Cáo SLA Kho Quỹ":
    if show_sla_report_page:
        show_sla_report_page()

elif report_type == "Báo Cáo KPI Cán Bộ":
    if kpi_module and hasattr(kpi_module, "show_kpi_report_page"):
        kpi_module.show_kpi_report_page()
        if "df_master_result" in st.session_state and st.session_state.df_master_result is not None:
            render_sla_violation_dashboard(st.session_state.df_master_result, SHARED_DATA_DIR, os.path.join(BASE_DIR, "parameters"), start_curr, end_curr)

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>Hệ thống báo cáo nội bộ tự động hóa • Tích hợp hiển thị Web & Tải file</p>", unsafe_allow_html=True)
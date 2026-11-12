import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import calendar
import os
import hashlib

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

try:
    import bao_cao_word_module as bc_word
except ImportError:
    bc_word = None

# ==================== ĐỊNH NGHĨA THƯ MỤC DÙNG CHUNG ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SHARED_DATA_DIR = os.path.join(BASE_DIR, "data_source")
OUTPUT_DIR = os.path.join(BASE_DIR, "templates")

os.makedirs(SHARED_DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==================== HỆ THỐNG XÁC THỰC & BẢO MẬT ĐĂNG NHẬP (CHUẨN MaNV) ====================
def hash_password(password):
    return hashlib.sha256(str(password).encode('utf-8')).hexdigest()

def check_login_system(base_dir):
    param_dir = os.path.join(base_dir, "parameters")
    path_canbo = os.path.join(param_dir, "Thamso_canbo.xlsx")
    os.makedirs(param_dir, exist_ok=True)

    if not os.path.exists(path_canbo):
        try:
            df_init_cb = pd.DataFrame([
                {'MaNV': '12345', 'MatKhau': hash_password('00'), 'TrangThai': True, 'TenNV': '12345 - Nguyễn Văn A', 'NhomQuyen': 'KSV', 'CoSo': 'Cơ sở 1'},
                {'MaNV': '67890', 'MatKhau': hash_password('00'), 'TrangThai': True, 'TenNV': '67890 - Trần Thị B', 'NhomQuyen': 'NV', 'CoSo': 'Cơ sở 2'}
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
                        is_first_login = raw_trang_thai.strip().lower() in ['true', '1', 'yes'] if isinstance(raw_trang_thai, str) else bool(raw_trang_thai)
                        
                        input_pass_clean = password.strip()
                        input_check = input_pass_clean if (is_first_login and (db_pass == input_pass_clean or db_pass == hash_password(input_pass_clean))) else hash_password(input_pass_clean)
                        
                        if db_pass == input_check or (is_first_login and db_pass == input_pass_clean):
                            col_ten = next((c for c in df_cb.columns if 'ten' in str(c).lower() or 'tên' in str(c).lower()), None)
                            ten_nv = str(user_row.iloc[0][col_ten]).strip() if (col_ten and pd.notnull(user_row.iloc[0][col_ten])) else username.strip()
                            
                            st.session_state.current_manv = username.strip()
                            st.session_state.current_user = ten_nv
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

    if st.session_state.force_change_pass:
        st.set_page_config(page_title="Đổi Mật Khẩu Lần Đầu", page_icon="⚠️", layout="centered")
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
                        target_ma = str(st.session_state.get('current_manv', '')).strip()
                        idx = df_cb[df_cb['MaNV'].str.strip() == target_ma].index
                        if not idx.empty:
                            df_cb.loc[idx, 'MatKhau'] = str(hash_password(new_pass.strip()))
                            df_cb.loc[idx, 'TrangThai'] = False
                            df_cb.to_excel(path_canbo, index=False)
                            st.session_state.force_change_pass = False
                            st.success("🎉 Đổi mật khẩu thành công!")
                            st.rerun()
                    except Exception as e:
                        st.error(f"❌ Lỗi cập nhật mật khẩu: {e}")
        return False

    return True

if not check_login_system(BASE_DIR):
    st.stop()

st.set_page_config(page_title="Hệ Thống Báo Cáo Trung Tâm Tiếp Quỹ & NHNN", page_icon="📊", layout="wide")

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
prev_month_year = today.year - 1 if today.month == 1 else today.year
prev_month = 12 if today.month == 1 else today.month - 1
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

# ==================== XỬ LÝ MODULE BÁO CÁO TỔNG HỢP ====================
if report_type == "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)":
    st.subheader("📋 MODULE BÁO CÁO TỔNG HỢP TIẾP QUỸ & THU GOM")
    
    st.markdown(f"""
    <div style="background-color: #F5EFF2; padding: 12px 18px; border-radius: 10px; border: 1px solid #EEDCE3; margin-bottom: 20px;">
        <h4 style="margin: 0 0 8px 0; color: #4A3B45;">📅 KỲ BÁO CÁO ĐỐI CHÍNH</h4>
        • <b>Kỳ trước:</b> Từ ngày {start_prev.strftime('%d/%m/%Y')} đến {end_prev.strftime('%d/%m/%Y')}<br>
        • <b>Kỳ báo cáo:</b> Từ ngày {start_curr.strftime('%d/%m/%Y')} đến {end_curr.strftime('%d/%m/%Y')}
    </div>
    """, unsafe_allow_html=True)

    PATH_FILE1 = os.path.join(SHARED_DATA_DIR, "Ds_TiepQuyThuGom.xlsx")
    PATH_FILE2 = os.path.join(SHARED_DATA_DIR, "Ds_Noplinhkhac.xlsx")
    PATH_PARAM_KM = os.path.join(BASE_DIR, "parameters", "Thamso_KM.xlsx")
    PATH_TEMPLATE = os.path.join(BASE_DIR, "templates", "Template_BaoCao1.xlsx")

    st.info("ℹ️ **Lưu ý:** Quá trình xử lý dữ liệu lớn có thể mất từ vài giây đến vài phút. Hệ thống sẽ tự động thông báo tiến trình.")

    if st.button("⚡ Chạy và Xem Báo Cáo Tổng Hợp", type="primary"):
        if bc is None:
            st.error("❌ Không tìm thấy module `BC_th.py`.")
        else:
            with st.status("⏳ Đang xử lý dữ liệu hệ thống...", expanded=True) as status:
                try:
                    st.write("📂 Đang đọc và làm sạch dữ liệu nguồn...")
                    df1 = bc.clean_and_prepare_file1(PATH_FILE1) if os.path.exists(PATH_FILE1) else pd.DataFrame()
                    df2 = bc.clean_and_prepare_file2(PATH_FILE2) if os.path.exists(PATH_FILE2) else pd.DataFrame()
                    df_all = pd.concat([df1, df2], ignore_index=True)
                    df_param_km = pd.read_excel(PATH_PARAM_KM) if os.path.exists(PATH_PARAM_KM) else pd.DataFrame()
                    
                    st.write("⚙️ Đang tiến hành tổng hợp số liệu...")
                    all_coso = ["Cơ sở 1", "Cơ sở 2", "Hub NVC", "CS HĐ", "CS HP", "Toàn trung tâm"]
                    all_nghiep_vu = [nv for nv in df_all['LoaiNghiepVu'].dropna().unique() if nv != 'Bỏ qua'] + ['Tổng Thu', 'Tổng Chi']
                    
                    df_prev_res = bc.aggregate_period(df_all, start_prev, end_prev, df_param_km)
                    df_curr_res = bc.aggregate_period(df_all, start_curr, end_curr, df_param_km)
                    final_table = bc.format_summary_dataframe(df_prev_res, df_curr_res, all_nghiep_vu, all_coso)
                    
                    st.write("💾 Đang định dạng và xuất file Excel thông minh...")
                    thang_nam_str = start_curr.strftime("%m%Y")
                    ten_file_excel = f"BC_{thang_nam_str}.xlsx"
                    path_output_excel = os.path.join(OUTPUT_DIR, ten_file_excel)
                    
                    if hasattr(bc, "write_to_template"):
                        try:
                            bc.write_to_template(final_table, PATH_TEMPLATE, path_output_excel, start_prev, end_prev, start_curr, end_curr)
                        except TypeError:
                            bc.write_to_template(final_table, PATH_TEMPLATE, path_output_excel)
                    
                    st.write("📝 Đang lập báo cáo phân tích và đánh giá dạng Word (.docx)...")
                    path_output_word = None
                    ten_file_word = f"Danh_Gia_Hoat_Dong_{thang_nam_str}.docx"
                    if bc_word:
                        path_output_word = os.path.join(OUTPUT_DIR, ten_file_word)
                        bc_word.tao_bao_cao_word_danh_gia(final_table, start_prev, end_prev, start_curr, end_curr, path_output_word)

                    status.update(label="🎉 Hoàn thành xử lý toàn bộ báo cáo!", state="complete", expanded=False)
                except Exception as e:
                    status.update(label="❌ Xảy ra lỗi trong quá trình xử lý!", state="error", expanded=True)
                    st.error(f"Chi tiết lỗi: {e}")
                    st.stop()

            st.markdown("---")
            st.subheader("📋 Bảng Tổng Hợp Kết Quả Tiếp Quỹ & Thu Gom")
            st.dataframe(final_table, use_container_width=True)
            
            st.markdown("---")
            st.subheader("📥 KẾT QUẢ XUẤT FILE BÁO CÁO")
            
            # Chia làm 2 cột song song để đặt 2 nút tải độc lập, không bị thoát app
            c_dl1, c_dl2 = st.columns(2)
            
            with c_dl1:
                if os.path.exists(path_output_excel):
                    with open(path_output_excel, "rb") as f_excel:
                        st.download_button(
                            label=f"📥 TẢI BÁO CÁO EXCEL ({ten_file_excel})",
                            data=f_excel,
                            file_name=ten_file_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
            with c_dl2:
                if path_output_word and os.path.exists(path_output_word):
                    with open(path_output_word, "rb") as f_word:
                        st.download_button(
                            label=f"📥 TẢI BÁO CÁO WORD ({ten_file_word})",
                            data=f_word,
                            file_name=ten_file_word,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

elif report_type == "Báo Cáo Nộp Lĩnh NHNN (Ds_NopLinhNHNN)":
    st.subheader("🏦 MODULE: DANH SÁCH NỘP LĨNH NHNN")
    thang_nam_nhnn = start_curr.strftime("%m%Y")
    path_output_nhnn = os.path.join(OUTPUT_DIR, f"Bao_Cao_DS_NopLinh_NHNN_{thang_nam_nhnn}.xlsx")
    if st.button("🚀 Chạy và Xem Báo Cáo NHNN", type="primary"):
        if BC_NHNN:
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
        if BC_ATM:
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

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>Hệ thống báo cáo nội bộ tự động hóa • Tích hợp hiển thị Web & Tải file</p>", unsafe_allow_html=True)
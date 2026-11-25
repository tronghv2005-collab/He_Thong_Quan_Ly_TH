import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import calendar
import os
import hashlib
import shutil
import uuid

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

try:
    import nhap_tien_moi
except ImportError:
    nhap_tien_moi = None


# ==================== CẤU HÌNH UPLOAD CHO FILE DUNG LƯỢNG LỚN ====================
def _configure_large_upload_limit(max_mb=512):
    try:
        current_limit = st.get_option("server.maxUploadSize")
        if current_limit is None or current_limit < max_mb:
            st.set_option("server.maxUploadSize", max_mb)
    except Exception:
        pass


_configure_large_upload_limit(max_mb=512)


# ==================== ĐỊNH NGHĨA THƯ MỤC DÙNG CHUNG ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SHARED_DATA_DIR = os.path.join(BASE_DIR, "data_source")
UPLOAD_ROOT_DIR = os.path.join(SHARED_DATA_DIR, "_uploaded_session")
OUTPUT_DIR = os.path.join(BASE_DIR, "templates")

os.makedirs(SHARED_DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_ROOT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Ánh xạ chính xác các file thuộc data_source theo từng chức năng/báo cáo
REPORT_REQUIRED_FILES = {
    "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)": [
        "Ds_TiepQuyThuGom.xlsx",
        "Ds_Noplinhkhac.xlsx",
    ],
    "Báo Cáo Nộp Lĩnh NHNN (Ds_NopLinhNHNN)": [
        "Ds_NopLinhNHNN.xlsx",
    ],
    "Báo Cáo Hoạt Động ATM (BC_atm)": [
        "BC_nhatkygiaodich.xlsx",
    ],
    "Báo Cáo XNK Theo Loại Tiền & Đối Tác": [
        "xnk_2026.xlsx",
    ],
    "Giao Dịch Phát Sinh Lớn (Ds_TiepQuyThuGom)": [
        "Ds_TiepQuyThuGom.xlsx",
    ],
    "Báo Cáo SLA Kho Quỹ": [
        "Ds_TiepQuyThuGom.xlsx",
        "QLNT.xlsx",
    ],
    "Báo Cáo KPI Cán Bộ": [
        "Ds_TiepQuyThuGom.xlsx",
        "Ds_Noplinhkhac.xlsx",
        "Ds_NopLinhNHNN.xlsx",
        "atm.xlsx",
        "kiemdem.xlsx",
    ],
}

ALL_DATA_SOURCE_FILES = sorted(
    {fname for flist in REPORT_REQUIRED_FILES.values() for fname in flist}
)


def get_session_upload_dir():
    if "upload_session_id" not in st.session_state:
        st.session_state.upload_session_id = uuid.uuid4().hex[:12]
    s_dir = os.path.join(UPLOAD_ROOT_DIR, st.session_state.upload_session_id)
    os.makedirs(s_dir, exist_ok=True)
    return s_dir


def save_uploaded_stream(uploaded_file, dest_path):
    """Ghi file tải lên theo từng khối (1MB) để tối ưu bộ nhớ với file > 150 MB."""
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    uploaded_file.seek(0)
    with open(dest_path, "wb") as f_out:
        shutil.copyfileobj(uploaded_file, f_out, length=1024 * 1024)


def get_data_file_path_with_source(filename, active_data_dir):
    """
    Trả về: (full_path, source_label)
    - Nếu chọn 'Kéo lên từ máy' và file tồn tại trong active_data_dir -> dùng file kéo lên ('Kéo từ máy').
    - Nếu file chưa kéo lên (hoặc chọn 'Thư mục HUB') -> tự động lấy từ thư mục HUB ('Thư mục HUB').
    """
    session_dir = get_session_upload_dir()
    uploaded_path = os.path.join(session_dir, filename)
    hub_path = os.path.join(SHARED_DATA_DIR, filename)

    if active_data_dir != SHARED_DATA_DIR:
        candidate = os.path.join(active_data_dir, filename)
        if os.path.exists(candidate):
            return candidate, "Kéo từ máy"
        if os.path.exists(uploaded_path):
            return uploaded_path, "Kéo từ máy"

    return hub_path, "Thư mục HUB"


def get_data_file_path(filename, active_data_dir):
    """Lấy đường dẫn file thực tế (ưu tiên file kéo lên nếu có, chưa kéo lên thì lấy ở HUB)."""
    fpath, _ = get_data_file_path_with_source(filename, active_data_dir)
    return fpath


def build_effective_data_dir(required_files, active_data_dir):
    """
    Tạo thư mục phiên hiệu lực chứa đúng các file cần thiết cho báo cáo:
    - File đã kéo lên trong phiên hiện tại: copy sang thư mục phiên hiệu lực.
    - File chưa kéo lên: tự động copy từ thư mục HUB (SHARED_DATA_DIR) để các module
      đọc cùng một thư mục mà không thiếu file nào.
    """
    if active_data_dir == SHARED_DATA_DIR or not required_files:
        return SHARED_DATA_DIR

    session_dir = get_session_upload_dir()
    effective_dir = os.path.join(session_dir, "_active_run")
    os.makedirs(effective_dir, exist_ok=True)

    for old_f in os.listdir(effective_dir):
        old_p = os.path.join(effective_dir, old_f)
        if os.path.isfile(old_p):
            try:
                os.remove(old_p)
            except OSError:
                pass

    for fname in required_files:
        up_file_p = os.path.join(session_dir, fname)
        hub_file_p = os.path.join(SHARED_DATA_DIR, fname)
        src_p = up_file_p if os.path.exists(up_file_p) else hub_file_p
        if os.path.exists(src_p):
            dst_p = os.path.join(effective_dir, fname)
            with open(src_p, "rb") as f_src, open(dst_p, "wb") as f_dst:
                shutil.copyfileobj(f_src, f_dst, length=1024 * 1024)

    return effective_dir


def render_data_source_selector(required_files, report_label=""):
    """
    Chỉ hiển thị đúng các file cần cho chức năng hiện tại.
    Hỗ trợ:
    1. Lấy toàn bộ từ thư mục HUB.
    2. Kéo lên từ máy (file nào chưa kéo lên tự động lấy từ thư mục HUB).
    """
    if "data_source_mode" not in st.session_state:
        st.session_state.data_source_mode = "Thư mục HUB"

    if not required_files:
        return "Thư mục HUB", SHARED_DATA_DIR

    session_dir = get_session_upload_dir()

    st.sidebar.markdown("---")
    st.sidebar.subheader(f"📂 Nguồn dữ liệu: {report_label}")

    mode = st.sidebar.radio(
        "Chọn cách lấy dữ liệu:",
        options=["Thư mục HUB", "Kéo lên từ máy"],
        index=0 if st.session_state.data_source_mode == "Thư mục HUB" else 1,
        key=f"radio_ds_mode_{report_label}",
    )
    st.session_state.data_source_mode = mode

    if mode == "Thư mục HUB":
        st.sidebar.caption("Đang lấy trực tiếp từ thư mục HUB (`data_source`):")
        lines = [f"🏢 `{fname}`" for fname in required_files]
        st.sidebar.markdown("<br>".join(lines), unsafe_allow_html=True)
        return "Thư mục HUB", SHARED_DATA_DIR

    st.sidebar.caption("Kéo thả nhiều file Excel cần cập nhật (file chưa kéo lên sẽ tự lấy từ HUB):")
    uploaded_files = st.sidebar.file_uploader(
        "Kéo thả file Excel vào đây:",
        type=["xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key=f"uploader_ds_{report_label}",
    )

    if uploaded_files:
        saved_names = []
        skipped_names = []
        allowed_lower = {f.lower(): f for f in required_files}

        for up_file in uploaded_files:
            up_lower = up_file.name.strip().lower()
            if up_lower in allowed_lower:
                std_name = allowed_lower[up_lower]
                save_path = os.path.join(session_dir, std_name)
                save_uploaded_stream(up_file, save_path)
                saved_names.append(std_name)
            else:
                skipped_names.append(up_file.name)

        if saved_names:
            st.sidebar.success(f"Đã nạp: {', '.join(saved_names)}")
        if skipped_names:
            st.sidebar.warning(f"Bỏ qua file không thuộc báo cáo này: {', '.join(skipped_names)}")

    status_lines = []
    for fname in required_files:
        up_p = os.path.join(session_dir, fname)
        if os.path.exists(up_p):
            sz_mb = os.path.getsize(up_p) / (1024 * 1024)
            status_lines.append(f"📤 `{fname}` → **Kéo từ máy** ({sz_mb:.1f} MB)")
        else:
            status_lines.append(f"🏢 `{fname}` → **Thư mục HUB**")
    st.sidebar.markdown("<br>".join(status_lines), unsafe_allow_html=True)

    if st.sidebar.button("↩️ Xóa file kéo lên (Dùng HUB)", key=f"btn_clear_up_{report_label}", use_container_width=True):
        for fname in required_files:
            up_p = os.path.join(session_dir, fname)
            if os.path.exists(up_p):
                try:
                    os.remove(up_p)
                except OSError:
                    pass
        st.rerun()

    effective_dir = build_effective_data_dir(required_files, session_dir)
    return "Kéo lên từ máy (thiếu tự lấy HUB)", effective_dir


def render_source_file_diagnostics(file_paths_dict):
    """Hiển thị thông tin đường dẫn, nguồn, dung lượng và thời gian cập nhật thực tế của các file nguồn."""
    for label, fpath in file_paths_dict.items():
        if os.path.exists(fpath):
            sz_mb = os.path.getsize(fpath) / (1024 * 1024)
            mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).strftime("%Y-%m-%d %H:%M:%S")
            st.write(f"📁 **{label}**: `{fpath}` | Dung lượng: **{sz_mb:.2f} MB** | Cập nhật: **{mtime}**")
        else:
            st.warning(f"⚠️ Không tìm thấy **{label}** tại: `{fpath}`")


def format_thousands_table(df_in, round_unit=True):
    """
    Làm tròn số đến hàng đơn vị (mặc định) và định dạng dấu phẩy (,) ngăn cách hàng nghìn
    cho các cột chứa dữ liệu số trên bảng hiển thị web.
    """
    if df_in is None or df_in.empty:
        return pd.DataFrame()
    df_fmt = df_in.copy()
    for col in df_fmt.columns:
        num_vals = pd.to_numeric(df_fmt[col], errors="coerce")
        if num_vals.notna().any():
            if round_unit:
                df_fmt[col] = [
                    "{:,.0f}".format(round(float(n))) if pd.notnull(n) else orig
                    for orig, n in zip(df_fmt[col], num_vals)
                ]
            else:
                df_fmt[col] = [
                    "{:,.2f}".format(float(n)) if pd.notnull(n) else orig
                    for orig, n in zip(df_fmt[col], num_vals)
                ]
    return df_fmt


# ==================== HỆ THỐNG XÁC THỰC & BẢO MẬT ĐĂNG NHẬP ====================
def hash_password(password):
    return hashlib.sha256(str(password).encode('utf-8')).hexdigest()


def check_login_system(base_dir):
    param_dir = os.path.join(base_dir, "parameters")
    path_canbo = os.path.join(param_dir, "Thamso_canbo.xlsx")
    os.makedirs(param_dir, exist_ok=True)

    if not os.path.exists(path_canbo):
        try:
            df_init_cb = pd.DataFrame([
                {'MaNV': 'admin', 'MatKhau': hash_password('admin123'), 'TrangThai': False, 'TenNV': 'Quản trị viên Hệ thống', 'NhomQuyen': 'ADMIN', 'CoSo': 'Toàn trung tâm'},
                {'MaNV': '97474', 'MatKhau': hash_password('00'), 'TrangThai': True, 'TenNV': '97474 - Nguyễn Văn A', 'NhomQuyen': 'KSV', 'CoSo': 'Cơ sở 1'},
                {'MaNV': '67890', 'MatKhau': hash_password('00'), 'TrangThai': True, 'TenNV': '67890 - Trần Thị B', 'NhomQuyen': 'NV', 'CoSo': 'Cơ sở 2'}
            ])
            df_init_cb['MaNV'] = df_init_cb['MaNV'].astype(str)
            df_init_cb['MatKhau'] = df_init_cb['MatKhau'].astype(str)
            df_init_cb['TrangThai'] = df_init_cb['TrangThai'].astype(bool)
            df_init_cb.to_excel(path_canbo, index=False)
        except Exception:
            pass

    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'current_user' not in st.session_state:
        st.session_state.current_user = None
    if 'current_manv' not in st.session_state:
        st.session_state.current_manv = None
    if 'current_role' not in st.session_state:
        st.session_state.current_role = None
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
            username = st.text_input("Mã cán bộ (MaNV) hoặc tài khoản Admin")
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
                            is_first_login = raw_trang_thai.strip().lower() in ['true', '1', 'yes', 't']
                        else:
                            is_first_login = bool(raw_trang_thai) if pd.notnull(raw_trang_thai) else False

                        input_pass_clean = password.strip()
                        input_hashed = hash_password(input_pass_clean)

                        if db_pass == input_hashed or db_pass == input_pass_clean:
                            col_ten = next((c for c in df_cb.columns if 'ten' in str(c).lower() or 'tên' in str(c).lower()), None)
                            col_quyen = next((c for c in df_cb.columns if 'quyen' in str(c).lower() or 'nhom' in str(c).lower()), None)

                            ten_nv = str(user_row.iloc[0][col_ten]).strip() if (col_ten and pd.notnull(user_row.iloc[0][col_ten])) else username.strip()
                            nhom_quyen = str(user_row.iloc[0][col_quyen]).strip().upper() if (col_quyen and pd.notnull(user_row.iloc[0][col_quyen])) else "NV"

                            if username.strip().lower() == "admin":
                                nhom_quyen = "ADMIN"

                            st.session_state.current_manv = username.strip()
                            st.session_state.current_user = ten_nv
                            st.session_state.current_role = nhom_quyen
                            st.session_state.logged_in = True

                            if nhom_quyen == "ADMIN":
                                st.session_state.force_change_pass = False
                            else:
                                st.session_state.force_change_pass = bool(is_first_login)

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
        st.warning("⚠️ **Đăng nhập lần đầu hoặc sau khi được Reset: Vui lòng nhập mật khẩu mới để tiếp tục!**")

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

                        if 'TrangThai' in df_cb.columns:
                            df_cb['TrangThai'] = df_cb['TrangThai'].astype(bool)
                        else:
                            df_cb['TrangThai'] = False

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

st.title("📊 HỆ THỐNG QUẢN LÝ TT KHO QUỸ")
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
st.sidebar.markdown(f"🔑 **Phân quyền:** `{st.session_state.current_role}`")
if st.sidebar.button("🚪 Đăng xuất", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.current_manv = None
    st.session_state.current_role = None
    st.session_state.force_change_pass = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("⚙️ TÙY CHỌN CHỨC NĂNG")

list_options = [
    "💵 Giao Dịch Tiền Mới",
    "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)",
    "Báo Cáo Nộp Lĩnh NHNN (Ds_NopLinhNHNN)",
    "Báo Cáo Hoạt Động ATM (BC_atm)",
    "Giao Dịch Phát Sinh Lớn (Ds_TiepQuyThuGom)",
    "Báo Cáo XNK Theo Loại Tiền & Đối Tác",
    "Nhập Đăng Ký Nộp Lĩnh",
    "Quản Lý Lỗi / Vi Phạm",
    "Báo Cáo SLA Kho Quỹ",
    "Báo Cáo KPI Cán Bộ"
]

if str(st.session_state.current_role).strip().upper() == "ADMIN":
    list_options.append("Quản Trị Hệ Thống (Reset Mật Khẩu NV)")

report_type = st.sidebar.selectbox("Chọn loại báo cáo / chức năng cần thực hiện:", tuple(list_options))

if report_type != "💵 Giao Dịch Tiền Mới":
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Thiết lập khoảng thời gian")
    start_prev = st.sidebar.date_input("Từ ngày (Kỳ trước)", value=default_kp_start)
    end_prev = st.sidebar.date_input("Đến ngày (Kỳ trước)", value=default_kp_end)
    start_curr = st.sidebar.date_input("Từ ngày (Kỳ này)", value=default_kc_start)
    end_curr = st.sidebar.date_input("Đến ngày (Kỳ này)", value=default_kc_end)

req_files = REPORT_REQUIRED_FILES.get(report_type, [])
data_mode, active_data_dir = render_data_source_selector(req_files, report_type)

# ==================== XỬ LÝ CÁC MODULE ====================
if report_type == "💵 Giao Dịch Tiền Mới":
    if nhap_tien_moi:
        nhap_tien_moi.show_nhap_tien_moi_page(
            base_dir=BASE_DIR,
            shared_data_dir=SHARED_DATA_DIR,
            current_user=st.session_state.get('current_user', 'System'),
            user_role=st.session_state.get('current_role', 'NV')
        )
    else:
        st.error("❌ Không tìm thấy module `nhap_tien_moi.py` trong thư mục hệ thống.")

elif report_type == "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)":
    st.subheader("📋 MODULE BÁO CÁO TỔNG HỢP HOẠT ĐỘNG CÁC PHÒNG NGÂN QUỸ")

    st.markdown(f"""
    <div style="background-color: #F5EFF2; padding: 12px 18px; border-radius: 10px; border: 1px solid #EEDCE3; margin-bottom: 20px;">
        <h4 style="margin: 0 0 8px 0; color: #4A3B45;">📅 KỲ BÁO CÁO ĐỐI CHIẾU</h4>
        • <b>Kỳ trước:</b> Từ ngày {start_prev.strftime('%d/%m/%Y')} đến {end_prev.strftime('%d/%m/%Y')}<br>
        • <b>Kỳ báo cáo:</b> Từ ngày {start_curr.strftime('%d/%m/%Y')} đến {end_curr.strftime('%d/%m/%Y')}<br>
        • <b>Nguồn dữ liệu:</b> {data_mode}
    </div>
    """, unsafe_allow_html=True)

    PATH_FILE1 = get_data_file_path("Ds_TiepQuyThuGom.xlsx", active_data_dir)
    PATH_FILE2 = get_data_file_path("Ds_Noplinhkhac.xlsx", active_data_dir)
    PATH_PARAM_KM = os.path.join(BASE_DIR, "parameters", "Thamso_KM.xlsx")
    PATH_TEMPLATE = os.path.join(BASE_DIR, "templates", "Template_BaoCao1.xlsx")

    st.info("ℹ️ **Lưu ý:** Quá trình xử lý file dữ liệu lớn (như `Ds_TiepQuyThuGom.xlsx` > 150 MB) có thể mất từ vài giây đến vài phút. Hệ thống sẽ tự động thông báo tiến trình.")

    if st.button("⚡ Chạy và Xem Báo Cáo Tổng Hợp", type="primary"):
        if bc is None:
            st.error("❌ Không tìm thấy module `BC_th.py`.")
        else:
            with st.status("⏳ Đang xử lý dữ liệu hệ thống...", expanded=True) as status:
                try:
                    st.write("📂 Đang kiểm tra nguồn và đọc dữ liệu...")
                    render_source_file_diagnostics({
                        "File 1 (Ds_TiepQuyThuGom.xlsx)": PATH_FILE1,
                        "File 2 (Ds_Noplinhkhac.xlsx)": PATH_FILE2,
                    })

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
            st.subheader("📋 BẢNG TỔNG HỢP HOẠT ĐỘNG CÁC PHÒNG NGÂN QUỸ")
            st.dataframe(format_thousands_table(final_table, round_unit=True), use_container_width=True)

            st.markdown("---")
            st.subheader("📥 KẾT QUẢ XUẤT FILE BÁO CÁO")

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
    ten_file_nhnn = f"Bao_Cao_DS_NopLinh_NHNN_{start_curr.strftime('%m%Y')}.xlsx"
    path_output_nhnn = os.path.join(OUTPUT_DIR, ten_file_nhnn)
    path_in_nhnn = get_data_file_path("Ds_NopLinhNHNN.xlsx", active_data_dir)

    if st.button("🚀 Chạy và Xem Báo Cáo NHNN", type="primary"):
        if BC_NHNN:
            render_source_file_diagnostics({"Ds_NopLinhNHNN.xlsx": path_in_nhnn})
            success = BC_NHNN.process_ds_nop_linh_nhnn_pipeline(
                active_data_dir, start_curr, end_curr, start_prev, end_prev, path_output_nhnn
            )
            if success and os.path.exists(path_output_nhnn):
                st.success("🎉 Xử lý báo cáo NHNN thành công!")
                df_nhnn = pd.read_excel(path_output_nhnn)
                st.dataframe(format_thousands_table(df_nhnn, round_unit=True), use_container_width=True)

                with open(path_output_nhnn, "rb") as f_nhnn:
                    st.download_button(
                        label=f"📥 TẢI BÁO CÁO EXCEL ({ten_file_nhnn})",
                        data=f_nhnn,
                        file_name=ten_file_nhnn,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
        else:
            st.error("❌ Không tìm thấy module `BC_NHNN.py`.")

elif report_type == "Báo Cáo Hoạt Động ATM (BC_atm)":
    st.subheader("🏧 MODULE BÁO CÁO HOẠT ĐỘNG ATM & CRM")
    PATH_NHATKY = get_data_file_path("BC_nhatkygiaodich.xlsx", active_data_dir)
    PATH_PARAM_ATM = os.path.join(BASE_DIR, "parameters", "Thamso_atm.xlsx")
    ten_file_atm = f"BC_atm_{start_curr.strftime('%m%Y')}.xlsx"
    path_output_atm = os.path.join(OUTPUT_DIR, ten_file_atm)

    if st.button("🚀 Chạy và Xuất Báo Cáo ATM", type="primary"):
        if BC_ATM:
            render_source_file_diagnostics({"BC_nhatkygiaodich.xlsx": PATH_NHATKY})
            df_res = BC_ATM.process_atm_summary_report(
                PATH_NHATKY, PATH_PARAM_ATM, start_curr, end_curr, start_prev, end_prev, path_output_atm
            )
            if df_res is not None and not df_res.empty:
                st.dataframe(format_thousands_table(df_res, round_unit=True), use_container_width=True)

                if os.path.exists(path_output_atm):
                    with open(path_output_atm, "rb") as f_atm:
                        st.download_button(
                            label=f"📥 TẢI BÁO CÁO EXCEL ({ten_file_atm})",
                            data=f_atm,
                            file_name=ten_file_atm,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
        else:
            st.error("❌ Không tìm thấy module `BC_ATM.py`.")

elif report_type == "Giao Dịch Phát Sinh Lớn (Ds_TiepQuyThuGom)":
    st.subheader("🔥 MODULE LỌC GIAO DỊCH PHÁT SINH LỚN")
    top_n = st.slider("Top N giao dịch:", 1, 50, 10)
    path_in_gdl = get_data_file_path("Ds_TiepQuyThuGom.xlsx", active_data_dir)

    if st.button("🚀 Chạy Lọc", type="primary"):
        if giao_dich_lon:
            render_source_file_diagnostics({"Ds_TiepQuyThuGom.xlsx": path_in_gdl})
            (vnd_c, usd_c), (vnd_p, usd_p) = giao_dich_lon.process_giao_dich_lon_report(
                active_data_dir, start_curr, end_curr, start_prev, end_prev, top_n
            )
            st.markdown("#### 💵 Kỳ này - VND")
            st.dataframe(format_thousands_table(vnd_c, round_unit=True), use_container_width=True)

            if usd_c is not None and not usd_c.empty:
                st.markdown("#### 💲 Kỳ này - Ngoại tệ / USD")
                st.dataframe(format_thousands_table(usd_c, round_unit=True), use_container_width=True)
        else:
            st.error("❌ Không tìm thấy module `giao_dich_lon.py`.")

elif report_type == "Báo Cáo XNK Theo Loại Tiền & Đối Tác":
    st.subheader("💱 MODULE BÁO CÁO XUẤT NHẬP KHẨU")
    path_in_xnk = get_data_file_path("xnk_2026.xlsx", active_data_dir)

    if st.button("🚀 Chạy XNK", type="primary"):
        if bc_xnk:
            render_source_file_diagnostics({"xnk_2026.xlsx": path_in_xnk})
            curr, prev, diff, out = bc_xnk.xu_ly_bao_cao_xnk_pipeline(
                active_data_dir, OUTPUT_DIR, start_curr, end_curr, start_prev, end_prev
            )
            st.dataframe(format_thousands_table(curr, round_unit=True), use_container_width=True)
        else:
            st.error("❌ Không tìm thấy module `bc_xnk.py`.")

elif report_type == "Nhập Đăng Ký Nộp Lĩnh":
    if nhapdknoplinh:
        nhapdknoplinh.show_dang_ky_nop_linh_page(BASE_DIR, SHARED_DATA_DIR, st.session_state.get('current_user', 'System'))

elif report_type == "Quản Lý Lỗi / Vi Phạm":
    if quanlyloi:
        quanlyloi.show_quan_ly_loi_page(BASE_DIR, SHARED_DATA_DIR, current_user=st.session_state.get('current_user', 'System'))

elif report_type == "Báo Cáo SLA Kho Quỹ":
    if show_sla_report_page:
        try:
            show_sla_report_page(active_data_dir)
        except TypeError:
            show_sla_report_page()

elif report_type == "Báo Cáo KPI Cán Bộ":
    if kpi_module and hasattr(kpi_module, "show_kpi_report_page"):
        try:
            kpi_module.show_kpi_report_page(active_data_dir)
        except TypeError:
            kpi_module.show_kpi_report_page()

elif report_type == "Quản Trị Hệ Thống (Reset Mật Khẩu NV)":
    st.subheader("🛠️ QUẢN TRỊ HỆ THỐNG - QUẢN LÝ TÀI KHOẢN & RESET MẬT KHẨU")

    path_canbo = os.path.join(BASE_DIR, "parameters", "Thamso_canbo.xlsx")
    if os.path.exists(path_canbo):
        try:
            df_cb = pd.read_excel(path_canbo, dtype={'MaNV': str})
            st.write("📋 **Danh sách cán bộ hiện tại trong hệ thống:**")
            st.dataframe(df_cb[['MaNV', 'TenNV', 'NhomQuyen', 'CoSo', 'TrangThai']], use_container_width=True)

            st.markdown("---")
            st.markdown("#### 🔄 Cấp lại / Reset Mật khẩu cho nhân viên")
            with st.form("admin_reset_form"):
                list_ma_nv = df_cb['MaNV'].tolist()
                selected_nv = st.selectbox("Chọn Mã Cán Bộ cần reset mật khẩu:", list_ma_nv)
                new_default_pass = st.text_input("Mật khẩu mới (hoặc mật khẩu mặc định):", value="00")

                submit_reset = st.form_submit_button("Xác nhận Reset Mật Khẩu", type="primary")
                if submit_reset:
                    if 'TrangThai' in df_cb.columns:
                        df_cb['TrangThai'] = df_cb['TrangThai'].astype(bool)
                    else:
                        df_cb['TrangThai'] = False

                    idx = df_cb[df_cb['MaNV'].str.strip() == str(selected_nv).strip()].index
                    if not idx.empty:
                        df_cb.loc[idx, 'MatKhau'] = str(hash_password(new_default_pass.strip()))
                        df_cb.loc[idx, 'TrangThai'] = True
                        df_cb.to_excel(path_canbo, index=False)
                        st.success(f"🎉 Đã reset mật khẩu thành công cho Mã cán bộ: **{selected_nv}** với mật khẩu mới: `{new_default_pass}`! (Nhân viên sẽ bắt buộc phải đổi mật khẩu khi đăng nhập).")
        except Exception as e:
            st.error(f"❌ Lỗi đọc file danh sách cán bộ: {e}")
    else:
        st.warning("⚠️ Không tìm thấy file `Thamso_canbo.xlsx` trong thư mục `parameters`.")

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>Hệ thống báo cáo nội bộ tự động hóa • Tích hợp hiển thị Web & Tải file</p>", unsafe_allow_html=True)
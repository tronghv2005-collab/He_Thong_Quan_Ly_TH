import streamlit as st
import pandas as pd
import numpy as np  # <--- ĐÃ BỔ SUNG KHAI BÁO NÀY ĐỂ HẾT LỖI
from datetime import datetime, date
import calendar
import os
import plotly.express as px
import plotly.graph_objects as go

# Import module Báo Cáo Tổng Hợp
try:
    import BC_th as bc
except ImportError:
    bc = None

# Import module Báo Cáo Nộp Lĩnh NHNN
try:
    import BC_NHNN
except ImportError:
    BC_NHNN = None

# Import module Báo Cáo Hoạt Động ATM & CRM
try:
    import BC_ATM
except ImportError:
    BC_ATM = None

# Import module Giao Dịch Phát Sinh Lớn
try:
    import giao_dich_lon
except ImportError:
    giao_dich_lon = None

# Import module Báo Cáo XNK mới tích hợp
try:
    import bc_xnk
except ImportError:
    bc_xnk = None

# Import module Báo Cáo SLA Kho Quỹ mới
try:
    from sla_report import show_sla_report_page
except ImportError:
    show_sla_report_page = None

# Import module Báo Cáo KPI Cán Bộ mới
try:
    import kpi_module
except ImportError:
    kpi_module = None

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Hệ Thống Báo Cáo Trung Tâm Tiếp Quỹ & NHNN",
    page_icon="📊",
    layout="wide"
)

# ==================== CSS TÙY CHỈNH PHONG CÁCH TINH TẾ, MỀM MẠI ====================
st.markdown("""
    <style>
    .main {
        background-color: #FAFAFC;
    }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #FFF9FB 0%, #F5EFF2 100%);
        border: 1px solid #EEDCE3;
        padding: 16px;
        border-radius: 14px;
        box-shadow: 0 4px 15px rgba(210, 180, 195, 0.12);
    }
    div[data-testid="stMetric"] label {
        color: #7D6B75 !important;
        font-weight: 500;
    }
    h3, h4 {
        color: #4A3B45;
        font-family: 'Helvetica Neue', sans-serif;
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

# ==================== TÍNH TOÁN NGÀY MẶC ĐỊNH (TỰ ĐỘNG THEO THÁNG) ====================
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
st.sidebar.header("⚙️ TÙY CHỌN CHỨC NĂNG")
report_type = st.sidebar.selectbox(
    "Chọn loại báo cáo cần thực hiện:", 
    (
        "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)", 
        "Báo Cáo Nộp Lĩnh NHNN (Ds_NopLinhNHNN)",
        "Báo Cáo Hoạt Động ATM (BC_atm)",
        "Giao Dịch Phát Sinh Lớn (Ds_Tiepquythugom)",
        "Báo Cáo XNK Theo Loại Tiền & Đối Tác",  # <--- ĐÃ TÍCH HỢP MỚI TẠI ĐÂY
        "Báo Cáo SLA Kho Quỹ",
        "Báo Cáo KPI Cán Bộ" 
    )
)

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Thiết lập khoảng thời gian")

st.sidebar.markdown("**1. Kỳ Trước (So sánh)**")
start_prev = st.sidebar.date_input("Từ ngày (Kỳ trước)", value=default_kp_start)
end_prev = st.sidebar.date_input("Đến ngày (Kỳ trước)", value=default_kp_end)

st.sidebar.markdown("**2. Kỳ Báo Cáo (Hiện tại)**")
start_curr = st.sidebar.date_input("Từ ngày (Kỳ này)", value=default_kc_start)
end_curr = st.sidebar.date_input("Đến ngày (Kỳ này)", value=default_kc_end)


# ==================== ĐỊNH NGHĨA THƯ MỤC DÙNG CHUNG ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SHARED_DATA_DIR = os.path.join(BASE_DIR, "data_source")
OUTPUT_DIR = os.path.join(BASE_DIR, "templates")

os.makedirs(SHARED_DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ==================== HÀM RENDER DASHBOARD SLA TRỰC QUAN ====================
def render_sla_violation_dashboard(df_master, shared_data_dir, param_dir, start_date, end_date):
    st.markdown("---")
    st.subheader("🌸 DASHBOARD TRỰC QUAN & GIÁM SÁT VI PHẠM SLA")

    @st.cache_data
    def load_mock_or_real_sla_data(df_m):
        if df_m.empty:
            return pd.DataFrame()
        
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
    
    if df_sla.empty:
        st.info("ℹ️ Không có dữ liệu SLA trong khoảng thời gian này.")
        return

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        selected_coso_sla = st.selectbox("🔍 Lọc theo Cơ sở:", ["Tất cả"] + list(df_sla['CoSo'].unique()), key="sla_filter_coso")
    with col_f2:
        selected_nv_sla = st.selectbox("🔍 Lọc theo Nghiệp vụ:", ["Tất cả"] + list(df_sla['LoaiNghiepVu'].unique()), key="sla_filter_nv")

    df_filtered = df_sla.copy()
    if selected_coso_sla != "Tất cả":
        df_filtered = df_filtered[df_filtered['CoSo'] == selected_coso_sla]
    if selected_nv_sla != "Tất cả":
        df_filtered = df_filtered[df_filtered['LoaiNghiepVu'] == selected_nv_sla]

    total_tx = len(df_filtered)
    total_violation = df_filtered['ViPham'].sum()
    rate_violation = (total_violation / total_tx * 100) if total_tx > 0 else 0.0

    mcol1, mcol2, mcol3 = st.columns(3)
    mcol1.metric("📌 Tổng số Giao dịch giám sát", f"{total_tx:,}")
    mcol2.metric("⚠️ Tổng số vụ Vi phạm SLA", f"{total_violation:,}", delta=f"-{rate_violation:.2f}%", delta_color="inverse")
    mcol3.metric("📈 Tỷ lệ vi phạm chung", f"{rate_violation:.2f}%")

    st.markdown("---")

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.markdown("#### 🏢 Tỷ lệ vi phạm theo Cơ sở")
        df_coso_agg = df_filtered.groupby('CoSo')['ViPham'].agg(['count', 'sum']).reset_index()
        df_coso_agg['TyLeViPham'] = (df_coso_agg['sum'] / df_coso_agg['count']) * 100
        
        fig_coso = px.bar(
            df_coso_agg, x='CoSo', y='TyLeViPham', text='TyLeViPham',
            labels={'TyLeViPham': 'Tỷ lệ vi phạm (%)', 'CoSo': 'Cơ sở'},
            color='TyLeViPham', color_continuous_scale='Sunset'
        )
        fig_coso.update_traces(texttemplate='%{text:.2f}%', textposition='outside', marker_line_width=0, opacity=0.85)
        fig_coso.update_layout(
            uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-20,
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_coso, use_container_width=True)

    with chart_col2:
        st.markdown("#### ⚙️ Tỷ lệ vi phạm theo Loại Nghiệp Vụ")
        df_nv_agg = df_filtered.groupby('LoaiNghiepVu')['ViPham'].agg(['count', 'sum']).reset_index()
        df_nv_agg['TyLeViPham'] = (df_nv_agg['sum'] / df_nv_agg['count']) * 100
        
        fig_nv = px.bar(
            df_nv_agg, x='LoaiNghiepVu', y='TyLeViPham', text='TyLeViPham',
            labels={'TyLeViPham': 'Tỷ lệ vi phạm (%)', 'LoaiNghiepVu': 'Nghiệp vụ'},
            color='TyLeViPham', color_continuous_scale='Purples'
        )
        fig_nv.update_traces(texttemplate='%{text:.2f}%', textposition='outside', marker_line_width=0, opacity=0.85)
        fig_nv.update_layout(
            uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-25,
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_nv, use_container_width=True)

    st.markdown("---")
    st.markdown("#### 👤 Danh sách Cán bộ có số lần Vi phạm SLA cần lưu ý")
    
    df_cb_agg = df_filtered.groupby(['MaSo', 'HoTen', 'CoSo', 'ViTri'])['ViPham'].agg(
        TongGiaoDich='count', 
        SoLanViPham='sum'
    ).reset_index()
    
    df_cb_agg['TyLeViPham (%)'] = (df_cb_agg['SoLanViPham'] / df_cb_agg['TongGiaoDich'] * 100).round(2)
    df_cb_violation = df_cb_agg[df_cb_agg['SoLanViPham'] > 0].sort_values(by=['SoLanViPham', 'TyLeViPham (%)'], ascending=False)

    if df_cb_violation.empty:
        st.success("🎉 Tuyệt vời! Không có cán bộ nào vi phạm SLA trong bộ lọc này.")
    else:
        st.dataframe(df_cb_violation, use_container_width=True, hide_index=True)


# ==================== XỬ LÝ GIAO DIỆN CHÍNH ====================

if report_type == "Báo Cáo Tổng Hợp (Tiếp Quỹ & Thu Gom)":
    st.subheader("📋 MODULE BÁO CÁO TỔNG HỢP TIẾP QUỸ & THU GOM")
    
    PATH_FILE1 = os.path.join(SHARED_DATA_DIR, "Ds_Tiepquythugom.xlsx")
    PATH_FILE2 = os.path.join(SHARED_DATA_DIR, "Ds_Noplinhkhac.xlsx")
    PATH_PARAM_KM = os.path.join(BASE_DIR, "parameters", "ThamSo_KM.xlsx")
    PATH_TEMPLATE = os.path.join(BASE_DIR, "templates", "Template_BaoCao1.xlsx")

    if st.button("⚡ Chạy và Xem Báo Cáo Tổng Hợp", type="primary"):
        if bc is None:
            st.error("❌ Không tìm thấy module `BC_th.py`. Vui lòng kiểm tra lại file code.")
            st.stop()
            
        with st.spinner("Đang kết nối hệ thống và tính toán tổng hợp số liệu..."):
            if not os.path.exists(PATH_PARAM_KM):
                st.error(f"❌ Không tìm thấy file tham số KM tại đường dẫn: {PATH_PARAM_KM}")
                st.stop()
                
            try:
                df_param_km = pd.read_excel(PATH_PARAM_KM)
            except Exception as e:
                st.error(f"❌ Lỗi khi đọc file tham số KM: {e}")
                st.stop()
                
            if not os.path.exists(PATH_FILE1) and not os.path.exists(PATH_FILE2):
                st.error(f"❌ Không tìm thấy file dữ liệu nào trong thư mục dùng chung: {SHARED_DATA_DIR}")
                st.stop()
                
            df1 = pd.DataFrame()
            df2 = pd.DataFrame()
            
            if os.path.exists(PATH_FILE1):
                df1 = bc.clean_and_prepare_file1(PATH_FILE1)
            if os.path.exists(PATH_FILE2):
                df2 = bc.clean_and_prepare_file2(PATH_FILE2)
                
            if df1.empty and df2.empty:
                st.error("❌ Tất cả dữ liệu đầu vào đều rỗng hoặc lỗi cấu trúc!")
            else:
                df_all = pd.concat([df1, df2], ignore_index=True)
                
                all_coso = ["Cơ sở 1", "Cơ sở 2", "Hub NVC", "CS HĐ", "CS HP", "Toàn trung tâm"]
                
                raw_nghiep_vu = df_all['LoaiNghiepVu'].dropna().unique()
                all_nghiep_vu = [nv for nv in raw_nghiep_vu if nv != 'Bỏ qua']
                
                extra_nv_summary = ['Tổng Thu', 'Tổng Chi']
                for nv_sum in extra_nv_summary:
                    if nv_sum not in all_nghiep_vu:
                        all_nghiep_vu.append(nv_sum)
                
                df_prev_res = bc.aggregate_period(df_all, start_prev, end_prev, df_param_km)
                df_curr_res = bc.aggregate_period(df_all, start_curr, end_curr, df_param_km)
                
                final_table = bc.format_summary_dataframe(df_prev_res, df_curr_res, all_nghiep_vu, all_coso)
                
                st.subheader("📋 Bảng Tổng Hợp Kết Quả Tiếp Quỹ & Thu Gom")
                
                string_columns_th = ['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong', 'Loại nghiệp vụ', 'Cơ sở']
                numeric_cols_th = [
                    c for c in final_table.columns 
                    if c not in string_columns_th and not str(c).startswith('Unnamed')
                ]
                
                format_dict_th = {}
                for col in numeric_cols_th:
                    if col in final_table.columns:
                        format_dict_th[col] = lambda x: "{:,.0f}".format(float(x)) if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else x

                st.dataframe(final_table.style.format(format_dict_th), use_container_width=True)
                
                thang_nam_str = start_curr.strftime("%m%Y")
                ten_file_excel = f"BC_{thang_nam_str}.xlsx"
                path_output_dong = os.path.join(OUTPUT_DIR, ten_file_excel)
                
                try:
                    bc.write_to_template(final_table, PATH_TEMPLATE, path_output_dong)
                    
                    with open(path_output_dong, "rb") as file:
                        st.success(f"🎉 Hệ thống đã tạo file báo cáo thành công: {ten_file_excel}")
                        st.download_button(
                            label=f"📥 Tải Báo Cáo Excel ({ten_file_excel})",
                            data=file,
                            file_name=ten_file_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.warning(f"⚠ Đã tổng hợp số liệu trên Web thành công, nhưng gặp lỗi khi tạo file Excel: {e}")


elif report_type == "Báo Cáo Nộp Lĩnh NHNN (Ds_NopLinhNHNN)":
    st.subheader("🏦 MODULE: DANH SÁCH NỘP LĨNH NHNN")
    st.markdown("""
    * **Nguyên tắc xử lý:** Lọc bỏ `NHNN chi nhánh khu vực 2`, phân tách tự động `Điện tổng` / `Điện con` và gom nhóm tiền tệ ĐTC (`VND`) - KĐTC (`VND_R`).
    """)
    
    absolute_folder_path = os.path.abspath(SHARED_DATA_DIR)
    st.caption(f"📂 Thư mục dữ liệu dùng chung đang quét: `{absolute_folder_path}` (Đọc file: `Ds_NopLinhNHNN.xlsx`)")
    
    thang_nam_nhnn = start_curr.strftime("%m%Y")
    output_filename_nhnn = f"Bao_Cao_DS_NopLinh_NHNN_{thang_nam_nhnn}.xlsx"
    path_output_nhnn = os.path.join(OUTPUT_DIR, output_filename_nhnn)
    
    if st.button("🚀 Chạy và Xem Báo Cáo NHNN", type="primary"):
        source_nhnn_file = os.path.join(SHARED_DATA_DIR, "Ds_NopLinhNHNN.xlsx")
        
        if BC_NHNN is None:
            st.error("❌ Không tìm thấy module `BC_NHNN.py`. Vui lòng kiểm tra lại file code.")
        elif not os.path.exists(source_nhnn_file):
            st.error(f"❌ Không tìm thấy file dữ liệu nguồn `Ds_NopLinhNHNN.xlsx` trong thư mục `{SHARED_DATA_DIR}`.")
        else:
            with st.spinner("Đang xử lý dữ liệu báo cáo NHNN... xin vui lòng đợi."):
                try:
                    success = BC_NHNN.process_ds_nop_linh_nhnn_pipeline(
                        folder_path=SHARED_DATA_DIR,
                        start_date_kc=start_curr,
                        end_date_kc=end_curr,
                        start_date_kp=start_prev,
                        end_date_kp=end_prev,
                        output_excel_path=path_output_nhnn
                    )
                    
                    if success and os.path.exists(path_output_nhnn):
                        st.success(f"🎉 Xử lý báo cáo NHNN thành công!")
                        
                        df_result_nhnn = pd.read_excel(path_output_nhnn)
                        st.subheader("📋 Bảng Kết Quả Chi Tiết Báo Cáo NHNN")
                        
                        cols_to_format_nhnn = [
                            'Số lươt kỳ trước', 'Số lượt kỳ này', 
                            'Số bó ĐTC Kì trước', 'Số Bó ĐTC kỳ báo cáo', 
                            'Số tiền ĐTC kỳ trước', 'Số tiền ĐTC kỳ báo cáo', 
                            'Số bó KĐTC Kì trước', 'Số Bó KĐTC kỳ báo cáo', 
                            'Số tiền KĐTC kỳ trước', 'Số tiền KĐTC kỳ báo cáo'
                        ]
                        
                        format_dict_nhnn = {}
                        for col in cols_to_format_nhnn:
                            if col in df_result_nhnn.columns:
                                format_dict_nhnn[col] = lambda x: "{:,.0f}".format(float(x)) if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else x
                        
                        st.dataframe(df_result_nhnn.style.format(format_dict_nhnn), use_container_width=True)
                        
                        with open(path_output_nhnn, "rb") as f:
                            st.download_button(
                                label=f"📥 Tải Xuống File Excel Báo Cáo NHNN ({output_filename_nhnn})",
                                data=f,
                                file_name=output_filename_nhnn,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.warning(f"⚠️ Không có dữ liệu thỏa mãn khoảng thời gian từ `{start_curr}` đến `{end_curr}` trong file `Ds_NopLinhNHNN.xlsx`.")
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {str(e)}")


elif report_type == "Báo Cáo Hoạt Động ATM (BC_atm)":
    st.subheader("🏧 MODULE BÁO CÁO HOẠT ĐỘNG ATM & CRM")
    st.markdown("""
    * **Nguyên tắc xử lý:** Lấy 8 ký tự bên phải cột `Diễn giải` làm `ID máy`, map với file `ThamSo_ATM.xlsx`, phân tách cấu trúc cha (`Cơ sở 1`, `Cơ sở 2`, `Hub NVC`, `CS HĐ`, `CS HP`, `Toàn trung tâm`) và con (`Tổng`, `ATM`, `CRM`) theo đúng yêu cầu nghiệp vụ.
    """)
    
    PATH_NHATKY = os.path.join(SHARED_DATA_DIR, "BC_nhatkygiaodich.xlsx")
    PATH_PARAM_ATM = os.path.join(BASE_DIR, "parameters", "ThamSo_ATM.xlsx")
    
    thang_nam_str = start_curr.strftime("%m%Y")
    output_filename_atm = f"BC_atm_{thang_nam_str}.xlsx"
    path_output_atm = os.path.join(OUTPUT_DIR, output_filename_atm)
    
    if st.button("🚀 Chạy và Xuất Báo Cáo ATM", type="primary"):
        if BC_ATM is None:
            st.error("❌ Không tìm thấy module `BC_ATM.py`. Vui lòng kiểm tra lại file code.")
        elif not os.path.exists(PATH_NHATKY):
            st.error(f"❌ Không tìm thấy file dữ liệu nguồn `BC_nhatkygiaodich.xlsx` trong thư mục `{SHARED_DATA_DIR}`.")
        elif not os.path.exists(PATH_PARAM_ATM):
            st.error(f"❌ Không tìm thấy file tham số `ThamSo_ATM.xlsx` trong thư mục `parameters`.")
        else:
            with st.spinner("Đang tính toán các chỉ tiêu báo cáo ATM... xin vui lòng đợi."):
                try:
                    df_result_atm = BC_ATM.process_atm_summary_report(
                        path_nhatky=PATH_NHATKY,
                        path_param=PATH_PARAM_ATM,
                        start_curr=start_curr,
                        end_curr=end_curr,
                        start_prev=start_prev,
                        end_prev=end_prev,
                        output_excel_path=path_output_atm
                    )
                    
                    if df_result_atm is not None and not df_result_atm.empty:
                        st.success("🎉 Xử lý và tạo file `BC_atm.xlsx` thành công!")
                        
                        st.subheader("📋 Bảng Kết Quả Chi Tiết Báo Cáo ATM")
                        
                        format_dict_atm = {}
                        for col in df_result_atm.columns:
                            if col not in ['Cơ sở', 'Phân loại']:
                                format_dict_atm[col] = lambda x: "{:,.0f}".format(float(x)) if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else x
                        
                        st.dataframe(df_result_atm.style.format(format_dict_atm, na_rep=""), use_container_width=True)
                        
                        with open(path_output_atm, "rb") as f:
                            st.download_button(
                                label=f"📥 Tải Xuống File Excel `BC_atm` ({output_filename_atm})",
                                data=f,
                                file_name=output_filename_atm,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.warning(f"⚠️ Không có dữ liệu thỏa mãn trong khoảng thời gian đã chọn.")
                except Exception as e:
                    st.error(f"❌ Đã xảy ra lỗi trong quá trình xử lý báo cáo ATM: {str(e)}")


elif report_type == "Giao Dịch Phát Sinh Lớn (Ds_Tiepquythugom)":
    st.subheader("🔥 MODULE LỌC GIAO DỊCH PHÁT SINH LỚN (KỲ NÀY & KỲ TRƯỚC)")
    st.markdown("Trích xuất các giao dịch có giá trị lớn từ file `Ds_Tiepquythugom.xlsx` phân tách theo VND và USD.")
    
    top_n_choice = st.slider("Chọn số lượng giao dịch lớn cần lấy (Top N theo mỗi nghiệp vụ/cơ sở):", min_value=1, max_value=50, value=10)
    
    if st.button("🚀 Chạy Lọc Giao Dịch Phát Sinh Lớn", type="primary"):
        if giao_dich_lon is None:
            st.error("❌ Không tìm thấy module `giao_dich_lon.py` trong hệ thống.")
        else:
            with st.spinner("Đang lọc và tổng hợp giao dịch phát sinh lớn cho cả 2 kỳ..."):
                (vnd_curr, usd_curr), (vnd_prev, usd_prev) = giao_dich_lon.process_giao_dich_lon_report(
                    shared_data_dir=SHARED_DATA_DIR,
                    start_curr=start_curr,
                    end_curr=end_curr,
                    start_prev=start_prev,
                    end_prev=end_prev,
                    top_n=top_n_choice
                )
                
                main_tab1, main_tab2 = st.tabs(["📊 Kỳ Báo Cáo (Hiện Tại)", "📊 Kỳ Trước (So Sánh)"])
                
                with main_tab1:
                    sub_t1, sub_t2 = st.tabs(["💵 VND Kỳ Này", "💵 USD Kỳ Này"])
                    with sub_t1:
                        if not vnd_curr.empty:
                            st.dataframe(vnd_curr.style.format({'Số tiền': '{:,.2f}'}), use_container_width=True, hide_index=True)
                        else:
                            st.info("Không có dữ liệu VND Kỳ này.")
                    with sub_t2:
                        if not usd_curr.empty:
                            st.dataframe(usd_curr.style.format({'Số tiền': '{:,.2f}'}), use_container_width=True, hide_index=True)
                        else:
                            st.info("Không có dữ liệu USD Kỳ này.")
                            
                with main_tab2:
                    sub_t3, sub_t4 = st.tabs(["💵 VND Kỳ Trước", "💵 USD Kỳ Trước"])
                    with sub_t3:
                        if not vnd_prev.empty:
                            st.dataframe(vnd_prev.style.format({'Số tiền': '{:,.2f}'}), use_container_width=True, hide_index=True)
                        else:
                            st.info("Không có dữ liệu VND Kỳ trước.")
                    with sub_t4:
                        if not usd_prev.empty:
                            st.dataframe(usd_prev.style.format({'Số tiền': '{:,.2f}'}), use_container_width=True, hide_index=True)
                        else:
                            st.info("Không có dữ liệu USD Kỳ trước.")
                
                excel_data = giao_dich_lon.export_to_excel(vnd_curr, usd_curr, vnd_prev, usd_prev)
                st.download_button(
                    label="📥 Tải Báo Cáo Giao Dịch Lớn (4 Sheet: VND & USD cho Kỳ Này & Kỳ Trước .xlsx)",
                    data=excel_data,
                    file_name=f"Giao_Dich_Phat_Sinh_Lon_{start_curr.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


elif report_type == "Báo Cáo XNK Theo Loại Tiền & Đối Tác":
    st.subheader("💱 MODULE BÁO CÁO XUẤT NHẬP KHẨU THEO LOẠI TIỀN & ĐỐI TÁC")
    st.markdown("""
    * **Chức năng:** Tổng hợp và so sánh số liệu Xuất Nhập Khẩu theo **Loại XNK**, **Đối tác** và **Loại tiền** giữa Kỳ báo cáo và Kỳ trước.
    """)
    
    if st.button("🚀 Chạy và Phân Tích Báo Cáo XNK", type="primary"):
        if bc_xnk is None:
            st.error("❌ Không tìm thấy module `bc_xnk.py`. Vui lòng đặt file `bc_xnk.py` vào cùng thư mục.")
        else:
            with st.spinner("Đang xử lý dữ liệu XNK và tính toán so sánh các kỳ..."):
                try:
                    df_pivot_curr, df_pivot_prev, df_diff, path_output = bc_xnk.xu_ly_bao_cao_xnk_pipeline(
                        shared_data_dir=SHARED_DATA_DIR,
                        output_dir=OUTPUT_DIR,
                        start_curr=start_curr,
                        end_curr=end_curr,
                        start_prev=start_prev,
                        end_prev=end_prev
                    )
                    
                    st.success("🎉 Xử lý và tổng hợp báo cáo XNK thành công!")
                    
                    tab_xnk1, tab_xnk2, tab_xnk3 = st.tabs(["📊 Kỳ Báo Cáo (Hiện Tại)", "📊 Kỳ Trước (So Sánh)", "📈 Chênh Lệch"])
                    with tab_xnk1:
                        st.dataframe(df_pivot_curr, use_container_width=True, hide_index=True)
                    with tab_xnk2:
                        st.dataframe(df_pivot_prev, use_container_width=True, hide_index=True)
                    with tab_xnk3:
                        st.dataframe(df_diff, use_container_width=True, hide_index=True)
                        
                    with open(path_output, "rb") as f:
                        st.download_button(
                            label="📥 Tải Xuống File Báo Cáo XNK (.xlsx)",
                            data=f,
                            file_name=os.path.basename(path_output),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"❌ Lỗi xử lý báo cáo XNK: {str(e)}")


elif report_type == "Báo Cáo SLA Kho Quỹ":
    if show_sla_report_page is None:
        st.error("❌ Không tìm thấy module `sla_report.py`. Vui lòng kiểm tra lại file code.")
    else:
        show_sla_report_page()


elif report_type == "Báo Cáo KPI Cán Bộ":
    if kpi_module is None:
        st.error("❌ Không tìm thấy module `kpi_module.py`. Vui lòng tạo file `kpi_module.py` trong thư mục dự án.")
    else:
        if hasattr(kpi_module, "show_kpi_report_page"):
            kpi_module.show_kpi_report_page()
            
            if "df_master_result" in st.session_state and st.session_state.df_master_result is not None:
                param_dir = os.path.join(BASE_DIR, "parameters")
                render_sla_violation_dashboard(
                    st.session_state.df_master_result, 
                    SHARED_DATA_DIR, 
                    param_dir, 
                    start_curr, 
                    end_curr
                )
        else:
            st.warning("⚠ Module `kpi_module.py` chưa có hàm `show_kpi_report_page()`. Vui lòng bổ sung hàm hiển thị.")

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>Hệ thống báo cáo nội bộ tự động hóa • Tích hợp hiển thị Web & Tải file</p>", unsafe_allow_html=True)
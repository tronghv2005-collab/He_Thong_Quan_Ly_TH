import streamlit as st
import pandas as pd
from datetime import datetime, date
import calendar
import os

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

st.title("📊 HỆ THỐNG XỬ LÝ VÀ XUẤT BÁO CÁO TỰ ĐỘNG")
st.markdown("---")

# ==================== TÍNH TOÁN NGÀY MẶC ĐỊNH (TỰ ĐỘNG THEO THÁNG) ====================
today = date.today()

# Kỳ báo cáo (Tháng này): Từ ngày 01 đến ngày hiện tại
default_kc_start = date(today.year, today.month, 1)
default_kc_end = today

# Kỳ trước (Tháng liền trước): Toàn bộ tháng trước
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
        "Báo Cáo SLA Kho Quỹ",
        "Báo Cáo KPI Cán Bộ" 
    )
)

st.sidebar.markdown("---")
st.sidebar.subheader("📅 Thiết lập khoảng thời gian")

# Khung chọn thời gian Kỳ trước
st.sidebar.markdown("**1. Kỳ Trước (So sánh)**")
start_prev = st.sidebar.date_input("Từ ngày (Kỳ trước)", value=default_kp_start)
end_prev = st.sidebar.date_input("Đến ngày (Kỳ trước)", value=default_kp_end)

# Khung chọn thời gian Kỳ báo cáo
st.sidebar.markdown("**2. Kỳ Báo Cáo (Hiện tại)**")
start_curr = st.sidebar.date_input("Từ ngày (Kỳ này)", value=default_kc_start)
end_curr = st.sidebar.date_input("Đến ngày (Kỳ này)", value=default_kc_end)


# ==================== ĐỊNH NGHĨA THƯ MỤC DÙNG CHUNG ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SHARED_DATA_DIR = os.path.join(BASE_DIR, "data_source")
OUTPUT_DIR = os.path.join(BASE_DIR, "templates")

# Tự động tạo các thư mục nếu chưa tồn tại
os.makedirs(SHARED_DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


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
                
                all_coso = ["Cơ sở 1", "Cơ sở 2", "HUB NVC", "CS HĐ", "CS HP", "Toàn trung tâm"]
                
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
    * **Nguyên tắc xử lý:** Lấy 8 ký tự bên phải cột `Diễn giải` làm `ID máy`, map với file `ThamSo_ATM.xlsx`, phân tách cấu trúc cha (`Cơ sở 1`, `Cơ sở 2`, `HUB NVC`, `CS HĐ`, `CS HP`, `Toàn trung tâm`) và con (`Tổng`, `ATM`, `CRM`) theo đúng yêu cầu nghiệp vụ.
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
        else:
            st.warning("⚠ Module `kpi_module.py` chưa có hàm `show_kpi_report_page()`. Vui lòng bổ sung hàm hiển thị.")

# Phần chân trang
st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>Hệ thống báo cáo nội bộ tự động hóa • Tích hợp hiển thị Web & Tải file</p>", unsafe_allow_html=True)
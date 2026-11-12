import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import date
import calendar
import os

# IMPORT CÁC MODULE NGHIỆP VỤ XỬ LÝ TRÊN RAM
import DS_tiepquythugom
import DS_atm
import ds_noplinhkhac
import ds_noplinhnhnn
import ds_kiemdem
import th_kpi  # Tích hợp module tính toán Thủ kho và Cán bộ hỗ trợ

def render_kpi_time_filter():
    st.subheader("⏱ Thiết lập Kỳ Đánh Giá KPI")
    col1, col2 = st.columns(2)
    with col1:
        kpi_mode = st.radio("Chọn hình thức đánh giá:", ["Theo Tháng", "Theo Quý"], horizontal=True, key="kpi_mode_radio")
    with col2:
        today = date.today()
        years_list = list(range(today.year - 3, today.year + 2))
        selected_year = st.selectbox("Chọn Năm:", years_list, index=years_list.index(today.year) if today.year in years_list else 3, key="kpi_year_select")

    if kpi_mode == "Theo Tháng":
        months_dict = {f"Tháng {i}": i for i in range(1, 13)}
        selected_month_str = st.selectbox("Chọn Tháng:", list(months_dict.keys()), index=today.month - 1, key="kpi_month_select")
        selected_month = months_dict[selected_month_str]
        start_date = date(selected_year, selected_month, 1)
        last_day = int(calendar.monthrange(selected_year, selected_month)[1])
        end_date = date(selected_year, selected_month, last_day)
    else:
        quarters_dict = {
            "Quý 1 (Tháng 1 - Tháng 3)": (1, 3),
            "Quý 2 (Tháng 4 - Tháng 6)": (4, 6),
            "Quý 3 (Tháng 7 - Tháng 9)": (7, 9),
            "Quý 4 (Tháng 10 - Tháng 12)": (10, 12)
        }
        selected_quarter_str = st.selectbox("Chọn Quý:", list(quarters_dict.keys()), key="kpi_quarter_select")
        start_m, end_m = quarters_dict[selected_quarter_str]
        start_date = date(selected_year, start_m, 1)
        last_day_of_end_m = int(calendar.monthrange(selected_year, end_m)[1])
        end_date = date(selected_year, end_m, last_day_of_end_m)

    st.info(f"📅 Khoảng thời gian quét: Từ ngày **{start_date.strftime('%d/%m/%Y')}** đến ngày **{end_date.strftime('%d/%m/%Y')}**")
    return start_date, end_date


def export_kpi_to_excel(df_master):
    """Xuất dữ liệu KPI Master ra file Excel định dạng chuyên nghiệp trực tiếp từ RAM."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_master.to_excel(writer, sheet_name='Báo cáo KPI Master', index=False)
        workbook  = writer.book
        worksheet = writer.sheets['Báo cáo KPI Master']
        
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F4E78', 'font_color': 'white', 'border': 1
        })
        data_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        number_format = workbook.add_format({'border': 1, 'valign': 'vcenter', 'align': 'right', 'num_format': '#,##0.00'})
        
        for col_num, column_title in enumerate(df_master.columns):
            worksheet.write(0, col_num, column_title, header_format)
            
        for i, col in enumerate(df_master.columns):
            max_len = max(df_master[col].astype(str).map(len).max(), len(str(col))) + 3
            if df_master[col].dtype in ['float64', 'int64'] and col != 'STT':
                worksheet.set_column(i, i, max_len, number_format)
            else:
                worksheet.set_column(i, i, max_len, data_format)
                
        worksheet.freeze_panes(1, 0)
        worksheet.set_row(0, 28)
    return output.getvalue()


def get_master_kpi_report(shared_data_dir, param_dir, start_date, end_date):
    start_str = start_date.strftime('%Y-%m-%d')
    end_str = end_date.strftime('%Y-%m-%d')

    # BƯỚC 1: XÂY DỰNG TRỤC NHÂN SỰ GỐC
    df_tq_emp = DS_tiepquythugom.build_master_employee_pipeline(shared_data_dir, start_str, end_str)
    df_atm_emp = DS_atm.build_atm_employee_pipeline(shared_data_dir, start_str, end_str)
    
    df_master_raw = pd.concat([df_tq_emp, df_atm_emp], ignore_index=True)
    if df_master_raw.empty:
        return pd.DataFrame()

    df_master_raw['MaSo'] = df_master_raw['MaSo'].astype(str).str.strip().str.lower()
    df_master_raw['HoTen'] = df_master_raw['HoTen'].astype(str).str.strip().str.title()
    df_master_raw['CoSo'] = df_master_raw['CoSo'].astype(str).str.strip()
    df_master_raw['ViTri'] = df_master_raw['ViTri'].astype(str).str.strip()

    df_emp_grouped = df_master_raw.groupby(['MaSo', 'HoTen', 'CoSo', 'ViTri'], as_index=False).agg({'NgayCong': 'sum'})
    df_emp_sorted = df_emp_grouped.sort_values(by=['MaSo', 'NgayCong'], ascending=[True, False])
    df_master_clean = df_emp_sorted.drop_duplicates(subset=['MaSo', 'ViTri'], keep='first').copy()

    # BƯỚC 2: TẬP HỢP CÁC CHỈ SỐ KPI TIẾP QUỸ
    df_kpi_tq = DS_tiepquythugom.process_tiepquy_thugom_matrix_with_params(shared_data_dir, param_dir, start_str, end_str)
    tq_cols = ['SoLuot_DCN', 'SoPhong_DCN', 'GDC', 'SoLuot_TN', 'SoPhong_TN', 'GDC_TN']
    
    if not df_kpi_tq.empty:
        df_kpi_tq['MaSo'] = df_kpi_tq['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_tq['ViTri'] = df_kpi_tq['ViTri'].astype(str).str.strip()
        group_keys_tq = ['MaSo', 'ViTri']
        if 'CoSo' in df_kpi_tq.columns:
            df_kpi_tq['CoSo'] = df_kpi_tq['CoSo'].astype(str).str.strip()
            group_keys_tq = ['MaSo', 'CoSo', 'ViTri']
        df_kpi_tq_clean = df_kpi_tq.groupby(group_keys_tq, as_index=False)[tq_cols].sum()
    else:
        df_kpi_tq_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + tq_cols)

    # BƯỚC 3: TẬP HỢP CÁC CHỈ SỐ KPI ATM
    df_kpi_atm = DS_atm.process_atm_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    atm_cols = ['SoMay_Tiep', 'GDC_Tiep', 'SoMay_Sua', 'GDC_Sua', 'SoMay_BD', 'GDC_BD']

    if not df_kpi_atm.empty:
        df_kpi_atm['MaSo'] = df_kpi_atm['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_atm['ViTri'] = df_kpi_atm['ViTri'].astype(str).str.strip()
        df_kpi_atm_clean = df_kpi_atm.groupby(['MaSo', 'ViTri'], as_index=False)[atm_cols].sum()
    else:
        df_kpi_atm_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + atm_cols)

    # BƯỚC 4: TẬP HỢP CÁC CHỈ SỐ KPI NỘP LĨNH KHÁC
    df_kpi_noplinh = ds_noplinhkhac.process_noplinhkhac_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    noplinh_cols = ['SoLuot_TCH', 'SoDiem_TCH', 'GDC_TCH', 'SoLuot_TCTD', 'So_TCTD', 'GDC_TCTD', 'SoLuot_NopLinhCS', 'GDC_NopLinhCS']

    if not df_kpi_noplinh.empty:
        df_kpi_noplinh['MaSo'] = df_kpi_noplinh['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_noplinh['ViTri'] = df_kpi_noplinh['ViTri'].astype(str).str.strip()
        df_kpi_noplinh_clean = df_kpi_noplinh.groupby(['MaSo', 'ViTri'], as_index=False)[noplinh_cols].sum()
    else:
        df_kpi_noplinh_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + noplinh_cols)

    # BƯỚC 5: TẬP HỢP CÁC CHỈ SỐ KPI NỘP LĨNH NHNN
    df_kpi_nhnn = ds_noplinhnhnn.process_noplinhnhnn_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    nhnn_cols = ['SoLuot_NopLinhNHNN', 'GDC_NHNN']

    if not df_kpi_nhnn.empty:
        df_kpi_nhnn['MaSo'] = df_kpi_nhnn['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_nhnn['ViTri'] = df_kpi_nhnn['ViTri'].astype(str).str.strip()
        df_kpi_nhnn_clean = df_kpi_nhnn.groupby(['MaSo', 'ViTri'], as_index=False)[nhnn_cols].sum()
    else:
        df_kpi_nhnn_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + nhnn_cols)

    # BƯỚC 6: GHÉP NỐI (LEFT JOIN) CÁC DỮ LIỆU KPI VÀO TRỤC NHÂN SỰ CHÍNH
    df_final = pd.merge(df_master_clean, df_kpi_tq_clean, on=['MaSo', 'ViTri'], how='left')
    df_final = pd.merge(df_final, df_kpi_atm_clean, on=['MaSo', 'ViTri'], how='left')
    df_final = pd.merge(df_final, df_kpi_noplinh_clean, on=['MaSo', 'ViTri'], how='left')
    df_final = pd.merge(df_final, df_kpi_nhnn_clean, on=['MaSo', 'ViTri'], how='left')
    
    all_kpi_cols = tq_cols + atm_cols + noplinh_cols + nhnn_cols
    for col in all_kpi_cols:
        if col not in df_final.columns:
            df_final[col] = 0.0
    df_final[all_kpi_cols] = df_final[all_kpi_cols].fillna(0.0)

    # BƯỚC 7: TẬP HỢP VÀ TÍNH TOÁN DỮ LIỆU KIỂM ĐÊM
    df_kpi_kd = ds_kiemdem.process_kiemdem_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    if not df_kpi_kd.empty:
        df_kpi_kd['MaSo'] = df_kpi_kd['MaSo'].astype(str).str.strip().str.lower()
        df_final = pd.merge(df_final, df_kpi_kd, on=['MaSo'], how='left')
    else:
        for c in ['MG_Lon', 'MG_Nho', 'NgoaiTe']:
            df_final[c] = 0.0

    kd_raw_cols = ['MG_Lon', 'MG_Nho', 'NgoaiTe']
    for col in kd_raw_cols:
        if col not in df_final.columns:
            df_final[col] = 0.0
    df_final[kd_raw_cols] = df_final[kd_raw_cols].fillna(0.0)

    f2_path = ds_kiemdem._find_file_by_keyword(param_dir, "thamso_nsld") or ds_kiemdem._find_file_by_keyword(shared_data_dir, "thamso_nsld")
    df2_param = ds_kiemdem._read_flexible_file(f2_path, header_row=0)
    p_dict = {}
    if not df2_param.empty:
        chitieu_col = ds_kiemdem._get_flexible_column(df2_param, ['chỉ tiêu', 'chi tieu', 'ten chi tieu'])
        gdc_col = ds_kiemdem._get_flexible_column(df2_param, ['gdc', 'hệ số', 'he so', 'gia tri'])
        if chitieu_col and gdc_col:
            for _, r in df2_param.iterrows():
                p_dict[str(r[chitieu_col]).strip().upper()] = ds_kiemdem._safe_float(r[gdc_col])

    t_vnd_cs1 = p_dict.get('TYLEVNDCS1', 0.0)
    t_vnd_cs2 = p_dict.get('TYLEVNDCS2', 0.0)
    t_vnd_nvc = p_dict.get('TYLEVNDNVC', 0.0)
    t_vnd_hd  = p_dict.get('TYLEVNDHD', 0.0)
    t_vnd_hp  = p_dict.get('TYLEVNDHP', 0.0)

    t_nt_cs1  = p_dict.get('TYLENTCS1', 0.0)
    t_nt_cs2  = p_dict.get('TYLENTCS2', 0.0)
    t_nt_nvc  = p_dict.get('TYLENTNVC', 0.0)
    t_nt_hd   = p_dict.get('TYLENTHD', 0.0)
    t_nt_hp   = p_dict.get('TYLENTHP', 0.0)

    h_kd_vnd_lon = p_dict.get('KIEMDEMVNDLON', 0.0)
    h_kd_vnd_nho = p_dict.get('KIEMDEMVNDNHO', 0.0)
    h_kd_nt_lon  = p_dict.get('KIEMDEMNTLON', 0.0)
    h_kd_nt_nho  = p_dict.get('KIEMDEMNTNHO', 0.0)

    def calc_kiemdem_metrics(row):
        coso = str(row.get('CoSo', '')).strip()
        mg_lon = row.get('MG_Lon', 0.0)
        mg_nho = row.get('MG_Nho', 0.0)
        nt = row.get('NgoaiTe', 0.0)

        if coso == 'Cơ sở 1':
            t_vnd, t_nt = t_vnd_cs1, t_nt_cs1
        elif coso == 'Cơ sở 2':
            t_vnd, t_nt = t_vnd_cs2, t_nt_cs2
        elif coso == 'Hub NVC':
            t_vnd, t_nt = t_vnd_nvc, t_nt_nvc
        elif coso == 'CS HĐ':
            t_vnd, t_nt = t_vnd_hd, t_nt_hd
        elif coso == 'CS HP':
            t_vnd, t_nt = t_vnd_hp, t_nt_hp
        else:
            t_vnd, t_nt = 0.0, 0.0

        kd_mg_lon = mg_lon * t_vnd
        kd_mg_nho = mg_nho + mg_lon * (1.0 - t_vnd)
        kd_nt_mg_lon = nt * t_nt
        kd_nt_mg_nho = nt * (1.0 - t_nt)

        gdc_kd = (kd_mg_lon * h_kd_vnd_lon) + (kd_mg_nho * h_kd_vnd_nho) + \
                 (kd_nt_mg_lon * h_kd_nt_lon) + (kd_nt_mg_nho * h_kd_nt_nho)

        return pd.Series({
            'KD_MG_Lon': kd_mg_lon, 'KD_MG_Nho': kd_mg_nho,
            'KD_NT_MG_Lon': kd_nt_mg_lon, 'KD_NT_MG_Nho': kd_nt_mg_nho,
            'GDC_KiemDem': gdc_kd
        })

    kd_results = df_final.apply(calc_kiemdem_metrics, axis=1)
    for col_res in kd_results.columns:
        df_final[col_res] = kd_results[col_res]

    # BƯỚC 8: TÍNH TOÁN BỔ SUNG 5 CỘT CHO THỦ KHO & CÁN BỘ HỖ TRỢ (TỪ th_kpi)
    df_th_result = th_kpi.process_thukho_hotro_kpi_matrix(shared_data_dir, param_dir, start_str, end_str, df_final)
    
    if not df_th_result.empty:
        df_final['MaSo_Key'] = df_final['MaSo'].astype(str).str.strip().str.lower()
        df_final['ViTri_Key'] = df_final['ViTri'].astype(str).str.strip().str.upper()
        df_th_result['MaSo_Key'] = df_th_result['MaSo'].astype(str).str.strip().str.lower()
        df_th_result['ViTri_Key'] = df_th_result['ViTri'].astype(str).str.strip().str.upper()
        
        df_final = pd.merge(
            df_final, 
            df_th_result[['MaSo_Key', 'ViTri_Key', 'SoBoTien_XL', 'GDC_XLTien', 'GDC_CongTac_DCN', 'GDC_Khac', 'TongCong']], 
            on=['MaSo_Key', 'ViTri_Key'], 
            how='left',
            suffixes=('', '_th')
        )
        df_final.drop(columns=['MaSo_Key', 'ViTri_Key'], inplace=True, errors='ignore')

    for c in ['SoBoTien_XL', 'GDC_XLTien', 'GDC_CongTac_DCN', 'GDC_Khac', 'TongCong']:
        if c not in df_final.columns:
            df_final[c] = 0.0
        else:
            df_final[c] = df_final[c].fillna(0.0)

    # BƯỚC 9: SẮP XẾP THỨ TỰ BÁO CÁO & XUẤT KẾT QUẢ
    valid_coso_list = ["Cơ sở 1", "Cơ sở 2", "Hub NVC", "CS HĐ", "CS HP"]
    position_order = ["Thủ quỹ", "GDV NQC", "GDV ATM", "Thủ kho", "Kiểm soát ATM", "Chấm báo cáo", "Lái xe"]
    
    df_final['CoSo_Cat'] = pd.Categorical(df_final['CoSo'], categories=valid_coso_list, ordered=True)
    df_final['ViTri_Cat'] = pd.Categorical(df_final['ViTri'], categories=position_order, ordered=True)
    df_final.sort_values(by=['CoSo_Cat', 'ViTri_Cat', 'HoTen'], inplace=True, na_position='last')
    
    target_order = [
        'HoTen', 'CoSo', 'ViTri', 'NgayCong',
        'SoLuot_DCN', 'SoPhong_DCN', 'GDC', 
        'SoLuot_TN', 'SoPhong_TN', 'GDC_TN',
        'SoMay_Tiep', 'GDC_Tiep', 
        'SoMay_Sua', 'GDC_Sua', 
        'SoMay_BD', 'GDC_BD',
        'SoLuot_TCH', 'SoDiem_TCH', 'GDC_TCH',
        'SoLuot_TCTD', 'So_TCTD', 'GDC_TCTD',
        'SoLuot_NopLinhCS', 'GDC_NopLinhCS',
        'SoLuot_NopLinhNHNN', 'GDC_NHNN',
        'KD_MG_Lon', 'KD_MG_Nho', 'KD_NT_MG_Lon', 'KD_NT_MG_Nho', 'GDC_KiemDem',
        'SoBoTien_XL', 'GDC_XLTien', 'GDC_CongTac_DCN', 'GDC_Khac', 'TongCong'
    ]
    return df_final[target_order].reset_index(drop=True)


def show_kpi_report_page():
    st.subheader("🎯 MODULE ĐÁNH GIÁ VÀ TỔNG HỢP KPI CÁN BỘ (TIẾP QUỸ, ATM, NỘP LĨNH, NHNN, KIỂM ĐÊM & THỦ KHO)")
    
    start_date, end_date = render_kpi_time_filter()
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    shared_data_dir = os.path.join(base_dir, "data_source")
    param_dir = os.path.join(base_dir, "parameters")
    
    if "df_master_result" not in st.session_state:
        st.session_state.df_master_result = None

    if st.button("🚀 Khởi Tạo Báo Cáo KPI Master (Tổng hợp toàn diện)", type="primary"):
        with st.spinner("Đang chạy tính toán và tổng hợp dữ liệu toàn diện... Vui lòng đợi..."):
            df_master = get_master_kpi_report(shared_data_dir, param_dir, start_date, end_date)
            
            if df_master.empty:
                st.warning("⚠️ Không tìm thấy dữ liệu báo cáo phát sinh trong khoảng thời gian này.")
                st.session_state.df_master_result = None
            else:
                if 'STT' not in df_master.columns:
                    df_master.insert(0, 'STT', range(1, len(df_master) + 1))
                st.session_state.df_master_result = df_master
                st.success(f"🎉 Hoàn tất tổng hợp KPI đồng bộ cho **{len(df_master)}** dòng nhân sự!")

    if st.session_state.df_master_result is not None:
        df_display = st.session_state.df_master_result
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        try:
            excel_data = export_kpi_to_excel(df_display)
            st.download_button(
                label="📥 Tải Báo Cáo KPI Master (.xlsx)",
                data=excel_data,
                file_name=f"Bao_Cao_KPI_Master_{start_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"⚠️ Không thể định dạng Excel nâng cao. Chi tiết lỗi: {e}")

if __name__ == "__main__":
    st.set_page_config(layout="wide")
    show_kpi_report_page()
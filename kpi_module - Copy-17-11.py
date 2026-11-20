import pandas as pd
import numpy as np
import os

import DS_tiepquythugom
import DS_atm
import ds_noplinhkhac
import ds_noplinhnhnn
import ds_kiemdem
import th_kpi

def get_master_kpi_report(shared_data_dir, param_dir, start_date, end_date):
    start_str = start_date.strftime('%Y-%m-%d')
    end_str = end_date.strftime('%Y-%m-%d')

    # ==========================================
    # BƯỚC 1: XÂY DỰNG TRỤC NHÂN SỰ GỐC & XỬ LÝ NGÀY CÔNG (CHỐNG DOUBLE TỪ 2 FILE TẠM)
    # ==========================================
    
    # File tạm 1: Lấy từ Tiếp quỹ thu gom
    df_tq_raw = DS_tiepquythugom.build_master_employee_pipeline(shared_data_dir, start_str, end_str)
    if not df_tq_raw.empty:
        df_tq_raw['MaSo'] = df_tq_raw['MaSo'].astype(str).str.strip().str.lower()
        df_tq_raw['HoTen'] = df_tq_raw['HoTen'].astype(str).str.strip().str.title()
        df_tq_raw['CoSo'] = df_tq_raw['CoSo'].astype(str).str.strip()
        df_tq_raw['ViTri'] = df_tq_raw['ViTri'].astype(str).str.strip()
        if 'NgayCong' in df_tq_raw.columns:
            df_tq_raw['NgayCong'] = pd.to_numeric(df_tq_raw['NgayCong'], errors='coerce').fillna(0.0)
        else:
            df_tq_raw['NgayCong'] = 0.0

    # File tạm 2: Lấy từ ATM
    df_atm_raw = DS_atm.build_atm_employee_pipeline(shared_data_dir, start_str, end_str)
    if not df_atm_raw.empty:
        df_atm_raw['MaSo'] = df_atm_raw['MaSo'].astype(str).str.strip().str.lower()
        df_atm_raw['HoTen'] = df_atm_raw['HoTen'].astype(str).str.strip().str.title()
        df_atm_raw['CoSo'] = df_atm_raw['CoSo'].astype(str).str.strip()
        df_atm_raw['ViTri'] = df_atm_raw['ViTri'].astype(str).str.strip()
        if 'NgayCong' in df_atm_raw.columns:
            df_atm_raw['NgayCong'] = pd.to_numeric(df_atm_raw['NgayCong'], errors='coerce').fillna(0.0)
        else:
            df_atm_raw['NgayCong'] = 0.0

    # Xử lý chống double ngày công: Cùng cán bộ, cùng vị trí nếu đã có ở file tạm 1 thì file tạm 2 không đếm ngày công nữa
    if not df_tq_raw.empty and not df_atm_raw.empty:
        tq_keys = set(zip(df_tq_raw['MaSo'], df_tq_raw['ViTri']))
        
        def filter_atm_ngaycong(row):
            if (row['MaSo'], row['ViTri']) in tq_keys:
                return 0.0
            return row['NgayCong']
        
        df_atm_raw['NgayCong'] = df_atm_raw.apply(filter_atm_ngaycong, axis=1)

    # Gộp 2 nguồn làm trục nhân sự chung
    df_master_raw = pd.concat([df_tq_raw, df_atm_raw], ignore_index=True)
    if df_master_raw.empty:
        df_master_raw = pd.DataFrame(columns=['MaSo', 'HoTen', 'CoSo', 'ViTri', 'NgayCong'])

    # Chuẩn hóa cuối cùng cho trục nhân sự
    df_master_clean = df_master_raw.groupby(['MaSo', 'HoTen', 'CoSo', 'ViTri'], as_index=False).agg({'NgayCong': 'sum'})
    df_master_clean = df_master_clean.drop_duplicates(subset=['MaSo'], keep='first').copy()

    # ==========================================
    # BƯỚC 2: TẬP HỢP KPI TIẾP QUỸ THU GOM
    # ==========================================
    tq_cols = ['SoLuot_DCN', 'SoPhong_DCN', 'GDC', 'SoLuot_TN', 'SoPhong_TN', 'GDC_TN']
    df_kpi_tq_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + tq_cols)
    
    df_kpi_tq = DS_tiepquythugom.process_tiepquy_thugom_matrix_with_params(shared_data_dir, param_dir, start_str, end_str)
    if not df_kpi_tq.empty:
        df_kpi_tq['MaSo'] = df_kpi_tq['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_tq['ViTri'] = df_kpi_tq['ViTri'].astype(str).str.strip()
        group_keys_tq = ['MaSo', 'ViTri']
        if 'CoSo' in df_kpi_tq.columns:
            df_kpi_tq['CoSo'] = df_kpi_tq['CoSo'].astype(str).str.strip()
            group_keys_tq = ['MaSo', 'CoSo', 'ViTri']
        df_kpi_tq_clean = df_kpi_tq.groupby(group_keys_tq, as_index=False)[tq_cols].sum()

    # ==========================================
    # BƯỚC 3: TẬP HỢP KPI ATM
    # ==========================================
    atm_cols = ['SoMay_Tiep', 'GDC_Tiep', 'SoMay_Sua', 'GDC_Sua', 'SoMay_BD', 'GDC_BD']
    df_kpi_atm_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + atm_cols)
    
    df_kpi_atm = DS_atm.process_atm_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    if not df_kpi_atm.empty:
        df_kpi_atm['MaSo'] = df_kpi_atm['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_atm['ViTri'] = df_kpi_atm['ViTri'].astype(str).str.strip()
        df_kpi_atm_clean = df_kpi_atm.groupby(['MaSo', 'ViTri'], as_index=False)[atm_cols].sum()

    # ==========================================
    # BƯỚC 4: TẬP HỢP KPI NỘP LĨNH KHÁC
    # ==========================================
    noplinh_cols = ['SoLuot_TCH', 'SoDiem_TCH', 'GDC_TCH', 'SoLuot_TCTD', 'So_TCTD', 'GDC_TCTD', 'SoLuot_NopLinhCS', 'GDC_NopLinhCS']
    df_kpi_noplinh_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + noplinh_cols)
    
    df_kpi_noplinh = ds_noplinhkhac.process_noplinhkhac_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    if not df_kpi_noplinh.empty:
        df_kpi_noplinh['MaSo'] = df_kpi_noplinh['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_noplinh['ViTri'] = df_kpi_noplinh['ViTri'].astype(str).str.strip()
        df_kpi_noplinh_clean = df_kpi_noplinh.groupby(['MaSo', 'ViTri'], as_index=False)[noplinh_cols].sum()

    # ==========================================
    # BƯỚC 5: TẬP HỢP KPI NỘP LĨNH NHNN
    # ==========================================
    nhnn_cols = ['SoLuot_NopLinhNHNN', 'GDC_NHNN']
    df_kpi_nhnn_clean = pd.DataFrame(columns=['MaSo', 'ViTri'] + nhnn_cols)
    
    df_kpi_nhnn = ds_noplinhnhnn.process_noplinhnhnn_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    if not df_kpi_nhnn.empty:
        df_kpi_nhnn['MaSo'] = df_kpi_nhnn['MaSo'].astype(str).str.strip().str.lower()
        df_kpi_nhnn['ViTri'] = df_kpi_nhnn['ViTri'].astype(str).str.strip()
        df_kpi_nhnn_clean = df_kpi_nhnn.groupby(['MaSo', 'ViTri'], as_index=False)[nhnn_cols].sum()

    # ==========================================
    # BƯỚC 6: MERGE DỮ LIỆU VÀO TRỤC CHÍNH
    # ==========================================
    df_final = pd.merge(df_master_clean, df_kpi_tq_clean, on=['MaSo', 'ViTri'], how='left')
    df_final = pd.merge(df_final, df_kpi_atm_clean, on=['MaSo', 'ViTri'], how='left')
    df_final = pd.merge(df_final, df_kpi_noplinh_clean, on=['MaSo', 'ViTri'], how='left')
    df_final = pd.merge(df_final, df_kpi_nhnn_clean, on=['MaSo', 'ViTri'], how='left')
    
    all_kpi_cols = tq_cols + atm_cols + noplinh_cols + nhnn_cols
    for col in all_kpi_cols:
        if col not in df_final.columns:
            df_final[col] = 0.0
    df_final[all_kpi_cols] = df_final[all_kpi_cols].fillna(0.0)

    # ==========================================
    # BƯỚC 7: XỬ LÝ KIỂM ĐÊM
    # ==========================================
    df_kpi_kd = ds_kiemdem.process_kiemdem_kpi_matrix(shared_data_dir, param_dir, start_str, end_str)
    if not df_kpi_kd.empty:
        df_kpi_kd['MaSo'] = df_kpi_kd['MaSo'].astype(str).str.strip().str.lower()
        df_final = pd.merge(df_final, df_kpi_kd, on=['MaSo'], how='left')

    kd_raw_cols = ['MG_Lon', 'MG_Nho', 'NgoaiTe']
    for col in kd_raw_cols:
        if col not in df_final.columns:
            df_final[col] = 0.0
    df_final[kd_raw_cols] = df_final[kd_raw_cols].fillna(0.0)

    f2_path = ds_kiemdem._find_file_by_keyword(param_dir, "thamso_nsld") or ds_kiemdem._find_file_by_keyword(shared_data_dir, "thamso_nsld")
    df2_param = ds_kiemdem._read_flexible_file(f2_path, header_row=0) if f2_path else pd.DataFrame()
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

    # ==========================================
    # BƯỚC 8: THỦ KHO VÀ CÁN BỘ HỖ TRỢ
    # ==========================================
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

    # ==========================================
    # BƯỚC 9: SẮP XẾP VÀ XUẤT KẾT QUẢ
    # ==========================================
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
    
    existing_target_order = [col for col in target_order if col in df_final.columns]
    df_result = df_final[existing_target_order].reset_index(drop=True)

    numeric_cols = df_result.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        if col != 'STT':
            df_result[col] = df_result[col].round(0)

    return df_result
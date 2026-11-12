import pandas as pd
import numpy as np
import os

def _find_file_by_keyword(folder_path, keyword):
    """Tìm file trong thư mục dựa theo từ khóa không phân biệt hoa thường."""
    if not os.path.exists(folder_path):
        return None
    for f in os.listdir(folder_path):
        if keyword.lower() in f.lower() and f.endswith(('.xlsx', '.xls')):
            return os.path.join(folder_path, f)
    return None

def _read_flexible_file(file_path, header_row=1):
    """Đọc file Excel linh hoạt theo dòng tiêu đề chỉ định."""
    if not file_path or not os.path.exists(file_path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(file_path, header=header_row)
    except Exception:
        try:
            df = pd.read_excel(file_path, header=0)
        except Exception:
            return pd.DataFrame()
    
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _extract_id(val):
    """Lấy dãy số/ký tự trước dấu '-' để làm mã nhân viên duy nhất."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if '-' in val_str:
        return val_str.split('-')[0].strip().lower()
    return val_str.lower()

def process_money_columns(df):
    """
    Tính tổng số tờ tiền từ các cột mệnh giá chuẩn. 
    Loại bỏ hoàn toàn các cột chứa từ khóa 'thành tiền' để tránh cộng nhầm giá trị tiền VNĐ.
    """
    if df.empty:
        return pd.Series(0.0, index=df.index)
    
    standard_denominations = [
        '500,000', '200,000', '100,000', '50,000', '20,000', '10,000', 
        '5,000', '2,000', '1,000', '500', '200', '100', '50', '20', '10', '5', '2', '1'
    ]
    
    total_sheets = pd.Series(0.0, index=df.index)
    for col_name in df.columns.unique():
        col_str = str(col_name).strip()
        col_lower = col_str.lower()
        
        if 'thành tiền' in col_lower or 'thanh tien' in col_lower:
            continue
            
        if any(d in col_str for d in standard_denominations):
            sub_col_data = df[col_name]
            if isinstance(sub_col_data, pd.DataFrame):
                for c_idx in range(sub_col_data.shape[1]):
                    series_data = sub_col_data.iloc[:, c_idx]
                    total_sheets += pd.to_numeric(series_data, errors='coerce').fillna(0)
            else:
                total_sheets += pd.to_numeric(sub_col_data, errors='coerce').fillna(0)
            
    return total_sheets

def process_thukho_hotro_kpi_matrix(shared_data_dir, param_dir, start_date_str, end_date_str, df_master_emp):
    f_tq_path = _find_file_by_keyword(shared_data_dir, "Ds_tiepquythugom")
    f_khac_path = _find_file_by_keyword(shared_data_dir, "Ds_Noplinhkhac")
    f_nhnn_path = _find_file_by_keyword(shared_data_dir, "Ds_NoplinhNHNN")
    f_param_path = _find_file_by_keyword(param_dir, "Thamso_NSLD") or _find_file_by_keyword(shared_data_dir, "Thamso_NSLD")

    df_f1 = _read_flexible_file(f_tq_path, header_row=1)
    df_f2 = _read_flexible_file(f_khac_path, header_row=1)
    df_f3 = _read_flexible_file(f_nhnn_path, header_row=1)

    start_dt = pd.to_datetime(start_date_str)
    end_dt = pd.to_datetime(end_date_str)

    # Lọc thời gian theo kỳ cho cả 3 file
    for df in [df_f1, df_f2, df_f3]:
        if not df.empty:
            date_col = next((c for c in df.columns if 'ngày hiệu lực' in c.lower() or 'ngay hieu luc' in c.lower()), None)
            if date_col:
                date_series = df[date_col]
                if isinstance(date_series, pd.DataFrame):
                    date_series = date_series.iloc[:, 0]
                df['__temp_date'] = pd.to_datetime(date_series, errors='coerce')
                df.drop(df[(df['__temp_date'] < start_dt) | (df['__temp_date'] > end_dt)].index, inplace=True)
                df.drop(columns=['__temp_date'], inplace=True, errors='ignore')

    # Đọc file tham số định mức
    param_dict = {}
    if f_param_path and os.path.exists(f_param_path):
        df_param = pd.read_excel(f_param_path, header=0)
        df_param.columns = [str(c).strip() for c in df_param.columns]
        c_chitieu = next((c for c in df_param.columns if 'chỉ tiêu' in c.lower()), None)
        c_gdc = next((c for c in df_param.columns if 'gdc' in c.lower()), None)
        if c_chitieu and c_gdc:
            for _, r in df_param.iterrows():
                key = str(r[c_chitieu]).strip().upper()
                val = pd.to_numeric(r[c_gdc], errors='coerce')
                if not pd.isna(key):
                    param_dict[key] = val if not pd.isna(val) else 0.0

    dm_thukho_nhan = param_dict.get('THUKHONHAN', 0.0)
    dm_thuquy_hotro = param_dict.get('THUQUYHOTRO', 0.0)
    dm_thuquy_hotro_nhnn = param_dict.get('THUQUYHOTRONHNN', 0.0)
    
    dm_tq = param_dict.get('DAUCUOINGAYTQ', 0.0)
    dm_tk = param_dict.get('DAUCUOINGAYTK', 0.0)
    dm_kiem = param_dict.get('DAUCUOINGAYTKKIEM', 0.0)
    dm_gdv = param_dict.get('DAUCUOINGAYGDV', 0.0)
    dm_gdv_atm = param_dict.get('DAUCUOINGAYGDVATM', 0.0)
    dm_ks_atm = param_dict.get('DAUCUOINGAYKSATM', 0.0)

    results = []
    if df_master_emp is None or df_master_emp.empty:
        return pd.DataFrame(columns=['MaSo', 'ViTri', 'SoBoTien_XL', 'GDC_XLTien', 'GDC_CongTac_DCN', 'GDC_Khac', 'TongCong'])

    for _, row in df_master_emp.iterrows():
        raw_maso = row.get('MaSo', '')
        ma_so = _extract_id(raw_maso)
        
        co_so = str(row.get('CoSo', '')).strip().lower()
        vi_tri = str(row.get('ViTri', '')).strip().upper()
        so_ngay_cong = float(row.get('NgayCong', 0.0))

        so_bo_xl = 0.0
        gdc_xl_tien = 0.0
        gdc_cong_tac_dcn = 0.0
        gdc_khac = 0.0

        if 'THỦ KHO' in vi_tri:
            # 1. Lọc File 1 (Tiếp quỹ thủ gom)
            def filter_thukho_f1(df):
                if df.empty: return pd.DataFrame()
                c_coso = next((c for c in df.columns if 'cơ sở' in c.lower() or 'coso' in c.lower()), None)
                c_thukho = next((c for c in df.columns if 'thủ kho' in c.lower() or 'thukho' in c.lower()), None)
                if not c_coso or not c_thukho: return pd.DataFrame()
                
                s_coso = df[c_coso].iloc[:, 0] if isinstance(df[c_coso], pd.DataFrame) else df[c_coso]
                s_thukho = df[c_thukho].iloc[:, 0] if isinstance(df[c_thukho], pd.DataFrame) else df[c_thukho]

                match_coso = s_coso.astype(str).str.strip().str.lower() == co_so
                match_id = s_thukho.apply(_extract_id) == ma_so
                return df[match_coso & match_id]

            # 2. Lọc File 2 (Nộp lĩnh khác)
            def filter_thukho_f2(df):
                if df.empty: return pd.DataFrame()
                c_coso = next((c for c in df.columns if 'cơ sở' in c.lower() or 'coso' in c.lower()), None)
                c_thukho = next((c for c in df.columns if 'thủ kho' in c.lower() or 'thukho' in c.lower()), None)
                c_nghiepvu = next((c for c in df.columns if 'loại nghiệp vụ' in c.lower() or 'nghiep vu' in c.lower()), None)
                c_nguoicam = next((c for c in df.columns if 'người cầm tiền' in c.lower() or 'nguoi cam' in c.lower()), None)
                
                if not c_coso: return pd.DataFrame()
                s_coso = df[c_coso].iloc[:, 0] if isinstance(df[c_coso], pd.DataFrame) else df[c_coso]
                match_coso = s_coso.astype(str).str.strip().str.lower() == co_so

                match_tk_chinh = pd.Series(False, index=df.index)
                if c_thukho:
                    s_thukho = df[c_thukho].iloc[:, 0] if isinstance(df[c_thukho], pd.DataFrame) else df[c_thukho]
                    match_tk_chinh = (s_thukho.apply(_extract_id) == ma_so)

                match_chitien_camtien = pd.Series(False, index=df.index)
                if c_nghiepvu and c_nguoicam:
                    s_nv = df[c_nghiepvu].iloc[:, 0] if isinstance(df[c_nghiepvu], pd.DataFrame) else df[c_nghiepvu]
                    s_nc = df[c_nguoicam].iloc[:, 0] if isinstance(df[c_nguoicam], pd.DataFrame) else df[c_nguoicam]
                    
                    is_chitien = s_nv.astype(str).str.strip().str.lower().str.contains('chi tiền', na=False)
                    is_camtien = s_nc.apply(_extract_id) == ma_so
                    match_chitien_camtien = is_chitien & is_camtien

                return df[match_coso & (match_tk_chinh | match_chitien_camtien)]

            # 3. Lọc File 3 (Nộp lĩnh NHNN) cho Thủ kho: Chỉ lấy Loại điện = 'Điện tổng'
            def filter_thukho_f3(df):
                if df.empty: return pd.DataFrame()
                c_coso = next((c for c in df.columns if 'cơ sở' in c.lower() or 'coso' in c.lower()), None)
                c_thukho = next((c for c in df.columns if 'thủ kho' in c.lower() or 'thukho' in c.lower()), None)
                c_loaidien = next((c for c in df.columns if 'loại điện' in c.lower() or 'loai dien' in c.lower()), None)
                
                if not c_coso: return pd.DataFrame()
                s_coso = df[c_coso].iloc[:, 0] if isinstance(df[c_coso], pd.DataFrame) else df[c_coso]
                match_coso = s_coso.astype(str).str.strip().str.lower() == co_so

                match_tk = pd.Series(False, index=df.index)
                if c_thukho:
                    s_thukho = df[c_thukho].iloc[:, 0] if isinstance(df[c_thukho], pd.DataFrame) else df[c_thukho]
                    match_tk = (s_thukho.apply(_extract_id) == ma_so)

                match_dientong = pd.Series(True, index=df.index)
                if c_loaidien:
                    s_dien = df[c_loaidien].iloc[:, 0] if isinstance(df[c_loaidien], pd.DataFrame) else df[c_loaidien]
                    match_dientong = s_dien.astype(str).str.strip().str.lower() == 'điện tổng'

                return df[match_coso & match_tk & match_dientong]

            sub_f1 = filter_thukho_f1(df_f1)
            sub_f2 = filter_thukho_f2(df_f2)
            sub_f3 = filter_thukho_f3(df_f3)
            
            sheets_f1 = process_money_columns(sub_f1)
            sheets_f2 = process_money_columns(sub_f2)
            sheets_f3 = process_money_columns(sub_f3)
            
            so_bo_xl = (sheets_f1.sum() + sheets_f2.sum() + sheets_f3.sum()) / 1000.0
            gdc_xl_tien = so_bo_xl * dm_thukho_nhan
            so_bo_xl =round(so_bo_xl ,0)
            
            has_gdv_nqc = False
            if not df_master_emp.empty and 'CoSo' in df_master_emp.columns and 'ViTri' in df_master_emp.columns:
                coso_match_emp = df_master_emp[df_master_emp['CoSo'].astype(str).str.strip().str.lower() == co_so]
                has_gdv_nqc = any('GDV NQC' in str(v).strip().upper() for v in coso_match_emp['ViTri'])

            gdc_cong_tac_dcn = so_ngay_cong * (dm_kiem if has_gdv_nqc else dm_tk)

        elif 'THỦ QUỸ' in vi_tri or 'HỖ TRỢ' in vi_tri:
            # Đối với Cán bộ hỗ trợ: Lấy từ File 2 và File 3 (File 3 lấy cả Điện tổng và Điện con như bình thường)
            def filter_hotro(df):
                if df.empty: return pd.DataFrame()
                c_coso = next((c for c in df.columns if 'cơ sở' in c.lower() or 'coso' in c.lower()), None)
                c_hs1 = next((c for c in df.columns if 'hỗ trợ 1' in c.lower() or 'hotro 1' in c.lower()), None)
                c_hs2 = next((c for c in df.columns if 'hỗ trợ 2' in c.lower() or 'hotro 2' in c.lower()), None)
                if not c_coso: return pd.DataFrame()

                s_coso = df[c_coso].iloc[:, 0] if isinstance(df[c_coso], pd.DataFrame) else df[c_coso]
                match_coso = s_coso.astype(str).str.strip().str.lower() == co_so
                match_hs = pd.Series(False, index=df.index)

                if c_hs1:
                    s_hs1 = df[c_hs1].iloc[:, 0] if isinstance(df[c_hs1], pd.DataFrame) else df[c_hs1]
                    match_hs |= (s_hs1.apply(_extract_id) == ma_so)
                if c_hs2:
                    s_hs2 = df[c_hs2].iloc[:, 0] if isinstance(df[c_hs2], pd.DataFrame) else df[c_hs2]
                    match_hs |= (s_hs2.apply(_extract_id) == ma_so)
                
                return df[match_coso & match_hs]

            sub_f2_hq = filter_hotro(df_f2)
            sub_f3_hq = filter_hotro(df_f3) # Lấy cả điện tổng và điện con cho cán bộ hỗ trợ

            def calc_hq_sheets(sub_df, dinh_muc):
                if sub_df.empty: return 0.0, 0.0
                c_hs1 = next((c for c in sub_df.columns if 'hỗ trợ 1' in c.lower()), None)
                c_hs2 = next((c for c in sub_df.columns if 'hỗ trợ 2' in c.lower()), None)
                
                total_b = 0.0
                total_g = 0.0
                for _, r in sub_df.iterrows():
                    id1 = ""
                    if c_hs1:
                        val1 = r[c_hs1]
                        if isinstance(val1, pd.Series): val1 = val1.iloc[0]
                        id1 = _extract_id(val1)

                    id2 = ""
                    if c_hs2:
                        val2 = r[c_hs2]
                        if isinstance(val2, pd.Series): val2 = val2.iloc[0]
                        id2 = _extract_id(val2)

                    has_both = (id1 != '') and (id2 != '')
                    he_so = 0.5 if has_both else 1.0
                    
                    row_df = pd.DataFrame([r.to_dict()])
                    row_sheets = process_money_columns(row_df)
                    bo = (row_sheets.sum() / 1000.0) * he_so
                    total_b += round(bo,0)
                    total_g += bo * dinh_muc
                return total_b, total_g

            bo_f2, gdc_f2 = calc_hq_sheets(sub_f2_hq, dm_thuquy_hotro)
            bo_f3, gdc_f3 = calc_hq_sheets(sub_f3_hq, dm_thuquy_hotro_nhnn)

            so_bo_xl = bo_f2 + bo_f3
            gdc_xl_tien = gdc_f2 + gdc_f3
            gdc_cong_tac_dcn = so_ngay_cong * dm_tq

        elif 'GDV ATM' in vi_tri:
            gdc_cong_tac_dcn = so_ngay_cong * dm_gdv_atm
        elif 'GDV' in vi_tri:
            gdc_cong_tac_dcn = so_ngay_cong * dm_gdv
        elif 'KIỂM SOÁT ATM' in vi_tri or 'KSATM' in vi_tri:
            gdc_cong_tac_dcn = so_ngay_cong * dm_ks_atm

        tong_cong = gdc_xl_tien + gdc_cong_tac_dcn + gdc_khac

        results.append({
            'MaSo': ma_so,
            'ViTri': row.get('ViTri', ''),
            'SoBoTien_XL': round(so_bo_xl, 2),
            'GDC_XLTien': round(gdc_xl_tien, 2),
            'GDC_CongTac_DCN': round(gdc_cong_tac_dcn, 2),
            'GDC_Khac': round(gdc_khac, 2),
            'TongCong': round(tong_cong, 2)
        })

    return pd.DataFrame(results)
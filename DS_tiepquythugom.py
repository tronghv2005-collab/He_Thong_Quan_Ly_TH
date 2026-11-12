
import os
import glob
import pandas as pd
import numpy as np

def _find_file_by_keyword(folder_path, keyword):
    """Hàm phụ trợ tìm file có chứa từ khóa trong tên (không phân biệt hoa thường)."""
    if not os.path.exists(folder_path):
        return None
    for f in os.listdir(folder_path):
        if keyword.lower() in f.lower() and (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.csv')):
            return os.path.join(folder_path, f)
    return None

def _read_flexible_file(file_path):
    """Đọc file Excel/CSV, tự động xử lý dòng tiêu đề (Header)."""
    if not file_path or not os.path.exists(file_path):
        return pd.DataFrame()
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, header=1)
            if df.empty or len(df.columns) <= 1:
                df = pd.read_csv(file_path, header=0)
        else:
            df = pd.read_excel(file_path, header=1)
            cols_str = [str(c).lower() for c in df.columns]
            if not any('cơ sở' in c or 'co so' in c for c in cols_str):
                df = pd.read_excel(file_path, header=0)
        return df
    except Exception as e:
        print(f"Lỗi đọc file {file_path}: {e}")
        return pd.DataFrame()

def _safe_float(val):
    """Hàm an toàn chuyển đổi giá trị sang float."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s_val = str(val).strip()
    if s_val == '' or s_val.lower() == 'nan':
        return 0.0
    try:
        return float(s_val)
    except ValueError:
        return 0.0

def build_master_employee_pipeline(shared_data_dir, start_str, end_str):
    """
    I - Tạo bảng Danh sách nhân viên tổng hợp từ CẢ FILE 1 (Ds_tiepquythugom) VÀ FILE 2 (Atm).
    Bảo đảm không bị sót nhân sự mảng ATM.
    """
    f1_path = _find_file_by_keyword(shared_data_dir, "Ds_TiepQuyThuGom") or _find_file_by_keyword(shared_data_dir, "tiepquythugom")
    f2_path = _find_file_by_keyword(shared_data_dir, "atm")
    
    df1 = _read_flexible_file(f1_path)
    df2 = _read_flexible_file(f2_path)
  
    
    start_dt = pd.to_datetime(start_str)
    end_dt = pd.to_datetime(end_str)

    records = []

    # 1. Quét nhân sự từ File Tiếp Quỹ - Thu Gom
    if not df1.empty:
        date_col = next((c for c in df1.columns if 'ngày' in str(c).lower() or 'date' in str(c).lower()), None)
        if date_col:
            df1['Ngay_Parsed'] = pd.to_datetime(df1[date_col], errors='coerce')
            df1 = df1[(df1['Ngay_Parsed'] >= start_dt) & (df1['Ngay_Parsed'] <= end_dt)]
        
        f1_roles = {'Thủ quỹ': 'Thủ quỹ', 'Lái xe': 'Lái xe', 'GDV NQC': 'GDV NQC', 'Thủ kho': 'Thủ kho'}
        
        for _, row in df1.iterrows():
            coso = str(row.get('Cơ sở', row.get('Co so', 'Chưa xác định'))).strip()
            loai_nv = str(row.get('Loại nghiệp vụ', row.get('Loai nghiep vu', ''))).strip()
            ngay = row.get('Ngay_Parsed', None)
            
            for col_name, role_name in f1_roles.items():
                val = row.get(col_name)
                if pd.notna(val) and str(val).strip() != '':
                    hoten_raw = str(val).strip()
                    if role_name == 'Lái xe' and ('không có' in hoten_raw.lower() or 'khong co' in hoten_raw.lower()):
                        continue
                    records.append({
                        'HoTen_Raw': hoten_raw,
                        'CoSo_GhiNhan': coso,
                        'ViTri': role_name,
                        'LoaiNghiepVu': loai_nv,
                        'Ngay': ngay,
                        'Source': 'File1'
                    })

    # 2. Quét nhân sự từ File ATM (BỔ SUNG ĐỂ KHÔNG BỊ SÓT)
    if not df2.empty:
        date_col2 = next((c for c in df2.columns if 'ngày' in str(c).lower() or 'date' in str(c).lower()), None)
        if date_col2:
            df2['Ngay_Parsed'] = pd.to_datetime(df2[date_col2], errors='coerce')
            df2 = df2[(df2['Ngay_Parsed'] >= start_dt) & (df2['Ngay_Parsed'] <= end_dt)]
            
        f2_roles = {
            'Thủ quỹ': 'Thủ quỹ', 
            'Lái xe': 'Lái xe', 
            'GDV ATM': 'GDV ATM', 
            'Kiểm soát ATM': 'Kiểm soát ATM', 
            'Chấm báo cáo': 'Chấm báo cáo'
        }
        
        for _, row in df2.iterrows():
            coso = str(row.get('Cơ sở', row.get('Co so', 'Chưa xác định'))).strip()
            ngay = row.get('Ngay_Parsed', None) if 'Ngay_Parsed' in df2.columns else None
            
            for col_name, role_name in f2_roles.items():
                if col_name in df2.columns:
                    val = row.get(col_name)
                    if pd.notna(val) and str(val).strip() != '':
                        hoten_raw = str(val).strip()
                        if role_name == 'Lái xe' and ('không có' in hoten_raw.lower() or 'khong co' in hoten_raw.lower()):
                            continue
                        records.append({
                            'HoTen_Raw': hoten_raw,
                            'CoSo_GhiNhan': coso,
                            'ViTri': role_name,
                            'LoaiNghiepVu': 'ATM',
                            'Ngay': ngay,
                            'Source': 'File2'
                        })

    if not records:
        return pd.DataFrame(columns=['MaSo', 'HoTen', 'CoSo', 'ViTri', 'NgayCong'])

    df_all = pd.DataFrame(records)
    
    def split_code_name(text):
        if '-' in text:
            parts = text.split('-', 1)
            return parts[0].strip().lower(), parts[1].strip()
        return text.strip().lower(), text.strip()

    df_all[['MaSo', 'HoTen']] = df_all['HoTen_Raw'].apply(lambda x: pd.Series(split_code_name(x)))

    # Xác định cơ sở gốc ưu tiên từ Tiếp Quỹ / Thu Gom hoặc từ File ATM
    f1_df = df_all[df_all['Source'] == 'File1'].copy()
    goc_df = f1_df[f1_df['LoaiNghiepVu'].isin(['Tiếp quỹ đầu ngày', 'Thu gom cuối ngày'])]
    
    home_map = {}
    if not goc_df.empty:
        for _, r in goc_df.iterrows():
            home_map[r['MaSo']] = r['CoSo_GhiNhan']
            
    if not home_map and not f1_df.empty:
        for _, r in f1_df.iterrows():
            if r['MaSo'] not in home_map:
                home_map[r['MaSo']] = r['CoSo_GhiNhan']

    # Nếu nhân sự chỉ làm việc ở file ATM mà không có trong file Tiếp quỹ
    for _, r in df_all.iterrows():
        if r['MaSo'] not in home_map:
            home_map[r['MaSo']] = r['CoSo_GhiNhan']

    df_all['CoSo'] = df_all['MaSo'].map(home_map).fillna(df_all['CoSo_GhiNhan'])

    df_valid_days = df_all.dropna(subset=['Ngay']).drop_duplicates(subset=['MaSo', 'HoTen', 'CoSo', 'ViTri', 'Ngay'])
    df_ngay_cong = df_valid_days.groupby(['MaSo', 'HoTen', 'CoSo', 'ViTri'], as_index=False).size()
    df_ngay_cong.rename(columns={'size': 'NgayCong'}, inplace=True)

    if df_ngay_cong.empty:
        df_ngay_cong = df_all.groupby(['MaSo', 'HoTen', 'CoSo', 'ViTri'], as_index=False).size()
        df_ngay_cong.rename(columns={'size': 'NgayCong'}, inplace=True)
        df_ngay_cong['NgayCong'] = 1.0

    return df_ngay_cong[['MaSo', 'HoTen', 'CoSo', 'ViTri', 'NgayCong']].copy()

def process_tiepquy_thugom_matrix_with_params(shared_data_dir, param_dir, start_str, end_str):
    """
    II - Tính KPI cho cán bộ. Tách bạch hoàn toàn GDC ĐCN và GDC TN, không gộp nhầm.
    """
    f1_path = _find_file_by_keyword(shared_data_dir, "ds_tiepquythugom") or _find_file_by_keyword(shared_data_dir, "tiepquythugom")
    f3_path = _find_file_by_keyword(param_dir, "thamso_km") or _find_file_by_keyword(shared_data_dir, "thamso_km")
    
    df1 = _read_flexible_file(f1_path)
    df3 = _read_flexible_file(f3_path)
    
    start_dt = pd.to_datetime(start_str)
    end_dt = pd.to_datetime(end_str)

    kpi_output_cols = ['MaSo', 'ViTri', 'SoLuot_DCN', 'SoPhong_DCN', 'GDC', 'SoLuot_TN', 'SoPhong_TN', 'GDC_TN']

    if df1.empty:
        return pd.DataFrame(columns=kpi_output_cols)

    date_col = next((c for c in df1.columns if 'ngày' in str(c).lower() or 'date' in str(c).lower()), None)
    if date_col:
        df1['Ngay_Parsed'] = pd.to_datetime(df1[date_col], errors='coerce')
        df1 = df1[(df1['Ngay_Parsed'] >= start_dt) & (df1['Ngay_Parsed'] <= end_dt)]

    param_dict = {}
    if not df3.empty:
        df3.columns = [str(c).strip() for c in df3.columns]
        ten_pgd_col = next((c for c in df3.columns if 'khachhang' in c.lower()), None)
        if ten_pgd_col:
            gdc_param_cols = ['GDCtn', 'GDCtuyendn', 'GDCtuyencn', 'GDCphongdn', 'GDCphongcn', 'GDCthuongdn', 'GDCthuongcn']
            for _, r in df3.iterrows():
                pgd_name = str(r[ten_pgd_col]).strip()
                param_dict[pgd_name] = {}
                for col in gdc_param_cols:
                    val = r.get(col, 0)
                    param_dict[pgd_name][col] = _safe_float(val)

    loai_nv_col = next((c for c in df1.columns if 'loại nghiệp vụ' in str(c).lower() or 'loai nghiep vu' in str(c).lower()), 'Loại nghiệp vụ')
    coso_col = next((c for c in df1.columns if 'cơ sở' in str(c).lower() or 'co so' in str(c).lower()), 'Cơ sở')
    tuyen_col = next((c for c in df1.columns if 'tuyến' in str(c).lower() or 'tuyen' in str(c).lower()), 'Tuyến')
    kh_col = next((c for c in df1.columns if 'khách hàng' in str(c).lower() or 'khach hang' in str(c).lower()), 'Tên khách hàng')

    home_map = {}
    if loai_nv_col in df1.columns:
        goc_df1 = df1[df1[loai_nv_col].isin(['Tiếp quỹ đầu ngày', 'Thu gom cuối ngày'])]
        for _, r in goc_df1.iterrows():
            f1_roles = ['Thủ quỹ', 'Lái xe', 'GDV NQC', 'Thủ kho']
            for role in f1_roles:
                val = r.get(role)
                if pd.notna(val) and str(val).strip() != '':
                    raw = str(val).strip()
                    if role == 'Lái xe' and ('không có' in raw.lower() or 'khong co' in raw.lower()):
                        continue
                    m_id = raw.split('-')[0].strip().lower() if '-' in raw else raw.lower()
                    home_map[m_id] = str(r.get(coso_col, '')).strip()

    records_kpi = []

    for _, row in df1.iterrows():
        loai_nv = str(row.get(loai_nv_col, '')).strip().lower()
        
        is_dn = 'đầu ngày' in loai_nv or 'dau ngay' in loai_nv
        is_cn = 'cuối ngày' in loai_nv or 'cuoi ngay' in loai_nv
        is_dcn = is_dn or is_cn
        is_tn = 'trong ngày' in loai_nv or 'trong ngay' in loai_nv or 'tn' in loai_nv
        
        kh_name = str(row.get(kh_col, '')).strip()
        p_vals = param_dict.get(kh_name, {
            'GDCtn': 0.0, 'GDCtuyendn': 0.0, 'GDCtuyencn': 0.0, 
            'GDCphongdn': 0.0, 'GDCphongcn': 0.0, 'GDCthuongdn': 0.0, 'GDCthuongcn': 0.0
        })
        
        gdc_tuyendn = p_vals['GDCtuyendn'] if is_dn else 0.0
        gdc_tuyencn = p_vals['GDCtuyencn'] if is_cn else 0.0
        gdc_phongdn = p_vals['GDCphongdn'] if is_dn else 0.0
        gdc_phongcn = p_vals['GDCphongcn'] if is_cn else 0.0
        gdc_thuongdn = p_vals['GDCthuongdn'] if is_dn else 0.0
        gdc_thuongcn = p_vals['GDCthuongcn'] if is_cn else 0.0
        gdc_tn = p_vals['GDCtn'] if is_tn else 0.0

        gdc_room_dcn = 0.0
        if is_dn:
            gdc_room_dcn = gdc_phongdn + gdc_thuongdn
        elif is_cn:
            gdc_room_dcn = gdc_phongcn + gdc_thuongcn

        ngay_val = row.get('Ngay_Parsed', None)
        tuyen_val = str(row.get(tuyen_col, '')).strip()

        target_kpi_roles = ['Thủ quỹ', 'Lái xe']
        for r in target_kpi_roles:
            val_nv = row.get(r)
            if pd.notna(val_nv) and str(val_nv).strip() != '':
                raw_name = str(val_nv).strip()
                
                if r == 'Lái xe' and ('không có' in raw_name.lower() or 'khong co' in raw_name.lower()):
                    continue

                maso = raw_name.split('-')[0].strip().lower() if '-' in raw_name else raw_name.lower()
                
                coso_hien_tai = str(row.get(coso_col, '')).strip()
                coso_goc = home_map.get(maso, coso_hien_tai)

                thu_quy_val = str(row.get('Thủ quỹ', '')).strip()
                lai_xe_val = str(row.get('Lái xe', '')).strip()

                records_kpi.append({
                    'MaSo': maso,
                    'CoSo': coso_goc,
                    'ViTri': r,
                    'Ngay': ngay_val,
                    'LoaiNghiepVu': row.get(loai_nv_col, ''),
                    'Tuyen': tuyen_val,
                    'KhachHang': kh_name,
                    'ThuQuy': thu_quy_val,
                    'LaiXe': lai_xe_val,
                    'GDCtuyendn': gdc_tuyendn,
                    'GDCtuyencn': gdc_tuyencn,
                    'GDCphongdn': gdc_phongdn,
                    'GDCphongcn': gdc_phongcn,
                    'GDCthuongdn': gdc_thuongdn,
                    'GDCthuongcn': gdc_thuongcn,
                    'GDCtn': gdc_tn,
                    'GDC_Room_DCN': gdc_room_dcn,
                    'Is_DCN': is_dcn,
                    'Is_TN': is_tn
                })

    if not records_kpi:
        return pd.DataFrame(columns=kpi_output_cols)

    df_kpi_raw = pd.DataFrame(records_kpi)

    # 1. Số lượt ĐCN & Phòng ĐCN
    df_dcn_base = df_kpi_raw[df_kpi_raw['Is_DCN']].copy()
    
    df_dcn_luot = df_dcn_base.drop_duplicates(subset=['MaSo', 'ViTri', 'Ngay', 'LoaiNghiepVu', 'ThuQuy', 'LaiXe'])
    sum_luot_dcn = df_dcn_luot.groupby(['MaSo', 'CoSo', 'ViTri'], as_index=False)['Is_DCN'].sum()
    sum_luot_dcn.rename(columns={'Is_DCN': 'SoLuot_DCN'}, inplace=True)

    sum_phong_dcn = df_dcn_base.groupby(['MaSo', 'CoSo', 'ViTri'], as_index=False).size()
    sum_phong_dcn.rename(columns={'size': 'SoPhong_DCN'}, inplace=True)

    # 2. Số lượt TN & Phòng TN
    df_tn_base = df_kpi_raw[df_kpi_raw['Is_TN']].copy()
    
    df_tn_luot = df_tn_base.drop_duplicates(subset=['MaSo', 'ViTri', 'Ngay', 'LoaiNghiepVu', 'Tuyen'])
    sum_luot_tn = df_tn_luot.groupby(['MaSo', 'CoSo', 'ViTri'], as_index=False)['Is_TN'].sum()
    sum_luot_tn.rename(columns={'Is_TN': 'SoLuot_TN'}, inplace=True)

    sum_phong_tn = df_tn_base.groupby(['MaSo', 'CoSo', 'ViTri'], as_index=False).size()
    sum_phong_tn.rename(columns={'size': 'SoPhong_TN'}, inplace=True)

    # 3. GDC ĐCN (GDC phòng + GDC tuyến ĐCN)
    room_sum = df_dcn_base.groupby(['MaSo', 'CoSo', 'ViTri'], as_index=False)['GDC_Room_DCN'].sum()

    tuyen_unique = df_dcn_base.drop_duplicates(subset=['MaSo', 'ViTri', 'Ngay', 'LoaiNghiepVu', 'ThuQuy', 'LaiXe'])
    tuyen_sum = tuyen_unique.groupby(['MaSo', 'CoSo', 'ViTri'], as_index=False).apply(
        lambda g: pd.Series({'GDC_Tuyen_Total': (g['GDCtuyendn'] + g['GDCtuyencn']).sum()})
    ).reset_index(drop=True)

    gdc_dcn_final = pd.merge(room_sum, tuyen_sum, on=['MaSo', 'CoSo', 'ViTri'], how='outer').fillna(0.0)
    gdc_dcn_final['GDC_DCN'] = gdc_dcn_final['GDC_Room_DCN'] + gdc_dcn_final['GDC_Tuyen_Total']

    # 4. GDC TN (Chỉ tính riêng phần GDCtn của trong ngày)
    gdc_tn_final = df_tn_base.groupby(['MaSo', 'CoSo', 'ViTri'], as_index=False)['GDCtn'].sum()
    gdc_tn_final.rename(columns={'GDCtn': 'GDC_TN'}, inplace=True)

    # ĐẶC BIỆT LƯU Ý: Cột 'GDC' tổng hợp chuẩn xác chỉ bao gồm GDC_DCN (hoặc tách bạch rõ, không cộng nhầm GDC_TN vào ĐCN)
    sum_gdc = pd.merge(gdc_dcn_final[['MaSo', 'CoSo', 'ViTri', 'GDC_DCN']], gdc_tn_final, on=['MaSo', 'CoSo', 'ViTri'], how='outer').fillna(0.0)
    # GDC tổng hợp chung nếu tính trên toàn bộ bảng master có thể quy định là GDC_DCN hoặc tổng 2 loại tùy báo cáo. 
    # Theo chuẩn bảng yêu cầu: Cột 'GDC' chính là GDC Đầu/Cuối ngày, còn 'GDC_TN' là của Trong ngày.
    sum_gdc['GDC'] = sum_gdc['GDC_DCN'] 

    sum_gdc = sum_gdc[['MaSo', 'CoSo', 'ViTri', 'GDC', 'GDC_TN']]

    # 5. Tổng hợp hoàn thiện bảng kết quả
    df_final_kpi = pd.DataFrame(columns=['MaSo', 'CoSo', 'ViTri'])
    for df_part in [sum_luot_dcn, sum_phong_dcn, sum_gdc, sum_luot_tn, sum_phong_tn]:
        if not df_part.empty:
            if df_final_kpi.empty or len(df_final_kpi) == 0:
                df_final_kpi = df_part
            else:
                df_final_kpi = pd.merge(df_final_kpi, df_part, on=['MaSo', 'CoSo', 'ViTri'], how='outer')

    for col in kpi_output_cols:
        if col not in df_final_kpi.columns:
            df_final_kpi[col] = 0.0

    df_final_kpi[kpi_output_cols] = df_final_kpi[kpi_output_cols].fillna(0.0)

    return df_final_kpi[kpi_output_cols]
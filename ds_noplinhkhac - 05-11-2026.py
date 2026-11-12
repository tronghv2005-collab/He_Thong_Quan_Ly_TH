import os
import pandas as pd
import numpy as np

def _find_file_by_keyword(folder_path, keyword):
    if not os.path.exists(folder_path):
        return None
    for f in os.listdir(folder_path):
        if keyword.lower() in f.lower() and (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.csv')):
            return os.path.join(folder_path, f)
    return None

def _read_flexible_file(file_path, header_row=1):
    if not file_path or not os.path.exists(file_path):
        return pd.DataFrame()
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, header=header_row)
        else:
            df = pd.read_excel(file_path, header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
        return df
    except Exception as e:
        print(f"Lỗi đọc file {file_path}: {e}")
        return pd.DataFrame()

def _safe_float(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s_val = str(val).strip().replace(',', '.')
    if s_val == '' or s_val.lower() == 'nan':
        return 0.0
    try:
        return float(s_val)
    except ValueError:
        return 0.0

def _get_flexible_column(df, keywords):
    for col in df.columns:
        col_clean = str(col).lower().replace('_', ' ').replace('-', ' ').strip()
        for kw in keywords:
            kw_clean = str(kw).lower().replace('_', ' ').replace('-', ' ').strip()
            if kw_clean in col_clean:
                return col
    return None

def process_noplinhkhac_kpi_matrix(shared_data_dir, param_dir, start_str, end_str):
    f1_path = _find_file_by_keyword(shared_data_dir, "noplinhkhac") or _find_file_by_keyword(shared_data_dir, "ds_noplinhkhac")
    f2_path = _find_file_by_keyword(param_dir, "thamso_nsld") or _find_file_by_keyword(shared_data_dir, "thamso_nsld")
    
    df1 = _read_flexible_file(f1_path, header_row=1)
    df2 = _read_flexible_file(f2_path, header_row=0)
    
    output_cols = [
        'MaSo', 'ViTri',
        'SoLuot_TCH', 'SoDiem_TCH', 'GDC_TCH',
        'SoLuot_TCTD', 'So_TCTD', 'GDC_TCTD',
        'SoLuot_NopLinhCS', 'GDC_NopLinhCS'
    ]

    if df1.empty:
        return pd.DataFrame(columns=output_cols)

    start_dt = pd.to_datetime(start_str)
    end_dt = pd.to_datetime(end_str)

    date_col = _get_flexible_column(df1, ['ngày hiệu lực', 'ngay hieu luc', 'ngày', 'date'])
    if date_col:
        df1['Ngay_Parsed'] = pd.to_datetime(df1[date_col], errors='coerce', dayfirst=True)
        df1 = df1[(df1['Ngay_Parsed'] >= start_dt) & (df1['Ngay_Parsed'] <= end_dt)]
    else:
        df1['Ngay_Parsed'] = pd.to_datetime(start_str)

    df1['Ngay_Key'] = pd.to_datetime(df1['Ngay_Parsed']).dt.normalize().dt.strftime('%Y-%m-%d')

    param_dict = {}
    if not df2.empty:
        chitieu_col = _get_flexible_column(df2, ['chỉ tiêu', 'chi tieu', 'ten chi tieu'])
        gdc_col = _get_flexible_column(df2, ['gdc', 'hệ số', 'he so', 'gia tri'])
        if chitieu_col and gdc_col:
            for _, r in df2.iterrows():
                key_name = str(r[chitieu_col]).strip().upper()
                param_dict[key_name] = _safe_float(r[gdc_col])

    # Lấy chính xác các tham số NSLD theo đúng tên chỉ tiêu gốc
    h_hachtoanthuho = param_dict.get('HACHTOANTHUHO', 0.0)
    h_hachtoantctd  = param_dict.get('HACHTOANTCTD', 0.0)
    
    h_chihotd       = param_dict.get('CHIHOTAIDIACHIKH', 0.0)
    h_thuhotd       = param_dict.get('THUHOTAIDIACHIKH', 0.0)
    
    h_tq_tctd       = param_dict.get('TIEPQUYTCTD', 0.0)  # Dành cho Chi tiền TCTD (Thủ quỹ / Lái xe)
    h_tq_tq         = param_dict.get('THUQUYTCTD', 0.0)    # Dành cho Thu tiền TCTD (Thủ quỹ / Lái xe)

    h_nop_hp = param_dict.get('NOPLINHHUBHP', 0.0)
    h_nop_hn = param_dict.get('NOPLINHHUBHN', 0.0)

    coso_col = _get_flexible_column(df1, ['cơ sở', 'co so']) or df1.columns[0]
    tuyen_col = _get_flexible_column(df1, ['tuyến', 'tuyen'])
    nghiepvu_col = _get_flexible_column(df1, ['loại nghiệp vụ', 'loai nghiep vu', 'nghiep vu'])

    records = []
    
    role_mapping = {
        'Thủ quỹ': ['thủ quỹ', 'thu quy'],
        #'Thủ kho': ['thủ kho', 'thu kho'],
        'Lái xe': ['lái xe', 'lai xe'],
        'GDV NQC': ['gdv nqc', 'nqc', 'giao dịch viên nqc']
    }

    for _, row in df1.iterrows():
        coso_raw = str(row.get(coso_col, 'Cơ sở 1')).strip()
        ngay_key = row['Ngay_Key']
        tuyen = str(row.get(tuyen_col, '')).strip().upper()
        nghiepvu = str(row.get(nghiepvu_col, '')).strip()

        nv_lower = nghiepvu.lower()
        
        # Nhận diện nghiệp vụ
        is_tch = ('chi hộ' in nv_lower or 'thu hộ' in nv_lower or 'chi ho' in nv_lower or 'thu ho' in nv_lower)
        is_tctd = ('tctd' in nv_lower)
        is_noplinh_cs = ('cơ sở khác' in nv_lower) or ('co so khac' in nv_lower)

        gdc_noplinh_val = 0.0
        if is_noplinh_cs:
            is_hp = ('hp' in coso_raw.lower() or 'hải phòng' in coso_raw.lower() or 'hai phong' in coso_raw.lower())
            gdc_noplinh_val = h_nop_hp if is_hp else h_nop_hn

        for role_name, keywords in role_mapping.items():
            col_name = _get_flexible_column(df1, keywords)
            if col_name and col_name in df1.columns:
                val = row.get(col_name)
                if pd.notna(val) and str(val).strip() != '':
                    hoten_raw = str(val).strip()
                    if role_name == 'Lái xe' and ('không có' in hoten_raw.lower() or 'khong co' in hoten_raw.lower()):
                        continue
                    
                    if '-' in hoten_raw:
                        maso = hoten_raw.split('-', 1)[0].strip().lower()
                    else:
                        maso = hoten_raw.strip().lower()

                    # GDV NQC chỉ lấy Thu hộ/Chi hộ và TCTD, loại bỏ hoàn toàn nộp lĩnh cơ sở
                    if role_name == 'GDV NQC':
                        if not (is_tch or is_tctd):
                            continue

                    gdc_tch_val = 0.0
                    gdc_tctd_val = 0.0

                    if role_name == 'GDV NQC':
                        # GDV NQC: Thu/Chi hộ nhân HACHTOANTHUHO, TCTD nhân HACHTOANTCTD
                        if is_tch:
                            gdc_tch_val = h_hachtoanthuho
                        if is_tctd:
                            gdc_tctd_val = h_hachtoantctd
                    else:
                        # Thủ quỹ / Thủ kho / Lái xe: Áp dụng chuẩn từng chỉ tiêu Thu/Chi riêng biệt
                        if is_tch:
                            gdc_tch_val = h_chihotd if 'chi' in nv_lower else h_thuhotd
                        if is_tctd:
                            gdc_tctd_val = h_tq_tctd if 'chi' in nv_lower else h_tq_tq

                    records.append({
                        'MaSo': maso,
                        'ViTri': role_name,
                        'Ngay_Key': ngay_key,
                        'Tuyen': tuyen,
                        'NghiepVu': nghiepvu,
                        'Is_TCH': is_tch,
                        'Is_TCTD': is_tctd,
                        'Is_NopLinhCS': is_noplinh_cs,
                        'GDC_TCH_Val': gdc_tch_val,
                        'GDC_TCTD_Val': gdc_tctd_val,
                        'GDC_NopLinhCS_Val': gdc_noplinh_val
                    })

    if not records:
        return pd.DataFrame(columns=output_cols)

    df_res = pd.DataFrame(records)

    def compute_metrics(group):
        df_tch = group[group['Is_TCH']]
        so_luot_tch = len(df_tch.drop_duplicates(subset=['Ngay_Key', 'NghiepVu', 'Tuyen']))
        so_diem_tch = len(df_tch)
        gdc_tch_sum = df_tch['GDC_TCH_Val'].sum()

        df_tctd = group[group['Is_TCTD']]
        so_luot_tctd = len(df_tctd.drop_duplicates(subset=['Ngay_Key', 'NghiepVu', 'Tuyen']))
        so_tctd = len(df_tctd)
        gdc_tctd_sum = df_tctd['GDC_TCTD_Val'].sum()

        df_nop = group[group['Is_NopLinhCS']]
        so_luot_nop = len(df_nop.drop_duplicates(subset=['Ngay_Key', 'NghiepVu', 'Tuyen']))
        gdc_nop_sum = df_nop['GDC_NopLinhCS_Val'].sum()

        return pd.Series({
            'SoLuot_TCH': so_luot_tch,
            'SoDiem_TCH': so_diem_tch,
            'GDC_TCH': gdc_tch_sum,
            'SoLuot_TCTD': so_luot_tctd,
            'So_TCTD': so_tctd,
            'GDC_TCTD': gdc_tctd_sum,
            'SoLuot_NopLinhCS': so_luot_nop,
            'GDC_NopLinhCS': gdc_nop_sum
        })

    df_grouped = df_res.groupby(['MaSo', 'ViTri'], as_index=False).apply(compute_metrics)
    
    for col in output_cols:
        if col not in df_grouped.columns:
            df_grouped[col] = 0.0

    return df_grouped[output_cols].fillna(0.0)
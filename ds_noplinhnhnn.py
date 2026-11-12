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

def _read_flexible_file(file_path, header_row=2):
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

def process_noplinhnhnn_kpi_matrix(shared_data_dir, param_dir, start_str, end_str):
    f1_path = _find_file_by_keyword(shared_data_dir, "noplinhnhnn") or _find_file_by_keyword(shared_data_dir, "ds_noplinhnhnn")
    f2_path = _find_file_by_keyword(param_dir, "thamso_nsld") or _find_file_by_keyword(shared_data_dir, "thamso_nsld")
    
    df1 = _read_flexible_file(f1_path, header_row=2)
    df2 = _read_flexible_file(f2_path, header_row=0)
    
    output_cols = [
        'MaSo', 'ViTri',
        'SoLuot_NopLinhNHNN', 'GDC_NHNN'
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

    # Đọc file tham số NSLD
    param_dict = {}
    if not df2.empty:
        chitieu_col = _get_flexible_column(df2, ['chỉ tiêu', 'chi tieu', 'ten chi tieu'])
        gdc_col = _get_flexible_column(df2, ['gdc', 'hệ số', 'he so', 'gia tri'])
        if chitieu_col and gdc_col:
            for _, r in df2.iterrows():
                key_name = str(r[chitieu_col]).strip().upper()
                param_dict[key_name] = _safe_float(r[gdc_col])

    h_thq_tong_hp = param_dict.get('NHNNTHUQUYTONGHP', 0.0)
    h_thq_hp      = param_dict.get('NHNNTHUQUYHP', 0.0)
    h_lx_hp       = param_dict.get('NHNNLAIXEHP', 0.0)

    h_thq_tong    = param_dict.get('NHNNTHUQUYTONG', 0.0)
    h_thq         = param_dict.get('NHNNTHUQUY', 0.0)
    h_lx          = param_dict.get('NHNNLAIXE', 0.0)

    h_hachtoan    = param_dict.get('HACHTOANNHNN', 0.0)

    stt_col = _get_flexible_column(df1, ['stt', 'số thứ tự']) or df1.columns[0]
    coso_col = _get_flexible_column(df1, ['cơ sở', 'co so']) or df1.columns[1]
    loaidien_col = _get_flexible_column(df1, ['loại điện', 'loai dien'])
    khachhang_col = _get_flexible_column(df1, ['tên khách hàng', 'ten khach hang', 'khach hang'])
    
    # 1. Bỏ không xét tất cả các bản ghi có Tên khách hàng = NHNN chi nhánh khu vực 2
    if khachhang_col:
        df1 = df1[~df1[khachhang_col].astype(str).str.strip().str.lower().eq('nhnn chi nhánh khu vực 2')]

    if df1.empty:
        return pd.DataFrame(columns=output_cols)

    # 2. Xử lý cột STT hàng loạt bằng chuỗi vector
    def clean_stt_series(s_series):
        return s_series.astype(str).str.strip().str.split('.').str[0]

    df1['STT_Processed'] = clean_stt_series(df1[stt_col])
    df1['LoaiDien_Clean'] = df1[loaidien_col].astype(str).str.strip().str.lower() if loaidien_col else ''
    df1['Coso_Clean'] = df1[coso_col].astype(str).str.strip().str.lower() if coso_col else ''

    thug_col_name = _get_flexible_column(df1, ['thủ quỹ', 'thu quy'])
    laix_col_name = _get_flexible_column(df1, ['lái xe', 'lai xe'])

    df1['ThuQuy_Trip'] = df1[thug_col_name].astype(str).str.strip().str.lower() if thug_col_name else ''
    df1['LaiXe_Trip'] = df1[laix_col_name].astype(str).str.strip().str.lower() if laix_col_name else ''
    df1['Is_Tong'] = df1['LoaiDien_Clean'].str.contains('tổng|tong', regex=True, na=False)

    # Bước A: Lọc trùng cấp độ CHUYẾN ĐI bằng vector hóa sort và drop_duplicates
    df1 = df1.sort_values(
        by=['Ngay_Key', 'STT_Processed', 'ThuQuy_Trip', 'LaiXe_Trip', 'Is_Tong'],
        ascending=[True, True, True, True, False]
    )

    df_trip_dedup = df1.drop_duplicates(
        subset=['Ngay_Key', 'STT_Processed', 'ThuQuy_Trip', 'LaiXe_Trip'],
        keep='first'
    ).copy()

    if df_trip_dedup.empty:
        return pd.DataFrame(columns=output_cols)

    # Đánh dấu cờ Hải Phòng hàng loạt
    is_hp_series = df_trip_dedup['Coso_Clean'].str.contains('hp|hải phòng|hai phong', regex=True, na=False)
    is_tong_series = df_trip_dedup['Is_Tong']

    # Bước B: Bóc tách nhân sự đồng thời cho các vị trí (Stacking)
    role_mapping = {
        'Thủ quỹ': thug_col_name,
        'Lái xe': laix_col_name,
        'GDV NQC': _get_flexible_column(df_trip_dedup, ['gdv nqc', 'nqc', 'giao dịch viên nqc'])
    }

    stacked_records = []

    for role_name, col_name in role_mapping.items():
        if not col_name or col_name not in df_trip_dedup.columns:
            continue

        sub = df_trip_dedup[['Ngay_Key', 'STT_Processed', col_name]].copy()
        sub.columns = ['Ngay_Key', 'STT_Processed', 'HoTen_Raw']
        sub['HoTen_Raw'] = sub['HoTen_Raw'].astype(str).str.strip()
        
        mask = (sub['HoTen_Raw'].notna()) & (sub['HoTen_Raw'] != '') & (sub['HoTen_Raw'].str.lower() != 'nan')
        if role_name == 'Lái xe':
            mask &= ~sub['HoTen_Raw'].str.lower().str.contains('không có|khong co', regex=True)
            
        sub = sub[mask].copy()
        if sub.empty:
            continue

        sub['ViTri'] = role_name
        has_dash = sub['HoTen_Raw'].str.contains('-', na=False)
        sub['MaSo'] = np.where(has_dash, sub['HoTen_Raw'].str.split('-', n=1).str[0].str.strip().str.lower(), sub['HoTen_Raw'].str.strip().str.lower())

        # Gán GDC vector tương ứng theo vị trí
        p_sub = df_trip_dedup.loc[sub.index]
        is_hp_sub = is_hp_series.loc[sub.index]
        is_tong_sub = is_tong_series.loc[sub.index]

        if role_name == 'Thủ quỹ':
            gdc_val = np.where(
                is_hp_sub,
                np.where(is_tong_sub, h_thq_tong_hp, h_thq_hp),
                np.where(is_tong_sub, h_thq_tong, h_thq)
            )
        elif role_name == 'Lái xe':
            gdc_val = np.where(is_hp_sub, h_lx_hp, h_lx)
        else: # GDV NQC
            gdc_val = np.full(len(sub), h_hachtoan)

        sub['GDC_Val'] = gdc_val
        stacked_records.append(sub[['MaSo', 'ViTri', 'GDC_Val']])

    if not stacked_records:
        return pd.DataFrame(columns=output_cols)

    df_res = pd.concat(stacked_records, ignore_index=True)

    # Bước C: Tổng hợp kết quả
    df_grouped = df_res.groupby(['MaSo', 'ViTri'], as_index=False).agg(
        SoLuot_NopLinhNHNN=('GDC_Val', 'count'),
        GDC_NHNN=('GDC_Val', 'sum')
    )

    for col in output_cols:
        if col not in df_grouped.columns:
            df_grouped[col] = 0.0

    return df_grouped[output_cols].fillna(0.0)
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

    # Lấy hệ số tham số NHNN cho Thủ quỹ & Lái xe
    h_thq_tong_hp = param_dict.get('NHNNTHUQUYTONGHP', 0.0)
    h_thq_hp      = param_dict.get('NHNNTHUQUYHP', 0.0)
    h_lx_hp       = param_dict.get('NHNNLAIXEHP', 0.0)

    h_thq_tong    = param_dict.get('NHNNTHUQUYTONG', 0.0)
    h_thq         = param_dict.get('NHNNTHUQUY', 0.0)
    h_lx          = param_dict.get('NHNNLAIXE', 0.0)

    h_hachtoan    = param_dict.get('HACHTOANNHNN', 0.0)

    # Nhận diện tên cột
    stt_col = _get_flexible_column(df1, ['stt', 'số thứ tự']) or df1.columns[0]
    coso_col = _get_flexible_column(df1, ['cơ sở', 'co so']) or df1.columns[1]
    loaidien_col = _get_flexible_column(df1, ['loại điện', 'loai dien'])
    khachhang_col = _get_flexible_column(df1, ['tên khách hàng', 'ten khach hang', 'khach hang'])
    
    # 1. Bỏ không xét tất cả các bản ghi có Tên khách hàng = NHNN chi nhánh khu vực 2
    if khachhang_col:
        df1 = df1[~df1[khachhang_col].astype(str).str.strip().str.lower().eq('nhnn chi nhánh khu vực 2')]

    # 2. Xử lý cột STT: Bỏ dấu chấm (.) và số đứng sau dấu chấm
    def clean_stt(val):
        s = str(val).strip()
        if '.' in s:
            s = s.split('.')[0]
        return s

    df1['STT_Processed'] = df1[stt_col].apply(clean_stt)
    df1['LoaiDien_Clean'] = df1[loaidien_col].astype(str).str.strip().str.lower() if loaidien_col else ''
    df1['Coso_Clean'] = df1[coso_col].astype(str).str.strip().str.lower() if coso_col else ''

    # Lấy tên cột Thủ quỹ và Lái xe thực tế trong file
    thug_col_name = _get_flexible_column(df1, ['thủ quỹ', 'thu quy'])
    laix_col_name = _get_flexible_column(df1, ['lái xe', 'lai xe'])

    # Bước A: Xử lý lọc trùng (Deduplicate) cấp độ CHUYẾN ĐI / BỘ ĐÔI giữ nguyên chuẩn logic cũ
    temp_rows = []
    for _, row in df1.iterrows():
        thq_raw = str(row.get(thug_col_name, '')).strip().lower() if thug_col_name else ''
        lx_raw = str(row.get(laix_col_name, '')).strip().lower() if laix_col_name else ''
        loai_dien = row['LoaiDien_Clean']
        is_tong = ('tổng' in loai_dien or 'tong' in loai_dien)

        temp_rows.append({
            'Original_Row': row,
            'Ngay_Key': row['Ngay_Key'],
            'STT_Processed': row['STT_Processed'],
            'ThuQuy_Trip': thq_raw,
            'LaiXe_Trip': lx_raw,
            'Is_Tong': is_tong
        })

    if not temp_rows:
        return pd.DataFrame(columns=output_cols)

    df_trip = pd.DataFrame(temp_rows)

    df_trip = df_trip.sort_values(
        by=['Ngay_Key', 'STT_Processed', 'ThuQuy_Trip', 'LaiXe_Trip', 'Is_Tong'],
        ascending=[True, True, True, True, False]
    )

    df_trip_dedup = df_trip.drop_duplicates(
        subset=['Ngay_Key', 'STT_Processed', 'ThuQuy_Trip', 'LaiXe_Trip'],
        keep='first'
    )

    # Bước B: Bóc tách nhân sự cho từng vị trí
    role_mapping = {
        'Thủ quỹ': ['thủ quỹ', 'thu quy'],
        'Lái xe': ['lái xe', 'lai xe'],
        'GDV NQC': ['gdv nqc', 'nqc', 'giao dịch viên nqc']
    }

    records = []

    for _, trip_row in df_trip_dedup.iterrows():
        row = trip_row['Original_Row']
        ngay_key = trip_row['Ngay_Key']
        stt_proc = trip_row['STT_Processed']
        coso_raw = row['Coso_Clean']
        loai_dien = row['LoaiDien_Clean']
        is_tong = ('tổng' in loai_dien or 'tong' in loai_dien)
        is_hp = ('hp' in coso_raw or 'hải phòng' in coso_raw or 'hai phong' in coso_raw)

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

                    # Gán chính xác GDC cho Thủ quỹ phân biệt rõ điện tổng và điện con
                    gdc_val = 0.0
                    if role_name == 'Thủ quỹ':
                        if is_hp:
                            gdc_val = h_thq_tong_hp if is_tong else h_thq_hp
                        else:
                            gdc_val = h_thq_tong if is_tong else h_thq
                    elif role_name == 'Lái xe':
                        gdc_val = h_lx_hp if is_hp else h_lx
                    elif role_name == 'GDV NQC':
                        gdc_val = h_hachtoan

                    records.append({
                        'MaSo': maso,
                        'ViTri': role_name,
                        'Ngay_Key': ngay_key,
                        'STT_Processed': stt_proc,
                        'GDC_Val': gdc_val
                    })

    if not records:
        return pd.DataFrame(columns=output_cols)

    df_res = pd.DataFrame(records)

    # Bước C: Tổng hợp kết quả số lượt và tổng GDC NHNN theo nhân sự và vị trí
    def compute_metrics(group):
        so_luot = len(group)
        gdc_sum = group['GDC_Val'].sum()
        return pd.Series({
            'SoLuot_NopLinhNHNN': so_luot,
            'GDC_NHNN': gdc_sum
        })

    df_grouped = df_res.groupby(['MaSo', 'ViTri'], as_index=False).apply(compute_metrics)

    for col in output_cols:
        if col not in df_grouped.columns:
            df_grouped[col] = 0.0

    return df_grouped[output_cols].fillna(0.0)
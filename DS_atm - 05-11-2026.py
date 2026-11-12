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

def _read_flexible_file(file_path):
    if not file_path or not os.path.exists(file_path):
        return pd.DataFrame()
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, header=0)
        else:
            df = pd.read_excel(file_path, header=0)
        df.columns = [str(c).strip() for c in df.columns]
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

# Hàm làm sạch ID máy chuẩn hóa, khắc phục triệt để lỗi thêm đuôi .0 do Pandas tự động ép kiểu float
def _clean_id_to_str(val):
    if pd.isna(val):
        return ""
    s_val = str(val).strip()
    if s_val.endswith('.0'):
        s_val = s_val[:-2]
    return s_val.upper()

def build_atm_employee_pipeline(shared_data_dir, start_str, end_str):
    f1_path = _find_file_by_keyword(shared_data_dir, "atm")
    df1 = _read_flexible_file(f1_path)
    
    start_dt = pd.to_datetime(start_str)
    end_dt = pd.to_datetime(end_str)

    records = []
    if not df1.empty:
        date_col = _get_flexible_column(df1, ['ngày', 'date', 'ngay hieu luc'])
        if date_col:
            df1['Ngay_Parsed'] = pd.to_datetime(df1[date_col], errors='coerce', dayfirst=True)
            df1 = df1[(df1['Ngay_Parsed'] >= start_dt) & (df1['Ngay_Parsed'] <= end_dt)]
        else:
            df1['Ngay_Parsed'] = pd.to_datetime(start_str)
        
        coso_col = _get_flexible_column(df1, ['cơ sở', 'co so']) or df1.columns[0]
        
        role_mapping = {
            'Thủ quỹ': ['thủ quỹ', 'thu quy'],
            'Lái xe': ['lái xe', 'lai xe'],
            'GDV ATM': ['gdv atm', 'gdv'],
            'Kiểm soát ATM': ['kiểm soát', 'kiem soat', 'ks atm'],
            'Chấm báo cáo': ['chấm báo cáo', 'cham bao cao', 'cbc']
        }

        for _, row in df1.iterrows():
            coso = str(row.get(coso_col, 'Cơ sở 1')).strip()
            ngay = row.get('Ngay_Parsed', None)
            
            for role_name, keywords in role_mapping.items():
                col_name = _get_flexible_column(df1, keywords)
                if col_name and col_name in df1.columns:
                    val = row.get(col_name)
                    if pd.notna(val) and str(val).strip() != '':
                        hoten_raw = str(val).strip()
                        if role_name == 'Lái xe' and ('không có' in hoten_raw.lower() or 'khong co' in hoten_raw.lower()):
                            continue
                        records.append({
                            'HoTen_Raw': hoten_raw,
                            'CoSo': coso,
                            'ViTri': role_name,
                            'Ngay': ngay
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

    df_valid_days = df_all.drop_duplicates(subset=['MaSo', 'HoTen', 'CoSo', 'ViTri', 'Ngay'])
    df_ngay_cong = df_valid_days.groupby(['MaSo', 'HoTen', 'CoSo', 'ViTri'], as_index=False).size()
    df_ngay_cong.rename(columns={'size': 'NgayCong'}, inplace=True)

    if df_ngay_cong.empty:
        df_ngay_cong = df_all.groupby(['MaSo', 'HoTen', 'CoSo', 'ViTri'], as_index=False).size()
        df_ngay_cong.rename(columns={'size': 'NgayCong'}, inplace=True)
        df_ngay_cong['NgayCong'] = 1.0

    return df_ngay_cong[['MaSo', 'HoTen', 'CoSo', 'ViTri', 'NgayCong']].copy()


def process_atm_kpi_matrix(shared_data_dir, param_dir, start_str, end_str):
    f1_path = _find_file_by_keyword(shared_data_dir, "atm")
    f2_path = _find_file_by_keyword(param_dir, "thamso_atm") or _find_file_by_keyword(shared_data_dir, "thamso_atm")
    
    df1 = _read_flexible_file(f1_path)
    df2 = _read_flexible_file(f2_path)
    
    start_dt = pd.to_datetime(start_str)
    end_dt = pd.to_datetime(end_str)

    output_cols = [
        'MaSo', 'ViTri', 
        'SoMay_Tiep', 'GDC_Tiep', 
        'SoMay_Sua', 'GDC_Sua', 
        'SoMay_BD', 'GDC_BD'
    ]

    if df1.empty:
        return pd.DataFrame(columns=output_cols)

    date_col = _get_flexible_column(df1, ['ngày', 'date', 'ngay hieu luc'])
    if date_col:
        df1['Ngay_Parsed'] = pd.to_datetime(df1[date_col], errors='coerce', dayfirst=True)
        df1 = df1[(df1['Ngay_Parsed'] >= start_dt) & (df1['Ngay_Parsed'] <= end_dt)]
    else:
        df1['Ngay_Parsed'] = pd.to_datetime(start_str)

    # 1. Đọc và chuẩn hóa file tham số (df2) với hàm làm sạch ID chuẩn
    param_dict = {}
    if not df2.empty:
        id_may_col_f2 = _get_flexible_column(df2, ['id máy', 'id may', 'ma may'])
        if id_may_col_f2:
            for _, r in df2.iterrows():
                m_id = _clean_id_to_str(r[id_may_col_f2])
                if not m_id: 
                    continue
                kc_val = _safe_float(r.get(_get_flexible_column(df2, ['KC từ máy đến điểm TQTT', 'kc']), 0))
                param_dict[m_id] = {
                    'kc': kc_val,
                    'gdc_bd_ks': _safe_float(r.get(_get_flexible_column(df2, ['gdc bao duong ks']), 0)),
                    'gdc_bd_tq': _safe_float(r.get(_get_flexible_column(df2, ['gdc baoduong tq']), 0)),
                    'gdc_bd_lx': _safe_float(r.get(_get_flexible_column(df2, ['gdc baoduong lx']), 0)),
                    'gdc_sua_tq': _safe_float(r.get(_get_flexible_column(df2, ['gdc sua chua tq']), 0)),
                    'gdc_sua_lx': _safe_float(r.get(_get_flexible_column(df2, ['gdc sua chua lx']), 0)),
                    'gdc_sua_ks': _safe_float(r.get(_get_flexible_column(df2, ['gdc sua chua ks']), 0)),
                    'gdc_thuong_lx': _safe_float(r.get(_get_flexible_column(df2, ['gdc thuong lx']), 0)),
                    'gdc_thuong_tq': _safe_float(r.get(_get_flexible_column(df2, ['gdc thuong tq']), 0)),
                    'gdc_thuong_ks': _safe_float(r.get(_get_flexible_column(df2, ['gdc thuong ks']), 0)),
                    'gdc_khac_lx': _safe_float(r.get(_get_flexible_column(df2, ['gdc may khac lx']), 0)),
                    'gdc_khac_tq': _safe_float(r.get(_get_flexible_column(df2, ['gdc may khac tq']), 0)),
                    'gdc_khac_ks': _safe_float(r.get(_get_flexible_column(df2, ['gdc may khac ks']), 0)),
                    'gdc_may_xa': _safe_float(r.get(_get_flexible_column(df2, ['gdc may xa']), 0)),
                    'gdc_gdv_atm': _safe_float(r.get(_get_flexible_column(df2, ['gdc gdv atm']), 0)),
                    'gdc_cbc': _safe_float(r.get(_get_flexible_column(df2, ['gdc cbc']), 0))
                }

    id_may_col_f1 = _get_flexible_column(df1, ['id máy', 'id may', 'ma may']) or df1.columns[0]
    tuyen_col = _get_flexible_column(df1, ['tuyến', 'tuyen'])
    phien_col = _get_flexible_column(df1, ['phiên', 'phien'])
    tiep_col = _get_flexible_column(df1, ['tiếp quỹ', 'tiep quy'])
    sua_col = _get_flexible_column(df1, ['sữa chữa', 'sửa chữa', 'sua chua'])
    bd_col = _get_flexible_column(df1, ['bảo dưỡng', 'bao duong'])

    # 2. Chuẩn hóa ngày, tuyến và ID máy tạm thời (đảm bảo hoàn toàn đồng nhất key)
    df1['_Ngay_Key'] = pd.to_datetime(df1['Ngay_Parsed']).dt.normalize().dt.strftime('%Y-%m-%d')
    df1['_Tuyen_Key'] = df1[tuyen_col].astype(str).str.strip().str.upper() if tuyen_col else 'TUYEN_DEFAULT'
    
    def extract_session_number(val):
        s_val = str(val).strip()
        if '1' in s_val:
            return 'PHIEN_1'
        elif '2' in s_val:
            return 'PHIEN_2'
        else:
            return s_val.upper()

    if phien_col:
        df1['_Phien_Key'] = df1[phien_col].apply(extract_session_number)
    else:
        df1['_Phien_Key'] = 'PHIEN_DEFAULT'

    # Áp dụng hàm làm sạch ID chuyên biệt để tuyệt đối không dính đuôi .0
    df1['_IdMay_Temp'] = df1[id_may_col_f1].apply(_clean_id_to_str)
    df1['KC_Val'] = df1['_IdMay_Temp'].apply(lambda x: param_dict.get(x, {}).get('kc', 0.0))

    # 3. Đánh dấu trạng thái Tiếp quỹ == 'X'
    df1['_Is_Tiep'] = False
    if tiep_col:
        df1['_Is_Tiep'] = df1[tiep_col].astype(str).str.strip().str.upper() == 'X'

    # 4. Tìm khoảng cách lớn nhất theo (Ngày + Tuyến + Phiên) trong các máy tiếp quỹ = 'X'
    dict_max_km = {}
    for _, row in df1.iterrows():
        if row['_Is_Tiep']:
            key_chuyen = f"{row['_Ngay_Key']}|{row['_Tuyen_Key']}|{row['_Phien_Key']}"
            km_may = row['KC_Val']
            if key_chuyen not in dict_max_km or km_may > dict_max_km[key_chuyen]:
                dict_max_km[key_chuyen] = km_may

    # 5. Gắn cờ Is_MayXa chính xác
    def check_is_may_xa(row):
        if not row['_Is_Tiep']:
            return False
        key_chuyen = f"{row['_Ngay_Key']}|{row['_Tuyen_Key']}|{row['_Phien_Key']}"
        max_km = dict_max_km.get(key_chuyen, -1)
        return True if (row['KC_Val'] > 0 and row['KC_Val'] == max_km) else False

    df1['Is_MayXa'] = df1.apply(check_is_may_xa, axis=1)
    df1['Is_MayXa'] = df1['Is_MayXa'] & (~df1.duplicated(subset=['_Ngay_Key', '_Tuyen_Key', '_Phien_Key', 'KC_Val'], keep='first'))

    # Xuất file debug kiểm tra phân loại máy
    df1['Xa_Khac'] = df1['Is_MayXa'].apply(lambda x: 'MÁY XA' if x else 'MÁY KHÁC')
    debug_file = os.path.join(shared_data_dir, "kiem_tra_phan_loai_may.xlsx")
    try:
        df1.to_excel(debug_file, index=False)
    except Exception as e:
        print(f"Không thể ghi file debug: {e}")

    df1.drop(columns=['_Ngay_Key', '_Tuyen_Key', '_Phien_Key', '_IdMay_Temp', '_Is_Tiep'], inplace=True, errors='ignore')

    records_kpi = []
    role_mapping = {
        'Thủ quỹ': ['thủ quỹ', 'thu quy'],
        'Lái xe': ['lái xe', 'lai xe'],
        'GDV ATM': ['gdv atm', 'gdv'],
        'Kiểm soát ATM': ['kiểm soát', 'kiem soat', 'ks atm'],
        'Chấm báo cáo': ['chấm báo cáo', 'cham bao cao', 'cbc']
    }

    for _, row in df1.iterrows():
        id_may = _clean_id_to_str(row.get(id_may_col_f1, ''))
        ngay_row = row.get('Ngay_Parsed', None)
        
        # Tra cứu thông số GDC từ dictionary chuẩn xác
        p_data = param_dict.get(id_may, {
            'kc': 0.0, 'gdc_bd_ks': 0.0, 'gdc_bd_tq': 0.0, 'gdc_bd_lx': 0.0,
            'gdc_sua_tq': 0.0, 'gdc_sua_lx': 0.0, 'gdc_sua_ks': 0.0,
            'gdc_thuong_lx': 0.0, 'gdc_thuong_tq': 0.0, 'gdc_thuong_ks': 0.0,
            'gdc_khac_lx': 0.0, 'gdc_khac_tq': 0.0, 'gdc_khac_ks': 0.0,
            'gdc_may_xa': 0.0, 'gdc_gdv_atm': 0.0, 'gdc_cbc': 0.0
        })

        is_may_xa = row.get('Is_MayXa', False)
        is_tiep = str(row.get(tiep_col, '')).strip().upper() == 'X' if tiep_col else False
        is_sua = str(row.get(sua_col, '')).strip().upper() == 'X' if sua_col else False
        is_bd = str(row.get(bd_col, '')).strip().upper() == 'X' if bd_col else False

        for role_name, keywords in role_mapping.items():
            col_name = _get_flexible_column(df1, keywords)
            if col_name and col_name in df1.columns:
                val = row.get(col_name)
                if pd.notna(val) and str(val).strip() != '':
                    raw_name = str(val).strip()
                    if role_name == 'Lái xe' and ('không có' in raw_name.lower() or 'khong co' in raw_name.lower()):
                        continue
                    
                    maso = raw_name.split('-')[0].strip().lower() if '-' in raw_name else raw_name.lower()

                    gdc_tiep_row = 0.0
                    gdc_sua_row = 0.0
                    gdc_bd_row = 0.0
                    
                    is_participating_tiep = is_tiep

                    if is_tiep:
                        if role_name == 'Thủ quỹ':
                            base_val = p_data['gdc_may_xa'] if is_may_xa else p_data['gdc_khac_tq']
                            gdc_tiep_row = base_val + p_data['gdc_thuong_tq']
                        elif role_name == 'Lái xe':
                            base_val = p_data['gdc_may_xa'] if is_may_xa else p_data['gdc_khac_lx']
                            gdc_tiep_row = base_val + p_data['gdc_thuong_lx']
                        elif role_name == 'Kiểm soát ATM':
                            base_val = p_data['gdc_may_xa'] if is_may_xa else p_data['gdc_khac_ks']
                            gdc_tiep_row = base_val + p_data['gdc_thuong_ks']
                        elif role_name == 'GDV ATM':
                            gdc_tiep_row = p_data['gdc_gdv_atm']
                        elif role_name == 'Chấm báo cáo':
                            gdc_tiep_row = p_data['gdc_cbc']

                    is_main_role = role_name in ['Thủ quỹ', 'Lái xe', 'Kiểm soát ATM']

                    if is_sua and is_main_role:
                        if role_name == 'Thủ quỹ': gdc_sua_row = p_data['gdc_sua_tq']
                        elif role_name == 'Lái xe': gdc_sua_row = p_data['gdc_sua_lx']
                        elif role_name == 'Kiểm soát ATM': gdc_sua_row = p_data['gdc_sua_ks']

                    if is_bd and is_main_role:
                        if role_name == 'Thủ quỹ': gdc_bd_row = p_data['gdc_bd_tq']
                        elif role_name == 'Lái xe': gdc_bd_row = p_data['gdc_bd_lx']
                        elif role_name == 'Kiểm soát ATM': gdc_bd_row = p_data['gdc_bd_ks']

                    records_kpi.append({
                        'MaSo': maso,
                        'ViTri': role_name,
                        'ID_May': id_may,
                        'Ngay': ngay_row,
                        'SoMay_Tiep': 1 if is_participating_tiep else 0,
                        'GDC_Tiep': gdc_tiep_row,
                        'SoMay_Sua': 1 if (is_sua and is_main_role) else 0,
                        'GDC_Sua': gdc_sua_row,
                        'SoMay_BD': 1 if (is_bd and is_main_role) else 0,
                        'GDC_BD': gdc_bd_row
                    })

    if not records_kpi:
        return pd.DataFrame(columns=output_cols)

    df_kpi = pd.DataFrame(records_kpi)
    df_kpi = df_kpi.drop_duplicates(subset=['MaSo', 'ViTri', 'ID_May', 'Ngay', 'SoMay_Tiep', 'GDC_Tiep'])
    df_grouped = df_kpi.groupby(['MaSo', 'ViTri'], as_index=False)[['SoMay_Tiep', 'GDC_Tiep', 'SoMay_Sua', 'GDC_Sua', 'SoMay_BD', 'GDC_BD']].sum()

    for col in output_cols:
        if col not in df_grouped.columns:
            df_grouped[col] = 0.0

    return df_grouped[output_cols].fillna(0.0)
import os
import pandas as pd
import numpy as np
import openpyxl

def _find_file_by_keyword(folder_path, keyword):
    if not os.path.exists(folder_path):
        return None
    for f in os.listdir(folder_path):
        if keyword.lower() in f.lower() and (f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.csv')):
            return os.path.join(folder_path, f)
    return None

def _read_flexible_file(file_path, header_row=0):
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
            if kw_clean in col_clean or col_clean in kw_clean:
                return col
    return None

def process_kiemdem_kpi_matrix(shared_data_dir, param_dir, start_str, end_str):
    """
    Xử lý kiểm đếm: Cố định cột 3 (Cán bộ), tách lấy phần trước dấu gạch ngang '-' làm MaSo chuẩn xác.
    """
    f1_path = _find_file_by_keyword(shared_data_dir, "kiemdem") or _find_file_by_keyword(shared_data_dir, "ds_kiemdem")
    
    output_cols = [
        'MaSo', 'MG_Lon', 'MG_Nho', 'NgoaiTe', 'KD_MG_Lon', 'KD_MG_Nho', 'KD_NT_MG_Lon', 'KD_NT_MG_Nho', 'GDC_KiemDem'
    ]

    if not f1_path or not os.path.exists(f1_path):
        print(f"⚠️ Không tìm thấy file kiểm đếm tại: {f1_path}")
        return pd.DataFrame(columns=output_cols)

    records = []
    try:
        wb = openpyxl.load_workbook(f1_path, data_only=True)
        sheet = wb.active

        col_canbo_idx = 3     # Cột C: Cán bộ
        col_mg_lon_idx = 33   # Cột MG Lớn
        col_mg_nho_idx = 34   # Cột MG Nhỏ
        col_ngoai_te_idx = 35 # Cột Ngoại tệ

        last_row = sheet.max_row
        current_canbo = ""

        for r in range(9, last_row + 1):
            cell_cb_val = sheet.cell(row=r, column=col_canbo_idx).value
            
            # Xử lý merge cell dọc (ffill) của cột Cán bộ
            if cell_cb_val is not None and str(cell_cb_val).strip() != "":
                current_canbo = str(cell_cb_val).strip()

            if not current_canbo or current_canbo.lower() == 'nan' or 'tổng số' in current_canbo.lower():
                continue

            # Tách lấy phần trước dấu gạch ngang '-' làm mã nhân viên độc lập
            raw_str = current_canbo
            if '-' in raw_str:
                maso = raw_str.split('-')[0].strip().lower()
            else:
                maso = raw_str.strip().lower()

            # Bỏ qua nếu mã không hợp lệ hoặc chứa từ khóa không phải nhân sự
            if not maso.isdigit() or 'pgd' in maso or 'nhóm' in maso or 'stt' in maso:
                continue

            mg_lon = _safe_float(sheet.cell(row=r, column=col_mg_lon_idx).value)
            mg_nho = _safe_float(sheet.cell(row=r, column=col_mg_nho_idx).value)
            ngoai_te = _safe_float(sheet.cell(row=r, column=col_ngoai_te_idx).value)

            if mg_lon == 0.0 and mg_nho == 0.0 and ngoai_te == 0.0:
                continue

            records.append({
                'MaSo': maso,
                'MG_Lon': round(mg_lon/1000,0),
                'MG_Nho': round(mg_nho/1000,0),
                'NgoaiTe': round(ngoai_te/1000,0)
            })

    except Exception as e:
        print(f"❌ Lỗi xử lý kiểm đếm: {e}")

    if not records:
        return pd.DataFrame(columns=output_cols)

    df_kd = pd.DataFrame(records)
    # Gom nhóm theo mã nhân viên
    df_kd = df_kd.groupby('MaSo', as_index=False).agg({
        'MG_Lon': 'sum',
        'MG_Nho': 'sum',
        'NgoaiTe': 'sum'
    })

    return df_kd
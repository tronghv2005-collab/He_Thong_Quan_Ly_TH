import pandas as pd
import numpy as np
from openpyxl import load_workbook
import os, re, uuid
from pathlib import Path

def no_accent_vietnamese(s):
    s = str(s).lower()
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[òóọỏõôồốộổổơờớợởỡ]', 'o', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[đ]', 'd', s)
    return re.sub(r'\s+', '', s)

def process_flat_header(file_path):
    df = pd.read_excel(file_path, header=[0,1])
    new_columns = []
    current_currency = "THONGTIN"
    
    for col in df.columns:
        c1 = str(col[0]).strip() if isinstance(col, tuple) else str(col).strip()
        c2 = str(col[1]).strip() if isinstance(col, tuple) and len(col) >= 2 else ""
        
        if "Unnamed" not in c1 and c1 != "" and "nan" not in c1.lower(): 
            current_currency = c1.upper().strip()
            
        clean1, clean2 = no_accent_vietnamese(c1), no_accent_vietnamese(c2)
        c_combined = (clean1 + "_" + clean2)
        
        if "sodien" in c_combined or "so_dien" in c_combined:
            new_columns.append(f"IGNORE_COL_{len(new_columns)}")
            continue
            
        if "ngayhieuluc" in c_combined or ("ngay" in c_combined and "luc" in c_combined):
            new_columns.append("NgayHieuLuc")
        elif "loainghiepvu" in c_combined or "nghiepvu" in c_combined:
            new_columns.append("LoaiNghiepVu")
        elif "thuquy" in c_combined:
            new_columns.append("ThuQuy")
        elif "laixe" in c_combined:
            new_columns.append("LaiXe")
        elif "tuyentiepquy" in c_combined or "tuyen" in c_combined:
            new_columns.append("TuyenTiepQuy")
        elif "khachhang" in c_combined or "tenkh" in c_combined:
            new_columns.append("KhachHang")
        elif "coso" in c_combined:
            new_columns.append("CoSo")
        elif "diadiem" in c_combined or "noplinh" in c_combined:
            new_columns.append("DiaDiemNopLinh")
        elif "thành tiền" in c2.lower() or "thanhtien" in clean1: 
            new_columns.append(f"{current_currency}_ThanhTien")
        else: 
            denom_clean = re.sub(r'[.,\s]', '', c2)
            if denom_clean.isdigit():
                new_columns.append(f"{current_currency}_{denom_clean}")
            else:
                new_columns.append(f"{current_currency}_{c2}" if "Unnamed" not in c2 else f"Col_{len(new_columns)}")
                    
    counts = {}
    final_cols = []
    for c in new_columns:
        counts[c] = counts.get(c, -1) + 1
        final_cols.append(f"{c}.{counts[c]}" if counts[c] > 0 else c)
    df.columns = final_cols
    return df

def clean_and_prepare_file1(file_path):
    if not os.path.exists(file_path): return pd.DataFrame()
    df = process_flat_header(file_path)
    str_cols = ['NgayHieuLuc', 'LoaiNghiepVu', 'ThuQuy', 'LaiXe', 'TuyenTiepQuy', 'KhachHang', 'CoSo']
    for col in str_cols:
        if col in df.columns: df[col] = df[col].astype(str).str.strip()
    if 'NgayHieuLuc' in df.columns: df['NgayHieuLuc'] = pd.to_datetime(df['NgayHieuLuc'], errors='coerce')
    
    if 'CoSo' in df.columns:
        df['CoSo'] = df['CoSo'].apply(lambda x: "Hub NVC" if "hub" in str(x).lower() and "nvc" in str(x).lower() else str(x).strip())
    if 'KhachHang' in df.columns:
        df['LoaiDoiTuong'] = df['KhachHang'].apply(lambda x: "Chi nhánh" if str(x)[:3] == "000" else "Phòng giao dịch")
    
    df['Is_Count_Turn'] = 1
    return _finalize_file_data(df, is_file2=False)

def clean_and_prepare_file2(file_path):
    if not os.path.exists(file_path): return pd.DataFrame()
    df = process_flat_header(file_path)
    
    info_fields = ['LoaiNghiepVu', 'DiaDiemNopLinh', 'TuyenTiepQuy', 'KhachHang', 'CoSo']
    for f in info_fields:
        if f in df.columns: df[f] = df[f].astype(str).str.strip()
    if 'NgayHieuLuc' in df.columns: df['NgayHieuLuc'] = pd.to_datetime(df['NgayHieuLuc'], errors='coerce')
    
    if 'CoSo' in df.columns:
        df['CoSo'] = df['CoSo'].apply(lambda x: "Hub NVC" if "hub" in str(x).lower() and "nvc" in str(x).lower() else str(x).strip())
        
    if 'LoaiNghiepVu' in df.columns:
        records_to_add = []
        
        for idx, row in df.iterrows():
            nv = str(row.get('LoaiNghiepVu', '')).strip()
            nv_clean = " ".join(nv.lower().split())
            dc = str(row.get('DiaDiemNopLinh', '')).strip()
            
            if nv_clean in ["thu tiền ( thu hộ kh(có tk)", "thu tiền (thu hộ khác)", "chi tiền(chi hộ khác)", "chi tiền (chi hộ kh (co tk)"]:
                continue
                
            # ĐỀ XUẤT CỦA BẠN: BẮT ĐÚNG DÒNG "CƠ SỞ KHÁC" VÀ TẠO 2 DÒNG RIÊNG BIỆT ĐỂ GÉP KHÓA
            if "cơ sở khác" in nv_clean or "co so khac" in nv_clean:
                
                # 1. Dòng Chi: Dùng loại nghiệp vụ gốc hoặc chuẩn hóa thành 'Chi tiền cơ sở khác'
                r_chi = row.copy()
                r_chi['LoaiNghiepVu'] = 'Chi tiền cơ sở khác'
                cs_val = str(r_chi.get('CoSo', '')).strip()
                if cs_val and cs_val.lower() != 'nan':
                    r_chi['CoSo'] = cs_val
                r_chi['Is_Count_Turn'] = 1
                records_to_add.append(r_chi)
                
                # 2. Dòng Thu: Đổi tên thành 'Thu tiền cơ sở khác', và gán cột CoSo = Tên khách hàng
                r_thu = row.copy()
                r_thu['LoaiNghiepVu'] = 'Thu tiền cơ sở khác'
                kh_val = str(r_thu.get('KhachHang', '')).strip()
                if kh_val and kh_val.lower() != 'nan':
                    r_thu['CoSo'] = kh_val  # Lấy tên khách hàng làm chuẩn cho cơ sở thu
                r_thu['Is_Count_Turn'] = 1
                records_to_add.append(r_thu)
                
            elif "thu tiền" in nv_clean and ("thu hộ" in nv_clean or "thu hộ kh" in nv_clean):
                row['LoaiNghiepVu'] = 'Thu hộ KH tại địa chỉ'
                row['Is_Count_Turn'] = 1
                records_to_add.append(row)
            elif "chi tiền" in nv_clean and ("chi hộ" in nv_clean or "chi hộ kh" in nv_clean) :
                row['LoaiNghiepVu'] = 'Chi hộ KH tại địa chỉ' 
                row['Is_Count_Turn'] = 1
                records_to_add.append(row) 
            elif "tctd nop tien" in nv_clean or "tctd nộp tiền" in nv.lower():
                r_root = row.copy()
                r_root['LoaiNghiepVu'] = 'Thu tiền (TCTD Nộp tiền)'
                r_root['Is_Count_Turn'] = 1
                records_to_add.append(r_root)
                
                r_child = row.copy()
                r_child['LoaiNghiepVu'] = 'TCTD nộp' if dc == 'Tại trung tâm' else 'Thu quỹ TCTD'
                r_child['Is_Count_Turn'] = 1
                r_child['Is_Child_Tag'] = '_CHILD_CON_'
                records_to_add.append(r_child)
            elif "tctd linh tien" in nv_clean or "tctd lĩnh tiền" in nv.lower():
                r_root = row.copy()
                r_root['LoaiNghiepVu'] = 'Chi tiền (TCTD Lĩnh tiền)'
                r_root['Is_Count_Turn'] = 1
                records_to_add.append(r_root)
                
                r_child = row.copy()
                r_child['LoaiNghiepVu'] = 'TCTD lĩnh' if dc == 'Tại trung tâm' else 'Tiếp quỹ'
                r_child['Is_Count_Turn'] = 1
                r_child['Is_Child_Tag'] = '_CHILD_CON_'
                records_to_add.append(r_child)
            else:
                row['Is_Count_Turn'] = 1
                records_to_add.append(row)
                
        df = pd.DataFrame(records_to_add)
            
    df['LoaiDoiTuong'] = "Toàn bộ"
        
    return _finalize_file_data(df, is_file2=True)

def _finalize_file_data(df, is_file2=False):
    if df.empty: return df
    str_cols = ['NgayHieuLuc', 'LoaiNghiepVu', 'KhachHang', 'CoSo']
    for col in str_cols:
        if col in df.columns: df[col] = df[col].astype(str).str.strip()
    if 'NgayHieuLuc' in df.columns: df['NgayHieuLuc'] = pd.to_datetime(df['NgayHieuLuc'], errors='coerce')
        
    if 'KhachHang' in df.columns and not is_file2:
        df['LoaiDoiTuong'] = df['KhachHang'].apply(lambda x: "Chi nhánh" if str(x)[:3] == "000" else "Phòng giao dịch")
    else:
        df['LoaiDoiTuong'] = "Toàn bộ"

    cac_tien_chinh = ['VND', 'USD', 'EUR', 'JPY']
    main_money_cols = {cc: [] for cc in cac_tien_chinh}
    all_numeric_money_cols = []
    
    for col in df.columns:
        if "_" in col and "IGNORE_COL" not in col and not any(k in col for k in ['NgayHieuLuc', 'LoaiNghiepVu', 'ThuQuy', 'LaiXe', 'TuyenTiepQuy', 'KhachHang', 'CoSo', 'Chuyen_Key', 'LoaiDoiTuong', 'DiaDiemNopLinh', 'Is_Count_Turn', 'Is_Child_Tag']):
            parts = col.split("_")
            prefix = str(parts[0]).upper().strip()
            suffix = str(parts[-1]).split(".")[0].lower().strip()
            
            if "thanhtien" not in suffix and suffix.isdigit():
                all_numeric_money_cols.append(col)
                matched_cc = next((cc for cc in cac_tien_chinh if cc in prefix), None)
                if matched_cc:
                    main_money_cols[matched_cc].append(col)

    for cc in cac_tien_chinh:
        m_cols = main_money_cols[cc]
        df[f'To_Sum_{cc}'] = df[m_cols].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if m_cols else 0
        t_col = next((c for c in df.columns if f"{cc}_ThanhTien" in c), None)
        df[f'{cc}_ThanhTien_Raw'] = pd.to_numeric(df.get(t_col, 0), errors='coerce').fillna(0) if t_col else 0

    flat_main_cols = [c for cols in main_money_cols.values() for c in cols]
    khac_cols = [c for c in all_numeric_money_cols if c not in flat_main_cols]
    df['To_Sum_Khac'] = df[khac_cols].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if khac_cols else 0

    date_str = df['NgayHieuLuc'].dt.strftime('%Y-%m-%d') if not df['NgayHieuLuc'].isna().all() else ''
    
    if is_file2:
        is_ho = df['LoaiNghiepVu'].isin(['Thu hộ KH tại địa chỉ', 'Chi hộ KH tại địa chỉ'])
        base_uuid = [str(uuid.uuid4()) for _ in range(len(df))]
        child_tag = df.get('Is_Child_Tag', '').fillna('')
        df['Chuyen_Key'] = np.where(is_ho, date_str + "_" + df['LoaiNghiepVu'] + "_" + df.get('TuyenTiepQuy', ''), base_uuid) + child_tag
    else:
        is_tn = df['LoaiNghiepVu'].isin(['Tiếp quỹ trong ngày', 'Thu gom trong ngày'])
        df['Chuyen_Key'] = np.where(is_tn, date_str + "_" + df.get('TuyenTiepQuy', ''), date_str + "_" + df['LoaiNghiepVu'] + "_" + df.get('ThuQuy', '') + "_" + df.get('LaiXe', ''))
    return df

def aggregate_center_summary(df_period):
    if df_period.empty:
        return pd.DataFrame()

    agg_d = {
        'Tinh_Luot': 'sum', 'Tinh_Phong': 'sum', 
        'To_Sum_VND': 'sum', 'VND_ThanhTien_Raw': 'sum',
        'To_Sum_USD': 'sum', 'USD_ThanhTien_Raw': 'sum',
        'To_Sum_EUR': 'sum', 'EUR_ThanhTien_Raw': 'sum',
        'To_Sum_JPY': 'sum', 'JPY_ThanhTien_Raw': 'sum',
        'To_Sum_Khac': 'sum', 'Tinh_KM_Tuyen': 'sum', 'Tinh_KM_Phong': 'sum'
    }

    def process_division(df_grouped):
        for cc in ['VND', 'USD', 'EUR', 'JPY', 'Khac']:
            df_grouped[f'Bo_{cc}'] = (df_grouped[f'To_Sum_{cc}'] / 1000).round(2)
        df_grouped['VND_ThanhTien'] = df_grouped['VND_ThanhTien_Raw']
        df_grouped['USD_ThanhTien'] = df_grouped['USD_ThanhTien_Raw']
        df_grouped['EUR_ThanhTien'] = df_grouped['EUR_ThanhTien_Raw']
        df_grouped['JPY_ThanhTien'] = df_grouped['JPY_ThanhTien_Raw']
        return df_grouped

    thu_filters = ['Thu gom cuối ngày', 'Thu gom trong ngày', 'CN Nộp tiền', 'Thu tiền (CN Nộp tiền)' ]
    chi_filters = ['Tiếp quỹ đầu ngày', 'Tiếp quỹ trong ngày', 'CN Lĩnh tiền', 'Chi tiền (CN Lĩnh tiền)']

    coso_list = df_period['CoSo'].dropna().unique()
    list_per_coso = []

    for cs in coso_list:
        df_cs = df_period[df_period['CoSo'] == cs]
        df_cs_nv = df_cs.groupby(['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'], as_index=False).agg(agg_d)
          
        df_cs_thu = df_cs[df_cs['LoaiNghiepVu'].isin(thu_filters)]
        if not df_cs_thu.empty:
            s_thu = df_cs_thu.agg(agg_d).to_dict()
            s_thu['LoaiNghiepVu'] = 'Tổng Thu'
            s_thu['CoSo'] = cs
            s_thu['LoaiDoiTuong'] = 'Toàn bộ'
            df_cs_nv = pd.concat([df_cs_nv, pd.DataFrame([s_thu])], ignore_index=True)

        df_cs_chi = df_cs[df_cs['LoaiNghiepVu'].isin(chi_filters)]
        if not df_cs_chi.empty:
            s_chi = df_cs_chi.agg(agg_d).to_dict()
            s_chi['LoaiNghiepVu'] = 'Tổng Chi'
            s_chi['CoSo'] = cs
            s_chi['LoaiDoiTuong'] = 'Toàn bộ'
            df_cs_nv = pd.concat([df_cs_nv, pd.DataFrame([s_chi])], ignore_index=True)

        list_per_coso.append(df_cs_nv)

    df_all_coso = pd.concat(list_per_coso, ignore_index=True) if list_per_coso else pd.DataFrame()
    if not df_all_coso.empty:
        df_all_coso = process_division(df_all_coso)

    center_rows = []
    df_center_thu = df_period[df_period['LoaiNghiepVu'].isin(thu_filters)]
    if not df_center_thu.empty:
        s_cthu = df_center_thu.agg(agg_d).to_dict()
        s_cthu['LoaiNghiepVu'] = 'Tổng Thu'
        s_cthu['CoSo'] = 'Toàn trung tâm'
        s_cthu['LoaiDoiTuong'] = 'Toàn bộ'
        center_rows.append(s_cthu)

    df_center_chi = df_period[df_period['LoaiNghiepVu'].isin(chi_filters)]
    if not df_center_chi.empty:
        s_cchi = df_center_chi.agg(agg_d).to_dict()
        s_cchi['LoaiNghiepVu'] = 'Tổng Chi'
        s_cchi['CoSo'] = 'Toàn trung tâm'
        s_cchi['LoaiDoiTuong'] = 'Toàn bộ'
        center_rows.append(s_cchi)

    df_center_final = pd.DataFrame(center_rows) if center_rows else pd.DataFrame(columns=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'])
    if not df_center_final.empty:
        df_center_final = process_division(df_center_final)

    return pd.concat([df_all_coso, df_center_final], ignore_index=True)

def aggregate_period(df_all, start_date, end_date, df_param_km):
    df_all['Clean_Date'] = pd.to_datetime(df_all['NgayHieuLuc']).dt.normalize()
    s_date = pd.to_datetime(start_date).normalize()
    e_date = pd.to_datetime(end_date).normalize()
    
    df_period = df_all[(df_all['Clean_Date'] >= s_date) & (df_all['Clean_Date'] <= e_date)].copy()
    if df_period.empty: return pd.DataFrame(columns=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'])
    
    df_param = df_param_km.copy()
    df_param.columns = [no_accent_vietnamese(c) for c in df_param.columns]
    p_cols = {c: 'KhachHang' for c in df_param.columns if c in ['khachhang', 'tenkhachhang']}
    p_cols.update({c: 'KM_PGD' for c in df_param.columns if 'kmpgd' in c})
    p_cols.update({c: 'KM_Tuyen' for c in df_param.columns if 'kmtuyen' in c})
    df_param = df_param.rename(columns=p_cols)[['KhachHang', 'KM_PGD', 'KM_Tuyen']].drop_duplicates(subset=['KhachHang']).copy()
    
    df_param['Join_Key'] = df_param['KhachHang'].astype(str).apply(no_accent_vietnamese)
    df_period['Join_Key'] = df_period['KhachHang'].astype(str).apply(no_accent_vietnamese)
    
    df_period = pd.merge(df_period, df_param, on='Join_Key', how='left')
    df_period['KM_PGD'] = pd.to_numeric(df_period['KM_PGD'], errors='coerce').fillna(0)
    df_period['KM_Tuyen'] = pd.to_numeric(df_period['KM_Tuyen'], errors='coerce').fillna(0)
    
    df_period['Is_First'] = ~df_period['Chuyen_Key'].duplicated()
    df_period['Tinh_Luot'] = np.where((df_period['Is_First']) & (df_period.get('Is_Count_Turn', 1) == 1), 1, 0)
    
    df_period['Tinh_KM_Tuyen'] = np.where((df_period['Is_First']) & (df_period['LoaiNghiepVu'].isin(['Tiếp quỹ đầu ngày', 'Thu gom cuối ngày', 'Thu hộ KH tại địa chỉ', 'Chi hộ KH tại địa chỉ'])), df_period['KM_Tuyen'], 0)
    df_period['Tinh_KM_Phong'] = df_period['KM_PGD']
    df_period['Tinh_Phong'] = 1

    agg_d = {'Tinh_Luot':'sum','Tinh_Phong':'sum','To_Sum_VND':'sum','VND_ThanhTien_Raw':'sum','To_Sum_USD':'sum','USD_ThanhTien_Raw':'sum','To_Sum_EUR':'sum','EUR_ThanhTien_Raw':'sum','To_Sum_JPY':'sum','JPY_ThanhTien_Raw':'sum','To_Sum_Khac':'sum','Tinh_KM_Tuyen':'sum','Tinh_KM_Phong':'sum'}
    
    def process_division(df_grouped):
        for cc in ['VND', 'USD', 'EUR', 'JPY', 'Khac']:
            df_grouped[f'Bo_{cc}'] = (df_grouped[f'To_Sum_{cc}'] / 1000).round(2)
        df_grouped['VND_ThanhTien'] = df_grouped['VND_ThanhTien_Raw']
        df_grouped['USD_ThanhTien'] = df_grouped['USD_ThanhTien_Raw']
        df_grouped['EUR_ThanhTien'] = df_grouped['EUR_ThanhTien_Raw']
        df_grouped['JPY_ThanhTien'] = df_grouped['JPY_ThanhTien_Raw']
        return df_grouped

    sd = df_period.groupby(['LoaiNghiepVu','CoSo','LoaiDoiTuong'], as_index=False).agg(agg_d)
    sd = process_division(sd)
    
    file1_nv = ['Tiếp quỹ đầu ngày', 'Thu gom cuối ngày', 'Tiếp quỹ trong ngày', 'Thu gom trong ngày']
    df_f1_records = df_period[df_period['LoaiNghiepVu'].isin(file1_nv)].copy()
    if not df_f1_records.empty:
        st = df_f1_records.groupby(['LoaiNghiepVu','CoSo'], as_index=False).agg(agg_d)
        st['LoaiDoiTuong'] = "Toàn bộ"
        st = process_division(st)
        sd = pd.concat([sd, st], ignore_index=True).drop_duplicates(subset=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'])
        
    df_center = aggregate_center_summary(df_period)
    final_combined = pd.concat([sd, df_center], ignore_index=True)
    return final_combined.drop_duplicates(subset=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'])

def format_summary_dataframe(df_prev, df_curr, all_nghiep_vu, all_coso):
    rows = []
    file1_nv = ['Tiếp quỹ đầu ngày', 'Thu gom cuối ngày', 'Tiếp quỹ trong ngày', 'Thu gom trong ngày']
    
    desired_nv_order = [
        'Tiếp quỹ đầu ngày', 'Tiếp quỹ trong ngày', 'Thu gom trong ngày', 'Thu gom cuối ngày',
        'Thu hộ KH tại địa chỉ', 'Chi hộ KH tại địa chỉ', 'Thu khách hàng khác', 'Chi khách hàng khác',
        'Thu tiền (TCTD Nộp tiền)', 'Thu quỹ TCTD', 'TCTD nộp', 'Chi tiền (TCTD Lĩnh tiền)', 'Tiếp quỹ', 'TCTD lĩnh',
        'Chi nhánh nộp tiền', 'Chi nhánh lĩnh', 'Chi tiền (chi nhánh lĩnh tiền)', 
        'Thu tiền cơ sở khác', 'Chi tiền cơ sở khác'
    ]
    
    combined_nv = [nv for nv in desired_nv_order if nv in all_nghiep_vu]
    for nv in all_nghiep_vu:
        if nv not in combined_nv and nv not in ['Tổng Thu', 'Tổng Chi']:
            combined_nv.append(nv)

    for nv in combined_nv:
        for cs in all_coso:
            if cs == "Toàn trung tâm": continue
            if nv in file1_nv:
                for dt in ['Toàn bộ', 'Chi nhánh', 'Phòng giao dịch']: 
                    rows.append({'LoaiNghiepVu': nv, 'CoSo': cs, 'LoaiDoiTuong': dt})
            else: 
                rows.append({'LoaiNghiepVu': nv, 'CoSo': cs, 'LoaiDoiTuong': 'Toàn bộ'})

    for cs in all_coso:
        if cs != "Toàn trung tâm":
            rows.append({'LoaiNghiepVu': 'Tổng Thu', 'CoSo': cs, 'LoaiDoiTuong': 'Toàn bộ'})
            rows.append({'LoaiNghiepVu': 'Tổng Chi', 'CoSo': cs, 'LoaiDoiTuong': 'Toàn bộ'})

    rows.append({'LoaiNghiepVu': 'Tổng Thu', 'CoSo': 'Toàn trung tâm', 'LoaiDoiTuong': 'Toàn bộ'})
    rows.append({'LoaiNghiepVu': 'Tổng Chi', 'CoSo': 'Toàn trung tâm', 'LoaiDoiTuong': 'Toàn bộ'})
            
    base_df = pd.DataFrame(rows).drop_duplicates(subset=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'])
    if df_prev.empty or 'LoaiNghiepVu' not in df_prev.columns: df_prev = pd.DataFrame(columns=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'])
    if df_curr.empty or 'LoaiNghiepVu' not in df_curr.columns: df_curr = pd.DataFrame(columns=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'])
        
    f_df = pd.merge(base_df, df_prev, on=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'], how='left').fillna(0)
    f_df = pd.merge(f_df, df_curr, on=['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong'], how='left', suffixes=('_Kỳ Trước', '_Kỳ Báo Cáo')).fillna(0)
    
    metrics = ['Tinh_Luot', 'Tinh_Phong', 'Bo_VND', 'VND_ThanhTien', 'Bo_USD', 'USD_ThanhTien', 'Bo_EUR', 'EUR_ThanhTien', 'Bo_JPY', 'JPY_ThanhTien', 'Bo_Khac', 'Tinh_KM_Tuyen', 'Tinh_KM_Phong']
    final_cols = ['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong']
    for m in metrics: final_cols.extend([f"{m}_Kỳ Trước", f"{m}_Kỳ Báo Cáo"])
    f_df = f_df.reindex(columns=final_cols, fill_value=0.0)

    check_cols = [c for c in f_df.columns if c not in ['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong']]
    f_df = f_df[(f_df[check_cols] != 0).any(axis=1)].copy()
    
    return f_df

def write_to_template(final_table, template_path, output_path):
    if not os.path.exists(template_path): final_table.to_excel(output_path, index=False); return
    wb = load_workbook(template_path); ws = wb.active
    data_rows = final_table.to_dict(orient='records')
    for r in range(4, ws.max_row + 1):
        nv = str(ws.cell(row=r, column=1).value).strip()
        cs = str(ws.cell(row=r, column=2).value).strip()
        dt = str(ws.cell(row=r, column=3).value).strip()
        m = next((d for d in data_rows if str(d.get('LoaiNghiepVu', '')).strip() == nv and str(d.get('CoSo', '')).strip() == cs and str(d.get('LoaiDoiTuong', '')).strip() == dt), None)
        if m:
            v_keys = [k for k in m.keys() if k not in ['LoaiNghiepVu', 'CoSo', 'LoaiDoiTuong']]
            for idx, k in enumerate(v_keys): ws.cell(row=r, column=4 + idx).value = m[k]
    wb.save(output_path); wb.close()
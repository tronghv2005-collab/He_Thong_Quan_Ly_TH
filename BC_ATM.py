import pandas as pd
import numpy as np
import os
import unicodedata

def process_atm_summary_report(path_nhatky, path_param, start_curr, end_curr, start_prev, end_prev, output_excel_path, **kwargs):
    # =========================================================================
    # 1. ĐỌC VÀ CHUẨN HÓA FILE THAM SỐ (FILE 1)
    # =========================================================================
    df_param = pd.read_excel(path_param)
    df_param.columns = [str(c).strip() for c in df_param.columns]
    
    id_param_col = next((c for c in df_param.columns if 'id' in c.lower() or 'máy' in c.lower() or 'may' in c.lower()), df_param.columns[0])
    kc_may_col = next((c for c in df_param.columns if 'kc từ máy' in c.lower() or ('km' in c.lower() and 'tqtt' in c.lower())), df_param.columns[1])
    tuyen_param_col = next((c for c in df_param.columns if 'tuyến' in c.lower() and 'km' not in c.lower()), df_param.columns[2])
    coso_param_col = next((c for c in df_param.columns if 'cơ sở' in c.lower() or 'co so' in c.lower()), df_param.columns[3])
    km_tuyen_param_col = next((c for c in df_param.columns if 'tuyến' in c.lower() and 'km' in c.lower()), df_param.columns[4])
    
    df_param_clean = pd.DataFrame()
    df_param_clean['ID_MAY_CHUAN'] = df_param[id_param_col].astype(str).str.strip().str.upper()
    df_param_clean['KC_MAY_CHUAN'] = pd.to_numeric(df_param[kc_may_col], errors='coerce').fillna(0)
    df_param_clean['TUYEN_CHUAN'] = df_param[tuyen_param_col].astype(str).str.strip()
    df_param_clean['CO_SO_CHUAN'] = df_param[coso_param_col].astype(str).str.strip()
    df_param_clean['KM_TUYEN_CHUAN'] = pd.to_numeric(df_param[km_tuyen_param_col], errors='coerce').fillna(0)
    
    df_param_clean.drop_duplicates(subset=['ID_MAY_CHUAN'], keep='first', inplace=True)

    # =========================================================================
    # 2. ĐỌC VÀ CHUẨN HÓA FILE NHẬT KÝ GIAO DỊCH (FILE 2)
    # =========================================================================
    df_nhatky = pd.read_excel(path_nhatky, header=0)
    df_nhatky.columns = [str(c).strip() for c in df_nhatky.columns]
    
    def find_col_by_keywords(df, keywords):
        for col in df.columns:
            col_lower = str(col).lower()
            for kw in keywords:
                if kw in col_lower:
                    return col
        return None

    diengiai_col = find_col_by_keywords(df_nhatky, ['diễn giải', 'dien giai', 'nội dung', 'noi dung'])
    notco_col = find_col_by_keywords(df_nhatky, ['nợ/có', 'no/co', 'no co', 'nợ có', 'bút toán', 'loại gd'])
    sotien_col = find_col_by_keywords(df_nhatky, ['số tiền', 'so tien', 'thành tiền', 'thanh tien'])
    thoigian_col = find_col_by_keywords(df_nhatky, ['thời gian', 'thoi gian', 'giờ', 'gio'])
    ngay_col = find_col_by_keywords(df_nhatky, ['ngày', 'ngay', 'date'])

    if not notco_col: raise ValueError("Không tìm thấy cột 'Nợ/Có'.")
    if not sotien_col: raise ValueError("Không tìm thấy cột 'Số tiền'.")
    if not diengiai_col: raise ValueError("Không tìm thấy cột 'Diễn giải'.")

    df_nhatky_clean = pd.DataFrame()
    raw_ngay = df_nhatky[ngay_col] if ngay_col else df_nhatky.iloc[:, 0]
    df_nhatky_clean['NGAY_GD'] = pd.to_datetime(raw_ngay, errors='coerce')

    df_nhatky_clean['THOI_GIAN'] = df_nhatky[thoigian_col] if thoigian_col else df_nhatky.iloc[:, 1]
    df_nhatky_clean['NOT_CO'] = df_nhatky[notco_col].astype(str).str.strip().str.upper()
    df_nhatky_clean['SO_TIEN'] = pd.to_numeric(df_nhatky[sotien_col].astype(str).str.replace(',', '').str.replace('.', '', regex=False), errors='coerce').fillna(0)
    df_nhatky_clean['ID_MAY_CALC'] = df_nhatky[diengiai_col].astype(str).apply(lambda s: s[-8:].strip().upper() if len(s) >= 8 else s.strip().upper())

    # =========================================================================
    # 3. MERGE VÀ PHÂN CHIA KỲ BÁO CÁO (XUẤT FILE DEBUG)
    # =========================================================================
    merged_df = pd.merge(df_nhatky_clean, df_param_clean, left_on='ID_MAY_CALC', right_on='ID_MAY_CHUAN', how='left')
    
    merged_df['CO_SO_CHUAN'] = merged_df['CO_SO_CHUAN'].fillna('CHUA PHAN LOAI')
    merged_df['TUYEN_CHUAN'] = merged_df['TUYEN_CHUAN'].fillna('CHUA PHAN LOAI')

    row_dates = merged_df['NGAY_GD'].dt.date

    def assign_period(d):
        if pd.isna(d):
            return 'Khác'
        if start_curr <= d <= end_curr:
            return 'Kỳ này'
        elif start_prev <= d <= end_prev:
            return 'Kỳ trước'
        return 'Khác'

    merged_df['Ky_BaoCao'] = row_dates.apply(assign_period)
    
    print(f"-> [THỐNG KÊ PHÂN BỔ KỲ] Kỳ này: {len(merged_df[merged_df['Ky_BaoCao'] == 'Kỳ này'])} dòng | Kỳ trước: {len(merged_df[merged_df['Ky_BaoCao'] == 'Kỳ trước'])} dòng | Khác: {len(merged_df[merged_df['Ky_BaoCao'] == 'Khác'])} dòng")

    try:
        merged_df[merged_df['Ky_BaoCao'] == 'Kỳ trước'].to_excel("DuLieu_ChiTiet_KyTruoc.xlsx", index=False)
        merged_df[merged_df['Ky_BaoCao'] == 'Kỳ này'].to_excel("DuLieu_ChiTiet_KyNay.xlsx", index=False)
    except Exception as e:
        print(f"-> Không thể xuất file debug: {e}")

    # Phân loại ATM / CRM
    merged_df['Loai_May'] = merged_df['ID_MAY_CALC'].apply(lambda x: 'ATM' if len(x) >= 5 and x[3:5] == '98' else ('CRM' if len(x) >= 5 and x[3:5] == '99' else 'Khác'))
    
    def parse_hour(t):
        try: return pd.to_datetime(str(t).strip()).hour
        except:
            s = str(t).strip()
            if ':' in s:
                try: return int(s.split(':')[0])
                except: pass
            return 9
    merged_df['Phiên'] = merged_df['THOI_GIAN'].apply(lambda t: 'Sáng' if parse_hour(t) <= 12 else 'Chiều')

    # =========================================================================
    # 4. GOM NHÓM (KM MÁY DÙNG SUM() ĐỂ CỘNG DỒN TOÀN BỘ)
    # =========================================================================
    coso_thuc_te = sorted([c for c in merged_df['CO_SO_CHUAN'].unique() if c != 'CHUA PHAN LOAI'])
    if 'CHUA PHAN LOAI' in merged_df['CO_SO_CHUAN'].values:
        coso_thuc_te.append('CHUA PHAN LOAI')
        
    coso_display = coso_thuc_te + ["Toàn trung tâm"]
    
    result_rows = []
    for cs in coso_display:
        df_cs = merged_df.copy() if cs == "Toàn trung tâm" else merged_df[merged_df['CO_SO_CHUAN'] == cs].copy()
            
        for target_type in ['Tổng', 'ATM', 'CRM']:
            row_data = {'Cơ sở': cs, 'Phân loại': target_type}
            
            if target_type == 'Tổng':
                df_sub = df_cs.copy()
            else:
                df_sub = df_cs[df_cs['Loai_May'] == target_type].copy()
                
            for ky in ['Kỳ trước', 'Kỳ này']:
                suffix = ' kỳ trước' if ky == 'Kỳ trước' else ' Kỳ này'
                
                df_ky = df_sub[df_sub['Ky_BaoCao'] == ky]
                
                if len(df_ky) == 0:
                    for k in ['Số tiếp máy', 'Số tiền TQ', 'Số tiền HQ', 'KM máy', 'KM tuyến']: 
                        row_data[f'{k}{suffix}'] = 0
                    continue
                
                df_dr = df_ky[df_ky['NOT_CO'] == 'DR']
                row_data[f'Số tiếp máy{suffix}'] = len(df_dr)
                row_data[f'Số tiền TQ{suffix}'] = df_dr['SO_TIEN'].sum()
                row_data[f'Số tiền HQ{suffix}'] = df_ky[df_ky['NOT_CO'] == 'CR']['SO_TIEN'].sum()
                
                # [ĐÃ SỬA]: Dùng sum() để cộng dồn toàn bộ số KM máy phát sinh trong kỳ
                row_data[f'KM máy{suffix}'] = df_ky['KC_MAY_CHUAN'].sum() if len(df_ky) > 0 else 0
                
                # KM tuyến: Chuẩn hóa mỗi Ngày + Tuyến + Phiên chỉ tính tối đa 1 lần
                df_phien_tuyen = df_ky.drop_duplicates(subset=['NGAY_GD', 'TUYEN_CHUAN', 'Phiên'])
                row_data[f'KM tuyến{suffix}'] = df_phien_tuyen['KM_TUYEN_CHUAN'].sum()

            result_rows.append(row_data)

    # =========================================================================
    # 5. XUẤT BÁO CÁO 17 CỘT CHUẨN & LƯU FILE
    # =========================================================================
    final_report = pd.DataFrame(result_rows)
    for ky_prefix in ['kỳ trước', 'Kỳ này']:
        sfx = f' {ky_prefix}'
        final_report[f'BQ tiếp{sfx}'] = np.where(final_report[f'Số tiếp máy{sfx}'] > 0, final_report[f'Số tiền TQ{sfx}'] / final_report[f'Số tiếp máy{sfx}'], 0)
        final_report[f'BQ HQ{sfx}'] = np.where(final_report[f'Số tiếp máy{sfx}'] > 0, final_report[f'Số tiền HQ{sfx}'] / final_report[f'Số tiếp máy{sfx}'], 0)
        final_report[f'Tỷ lệ HQ{sfx}'] = np.where(final_report[f'Số tiền TQ{sfx}'] > 0, (final_report[f'Số tiền HQ{sfx}'] / final_report[f'Số tiền TQ{sfx}']) * 100, 0)

    ordered_cols = [
        'Cơ sở', 'Số tiếp máy kỳ trước', 'Số tiếp máy Kỳ này', 'Số tiền TQ kỳ trước', 'Số tiền TQ Kỳ này',
        'Số tiền HQ kỳ trước', 'Số tiền HQ Kỳ này', 'BQ tiếp kỳ trước', 'BQ tiếp Kỳ này', 'BQ HQ kỳ trước',
        'BQ HQ Kỳ này', 'Tỷ lệ HQ kỳ trước', 'Tỷ lệ HQ Kỳ này', 'KM máy kỳ trước', 'KM máy Kỳ này',
        'KM tuyến kỳ trước', 'KM tuyến Kỳ này'
    ]
    
    final_report['Cơ sở'] = np.where(final_report['Phân loại'] == 'Tổng', final_report['Cơ sở'], final_report['Cơ sở'].apply(lambda x: f"   + {x} ") + final_report['Phân loại'])
    final_report = final_report.reindex(columns=ordered_cols)

    try: 
        final_report.to_excel(output_excel_path, index=False)
    except PermissionError: 
        final_report.to_excel(f"{os.path.splitext(output_excel_path)}_duphong.xlsx", index=False)
    
    return final_report
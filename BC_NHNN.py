import pandas as pd
import numpy as np
import os, re
from pathlib import Path
from openpyxl.utils import get_column_letter

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

def clean_and_prepare_nhnn(file_path):
    if not os.path.exists(file_path): return pd.DataFrame()
    
    # 1. Đọc thô file Excel (Bỏ qua dòng 1 trắng, lấy dòng 2 và 3 làm tiêu đề gộp)
    df = pd.read_excel(file_path, header=[1, 2])
    
    raw_columns = list(df.columns)
    new_columns = []
    current_currency = "THONGTIN"
    
    # 2. Vòng lặp chuẩn hóa tên cột
    for col in raw_columns:
        c1 = str(col[0]).strip() if isinstance(col, tuple) else str(col).strip()
        c2 = str(col[1]).strip() if isinstance(col, tuple) and len(col) >= 2 else ""
        
        if "Unnamed" not in c1 and c1 != "" and "nan" not in c1.lower(): 
            current_currency = c1.upper().strip()
            
        clean1, clean2 = no_accent_vietnamese(c1), no_accent_vietnamese(c2)
        c_combined = (clean1 + "_" + clean2)
        
        if "sodien" in c_combined or "so_dien" in c_combined:
            new_columns.append("SoDien")
        elif "ngayhieuluc" in c_combined or ("ngay" in c_combined and "luc" in c_combined) or "ngaygiaodich" in c_combined:
            new_columns.append("NgayHieuLoc")
        elif "tenkhachhang" in c_combined or "khachhang" in c_combined or "tenkh" in c_combined:
            new_columns.append("TenKhachHang")
        elif "loaidien" in c_combined or "dien" in c_combined:
            new_columns.append("LoaiDien")
        elif "coso" in c_combined or "phongban" in c_combined or "donvi" in c_combined:
            new_columns.append("CoSo")
        else:
            denom_clean = re.sub(r'[.,\s]', '', c2)
            if "thanhtien" in clean2 or "thanhtien" in c_combined or "thanh_tien" in clean2:
                if "VND_R" in current_currency or "VNDR" in current_currency:
                    new_columns.append("VNDR_THANHTIEN")
                elif "VND" in current_currency:
                    new_columns.append("VND_THANHTIEN")
                else:
                    new_columns.append(f"{current_currency}_THANHTIEN")
            elif denom_clean.isdigit():
                if "VND_R" in current_currency or "VNDR" in current_currency:
                    new_columns.append(f"VNDR_{denom_clean}")
                elif "VND" in current_currency:
                    new_columns.append(f"VND_{denom_clean}")
                else:
                    new_columns.append(f"{current_currency}_{denom_clean}")
            else:
                new_columns.append(f"{current_currency}_{c2}" if "Unnamed" not in c2 else f"Col_{len(new_columns)}")
                    
    # Xử lý tránh trùng tên cột và ép phẳng tiêu đề 2 tầng
    counts = {}
    final_cols = []
    for c in new_columns:
        counts[c] = counts.get(c, -1) + 1
        final_cols.append(f"{c}.{counts[c]}" if counts[c] > 0 else c)
    df.columns = final_cols

    # 🎯 ĐỒNG BỘ: Tạo cột 'Loại nghiệp vụ' chính xác từ cột thứ 11 (index 10)
    if df.shape[1] > 10:
        df['Loại nghiệp vụ'] = df.iloc[:, 10].astype(str).str.strip()
    else:
        df['Loại nghiệp vụ'] = "Nộp/Lĩnh"

    # Làm sạch dữ liệu chuỗi cơ bản
    str_cols = ['SoDien', 'NgayHieuLoc', 'CoSo', 'LoaiDien', 'TenKhachHang', 'Loại nghiệp vụ']
    for col in str_cols:
        if col in df.columns: 
            df[col] = df[col].astype(str).str.strip()
        
    if 'NgayHieuLoc' in df.columns: 
        df['NgayHieuLoc'] = pd.to_datetime(df['NgayHieuLoc'], errors='coerce')
        
    if 'CoSo' in df.columns:
        def standardize_coso(val):
            v_low = str(val).lower()
            if "hub" in v_low and "nvc" in v_low:
                return "HUB NVC"
            elif "1" in v_low or "co so 1" in v_low or "cơ sở 1" in v_low:
                return "Cơ sở 1"
            elif "2" in v_low or "co so 2" in v_low or "cơ sở 2" in v_low:
                return "Cơ sở 2"
            elif "hd" in v_low or "hđ" in v_low or "hai duong" in v_low:
                return "CS HĐ"
            elif "hp" in v_low or "hai phong" in v_low:
                return "CS HP"
            return str(val).strip()
        df['CoSo'] = df['CoSo'].apply(standardize_coso)

    if 'TenKhachHang' in df.columns:
        df = df[df['TenKhachHang'].str.lower() != "nhnn chi nhánh khu vực 2"].copy()

    if 'LoaiDien' in df.columns:
        def standardize_loaidien(val):
            v_low = str(val).lower()
            if "tổng" in v_low:
                return "Điện tổng"
            elif "con" in v_low:
                return "Điện con"
            return str(val).strip() if str(val).strip() != "" and str(val).strip().lower() != "nan" else "Điện tổng"
        df['Loại điện'] = df['LoaiDien'].apply(standardize_loaidien)
    else:
        df['Loại điện'] = "Điện tổng"

    # Lọc tính toán số lượng bó và tiền của ĐTC và KĐTC
    dtc_cols = [c for c in df.columns if c.startswith("VND_") and c != "VND_THANHTIEN" and not "THANHTIEN" in c.upper()]
    dtc_tien_col = next((c for c in df.columns if c == "VND_THANHTIEN" or "VND_THANHTIEN" in c.upper() or "VND_THÀNHTIỀN" in c.upper()), None)

    kdtc_cols = [c for c in df.columns if c.startswith("VNDR_") and c != "VNDR_THANHTIEN" and not "THANHTIEN" in c.upper()]
    kdtc_tien_col = next((c for c in df.columns if c == "VNDR_THANHTIEN" or "VNDR_THANHTIEN" in c.upper() or "VNDR_THÀNHTIỀN" in c.upper()), None)

    tong_to_dtc = df[dtc_cols].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if dtc_cols else 0
    sobo_dtc = np.round(tong_to_dtc / 1000)
    sotien_dtc = pd.to_numeric(df[dtc_tien_col], errors='coerce').fillna(0) if dtc_tien_col else 0

    tong_to_kdtc = df[kdtc_cols].fillna(0).apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1) if kdtc_cols else 0
    sobo_kdtc = np.round(tong_to_kdtc / 1000)
    sotien_kdtc = pd.to_numeric(df[kdtc_tien_col], errors='coerce').fillna(0) if kdtc_tien_col else 0

    df['SoLuot'] = 1
    df['SoBo_ĐTC'] = sobo_dtc
    df['SoTien_ĐTC'] = sotien_dtc
    df['SoBo_KĐTC'] = sobo_kdtc
    df['SoTien_KĐTC'] = sotien_kdtc

    mask_dien_con = df['Loại điện'].str.lower().str.contains("điện con", na=False)
    df.loc[mask_dien_con, ['SoBo_ĐTC', 'SoTien_ĐTC', 'SoBo_KĐTC', 'SoTien_KĐTC']] = 0

    return df
def aggregate_nhnn_period(df_all, start_date, end_date):
    if df_all.empty: return pd.DataFrame()
    
    # Tạo bản sao sâu (deep copy) để tránh lỗi SettingWithCopyWarning
    df_working = df_all.copy()
    
    # 🎯 ĐỒNG BỘ: Rào chắn bảo vệ tên cột 'Loại nghiệp vụ'
    if 'Loại nghiệp vụ' not in df_working.columns:
        if df_working.shape[1] > 10:
            df_working['Loại nghiệp vụ'] = df_working.iloc[:, 10].astype(str).str.strip()
        else:
            df_working['Loại nghiệp vụ'] = "Nộp/Lĩnh"

    df_working['Clean_Date'] = pd.to_datetime(df_working['NgayHieuLoc'], errors='coerce').dt.normalize()
    s_date = pd.to_datetime(start_date).normalize()
    e_date = pd.to_datetime(end_date).normalize()
    
    df_period = df_working[(df_working['Clean_Date'] >= s_date) & (df_working['Clean_Date'] <= e_date)].copy()
    if df_period.empty: return pd.DataFrame()

    agg_dict = {
        'SoLuot': 'sum',
        'SoBo_ĐTC': 'sum',
        'SoTien_ĐTC': 'sum',
        'SoBo_KĐTC': 'sum',
        'SoTien_KĐTC': 'sum'
    }

    # 🎯 ĐỒNG BỘ: Groupby sử dụng cột 'Loại nghiệp vụ' có dấu
    grouped = df_period.groupby(['Loại nghiệp vụ', 'CoSo', 'Loại điện'], as_index=False).agg(agg_dict)
    return grouped

def process_ds_nop_linh_nhnn_pipeline(folder_path, start_date_kc, end_date_kc, start_date_kp, end_date_kp, output_excel_path):
    file_path = os.path.join(folder_path, "Ds_NopLinhNHNN.xlsx")
    
    if not os.path.exists(file_path):
        print(f"Không tìm thấy file Ds_NopLinhNHNN.xlsx tại thư mục {folder_path}.")
        return False
        
    df_temp = clean_and_prepare_nhnn(file_path)
    if df_temp.empty:
        print("File Ds_NopLinhNHNN.xlsx rỗng hoặc không có dữ liệu hợp lệ.")
        return False
        
    df_baocao = aggregate_nhnn_period(df_temp, start_date_kc, end_date_kc)
    df_kytrước = aggregate_nhnn_period(df_temp, start_date_kp, end_date_kp)
    
    # Kiểm tra rỗng và khởi tạo cấu trúc trống chuẩn để Merge không bị crash
    if df_baocao.empty and df_kytrước.empty:
        print("Cả hai kỳ dữ liệu lọc được đều trống.")
        return False
    elif df_baocao.empty:
        df_baocao = pd.DataFrame(columns=['Loại nghiệp vụ', 'CoSo', 'Loại điện', 'SoLuot', 'SoBo_ĐTC', 'SoTien_ĐTC', 'SoBo_KĐTC', 'SoTien_KĐTC'])
    elif df_kytrước.empty:
        df_kytrước = pd.DataFrame(columns=['Loại nghiệp vụ', 'CoSo', 'Loại điện', 'SoLuot', 'SoBo_ĐTC', 'SoTien_ĐTC', 'SoBo_KĐTC', 'SoTien_KĐTC'])

    # 🎯 ĐỒNG BỘ: Khớp nối dữ liệu hai kỳ qua cột 'Loại nghiệp vụ'
    merged_output = pd.merge(
        df_baocao, df_kytrước, 
        on=['Loại nghiệp vụ', 'CoSo', 'Loại điện'], 
        how='outer', 
        suffixes=('_kỳ báo cáo', '_Kì trước')
    ).fillna(0)
    
    rename_mapping = {
        'SoLuot_Kì trước': 'Số lươt kỳ trước',
        'SoLuot_kỳ báo cáo': 'Số lượt kỳ này',
        'SoBo_ĐTC_Kì trước': 'Số bó ĐTC Kì trước',
        'SoBo_ĐTC_kỳ báo cáo': 'Số Bó ĐTC kỳ báo cáo',
        'SoTien_ĐTC_Kì trước': 'Số tiền ĐTC kỳ trước',
        'SoTien_ĐTC_kỳ báo cáo': 'Số tiền ĐTC kỳ báo cáo',
        'SoBo_KĐTC_Kì trước': 'Số bó KĐTC Kì trước',
        'SoBo_KĐTC_kỳ báo cáo': 'Số Bó KĐTC kỳ báo cáo',
        'SoTien_KĐTC_Kì trước': 'Số tiền KĐTC kỳ trước',
        'SoTien_KĐTC_kỳ báo cáo': 'Số tiền KĐTC kỳ báo cáo',
        'CoSo': 'Cơ sở'
    }
    
    merged_output = merged_output.rename(columns=rename_mapping)
    
    def nghiệp_vụ_rank(val):
        v = str(val).lower()
        if "nộp" in v: return 1
        elif "lĩnh" in v or "linh" in v: return 2
        return 3

    coso_order = ["Cơ sở 1", "Cơ sở 2", "HUB NVC", "CS HĐ", "CS HP"]
    def coso_rank(val):
        return coso_order.index(val) if val in coso_order else 99

    def loaidien_rank(val):
        v = str(val).lower()
        if "tổng" in v: return 1
        elif "con" in v: return 2
        return 3

    # 🎯 ĐỒNG BỘ: Sắp xếp thứ tự ưu tiên dòng dựa trên cột 'Loại nghiệp vụ' chuẩn tiếng Việt
    if 'Loại nghiệp vụ' in merged_output.columns and 'Cơ sở' in merged_output.columns and 'Loại điện' in merged_output.columns:
        merged_output['NV_Rank'] = merged_output['Loại nghiệp vụ'].apply(nghiệp_vụ_rank)
        merged_output['CoSo_Rank'] = merged_output['Cơ sở'].apply(coso_rank)
        merged_output['Dien_Rank'] = merged_output['Loại điện'].apply(loaidien_rank)
        
        merged_output = merged_output.sort_values(by=['NV_Rank', 'CoSo_Rank', 'Dien_Rank']).drop(columns=['NV_Rank', 'CoSo_Rank', 'Dien_Rank'])

    # 🎯 ĐỒNG BỘ: Sắp xếp lại thứ tự phân bổ các cột trong file Excel báo cáo
    desired_columns = [
        'Loại nghiệp vụ', 'Cơ sở', 'Loại điện', 
        'Số lươt kỳ trước', 'Số lượt kỳ này', 
        'Số bó ĐTC Kì trước', 'Số Bó ĐTC kỳ báo cáo', 
        'Số tiền ĐTC kỳ trước', 'Số tiền ĐTC kỳ báo cáo', 
        'Số bó KĐTC Kì trước', 'Số Bó KĐTC kỳ báo cáo', 
        'Số tiền KĐTC kỳ trước', 'Số tiền KĐTC kỳ báo cáo'
    ]
    
    final_cols = [col for col in desired_columns if col in merged_output.columns]
    merged_output = merged_output[final_cols].copy()
    
    for c in final_cols[3:]:
        merged_output[c] = pd.to_numeric(merged_output[c], errors='coerce').fillna(0)

    # Ghi dữ liệu ra file Excel và format phân cách hàng nghìn (#,##0) thông qua openpyxl
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        merged_output.to_excel(writer, index=False, sheet_name='BaoCao_NHNN')
        worksheet = writer.sheets['BaoCao_NHNN']
        
        for col_idx in range(4, len(final_cols) + 1):
            col_letter = get_column_letter(col_idx)
            for row in range(2, len(merged_output) + 2):
                cell = worksheet[f"{col_letter}{row}"]
                if cell.value is not None:
                    try:
                        cell.value = float(cell.value)
                        cell.number_format = '#,##0'
                    except ValueError:
                        pass

    print(f"Đã xuất báo cáo NHNN thành công ra file: {output_excel_path}")
    return True

import streamlit as st
import pandas as pd
import os
import io

def _read_multi_header_file(file_path):
    """Đọc file Excel có dòng tiêu đề kép (multi-level headers)."""
    if not file_path or not os.path.exists(file_path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(file_path, header=[0, 1])
    except Exception:
        try:
            df = pd.read_excel(file_path, header=0)
        except Exception:
            return pd.DataFrame()
    return df

def _extract_period_data(df_raw, start_date, end_date, top_n):
    """Hàm phụ trợ lọc, sắp xếp theo thứ tự nghiệp vụ/cơ sở và lấy Top N theo kỳ."""
    if df_raw.empty:
        return pd.DataFrame(), pd.DataFrame()

    columns = df_raw.columns
    date_col_idx, coso_col_idx, nghiepvu_col_idx, khachhang_col_idx = None, None, None, None
    vnd_value_col_idx, usd_value_col_idx = None, None

    for idx, col in enumerate(columns):
        t1 = str(col[0]).strip().lower()
        t2 = str(col[1]).strip().lower()

        if 'ngày hiệu lực' in t2 or 'ngay hieu luc' in t2:
            date_col_idx = idx
        elif 'cơ sở' in t2 or 'coso' in t2:
            coso_col_idx = idx
        elif 'loại nghiệp vụ' in t2 or 'nghiep vu' in t2:
            nghiepvu_col_idx = idx
        elif 'tên khách hàng' in t2 or 'khách hàng' in t2 or 'đơn vị' in t2:
            khachhang_col_idx = idx

        if 'thành tiền' in t2 or 'thanh tien' in t2:
            if 'vnd' in t1 or 'đtc' in t1:
                vnd_value_col_idx = idx
            elif 'usd' in t1 or 'ngoại tệ' in t1:
                usd_value_col_idx = idx
            elif 'vnd' not in t1 and 'usd' not in t1 and 'unnamed' in t1:
                if vnd_value_col_idx is None:
                    vnd_value_col_idx = idx

    if date_col_idx is None or coso_col_idx is None:
        return pd.DataFrame(), pd.DataFrame()

    data_rows = []
    for row_idx in range(len(df_raw)):
        row_data = {}
        row_data['ngay_hieu_luc'] = df_raw.iloc[row_idx, date_col_idx]
        row_data['co_so'] = df_raw.iloc[row_idx, coso_col_idx]
        row_data['nghiep_vu'] = df_raw.iloc[row_idx, nghiepvu_col_idx] if nghiepvu_col_idx is not None else 'Giao dịch chung'
        row_data['ten_don_vi'] = df_raw.iloc[row_idx, khachhang_col_idx] if khachhang_col_idx is not None else 'Khác'
        
        row_data['val_vnd'] = df_raw.iloc[row_idx, vnd_value_col_idx] if vnd_value_col_idx is not None else 0.0
        row_data['val_usd'] = df_raw.iloc[row_idx, usd_value_col_idx] if usd_value_col_idx is not None else 0.0
        data_rows.append(row_data)

    df_flat = pd.DataFrame(data_rows)

    start_dt = pd.to_datetime(start_date)
    end_dt = pd.to_datetime(end_date)

    df_flat['__temp_date'] = pd.to_datetime(df_flat['ngay_hieu_luc'], errors='coerce')
    df_filtered = df_flat[(df_flat['__temp_date'] >= start_dt) & (df_flat['__temp_date'] <= end_dt)].copy()

    if df_filtered.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_filtered['__CoSo'] = df_filtered['co_so'].astype(str).str.strip()
    df_filtered['__NgayHieuLuc'] = df_filtered['__temp_date'].dt.strftime('%d/%m/%Y')
    df_filtered['__NghiepVu'] = df_filtered['nghiep_vu'].astype(str).str.strip()
    df_filtered['__DonVi'] = df_filtered['ten_don_vi'].astype(str).str.strip()
    
    df_filtered['__ValVND'] = pd.to_numeric(df_filtered['val_vnd'], errors='coerce').fillna(0.0)
    df_filtered['__ValUSD'] = pd.to_numeric(df_filtered['val_usd'], errors='coerce').fillna(0.0)

    # ĐỊNH NGHĨA THỨ TỰ SẮP XẾP CHUẨN
    coso_order = ["Cơ sở 1", "Cơ sở 2", "Hub NVC", "CS HĐ", "CS HP"]
    nghiepvu_order = ["Tiếp quỹ đầu ngày", "Tiếp quỹ trong ngày", "Thu gom trong ngày", "Thu gom cuối ngày"]

    df_filtered['__CoSo_Cat'] = pd.Categorical(df_filtered['__CoSo'], categories=coso_order, ordered=True)
    df_filtered['__NghiepVu_Cat'] = pd.Categorical(df_filtered['__NghiepVu'], categories=nghiepvu_order, ordered=True)

    def get_top_n_currency(sub_df, val_col_key):
        if sub_df.empty:
            return pd.DataFrame()
        valid_df = sub_df[sub_df[val_col_key] > 0].copy()
        if valid_df.empty:
            return pd.DataFrame()

        # Sắp xếp theo thứ tự Cơ sở -> Nghiệp vụ -> Giá trị giảm dần
        valid_df.sort_values(by=['__CoSo_Cat', '__NghiepVu_Cat', val_col_key], ascending=[True, True, False], inplace=True)

        grouped = []
        # Duyệt theo đúng thứ tự cơ sở và nghiệp vụ đã định nghĩa
        for coso in coso_order:
            df_coso = valid_df[valid_df['__CoSo'] == coso]
            if df_coso.empty:
                continue
            for nv in nghiepvu_order:
                df_nv = df_coso[df_coso['__NghiepVu'] == nv]
                if not df_nv.empty:
                    top_g = df_nv.sort_values(by=val_col_key, ascending=False).head(top_n)
                    grouped.append(top_g)
        
        # Bổ sung các nghiệp vụ phát sinh ngoài danh sách chuẩn (nếu có) vào cuối bảng
        other_df = valid_df[~valid_df['__NghiepVu'].isin(nghiepvu_order)]
        if not other_df.empty:
            grouped.append(other_df)

        return pd.concat(grouped, ignore_index=True) if grouped else pd.DataFrame()

    top_vnd_raw = get_top_n_currency(df_filtered, '__ValVND')
    top_usd_raw = get_top_n_currency(df_filtered, '__ValUSD')

    def format_final_df(raw_df, val_key):
        if raw_df.empty:
            return pd.DataFrame(columns=['Tên cơ sở', 'Ngày hiệu lực', 'Nghiệp vụ', 'Tên đơn vị', 'Số tiền'])
        out = pd.DataFrame()
        out['Tên cơ sở'] = raw_df['__CoSo']
        out['Ngày hiệu lực'] = raw_df['__NgayHieuLuc']
        out['Nghiệp vụ'] = raw_df['__NghiepVu']
        out['Tên đơn vị'] = raw_df['__DonVi']
        out['Số tiền'] = raw_df[val_key]
        return out

    return format_final_df(top_vnd_raw, '__ValVND'), format_final_df(top_usd_raw, '__ValUSD')

def process_giao_dich_lon_report(shared_data_dir, start_curr, end_curr, start_prev, end_prev, top_n=10):
    file_path = os.path.join(shared_data_dir, "Ds_TiepQuyThuGom.xlsx")
    if not os.path.exists(file_path):
        return (pd.DataFrame(), pd.DataFrame()), (pd.DataFrame(), pd.DataFrame())

    df_raw = _read_multi_header_file(file_path)
    if df_raw.empty:
        return (pd.DataFrame(), pd.DataFrame()), (pd.DataFrame(), pd.DataFrame())

    vnd_curr, usd_curr = _extract_period_data(df_raw, start_curr, end_curr, top_n)
    vnd_prev, usd_prev = _extract_period_data(df_raw, start_prev, end_prev, top_n)

    return (vnd_curr, usd_curr), (vnd_prev, usd_prev)

def export_to_excel(vnd_curr, usd_curr, vnd_prev, usd_prev):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1F4E78', 'font_color': 'white', 'border': 1
        })
        data_format = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        number_format = workbook.add_format({'border': 1, 'valign': 'vcenter', 'align': 'right', 'num_format': '#,##0.00'})

        sheets_data = [
            ('VND Kỳ Này', vnd_curr),
            ('USD Kỳ Này', usd_curr),
            ('VND Kỳ Trước', vnd_prev),
            ('USD Kỳ Trước', usd_prev)
        ]

        for sheet_name, df in sheets_data:
            if not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                for col_num, col_name in enumerate(df.columns):
                    ws.write(0, col_num, col_name, header_format)
                    ws.set_column(col_num, col_num, 22, number_format if col_name == 'Số tiền' else data_format)
                ws.freeze_panes(1, 0)

    return output.getvalue()
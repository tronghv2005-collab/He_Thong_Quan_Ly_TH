import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import os
import openpyxl
from io import BytesIO

DENOMINATIONS = [500000, 200000, 100000, 50000, 20000, 10000, 5000, 2000, 1000]

def _parse_int_formatted(val_raw):
    """Chuyển chuỗi có dấu phẩy hàng ngàn về số nguyên không âm"""
    try:
        clean_str = str(val_raw or "").replace(",", "").strip()
        num_val = int(float(clean_str)) if clean_str else 0
        return max(0, num_val)
    except:
        return 0

def init_default_gl(path_gl):
    """Khởi tạo file gl_tienmo.xlsx với đầy đủ giá trị 0.0 cho tất cả các ô số liệu chuẩn 2 khối"""
    if not os.path.exists(path_gl):
        row0 = [np.nan, np.nan, 'TRONG KHO'] + [np.nan]*18 + ['DON VI'] + [np.nan]*18
        row1 = ['MA DON VI', 'TEN DON VI']
        for _ in range(2):
            for d in DENOMINATIONS:
                row1.extend([d, 'THANHTIEN'])
            row1.append('TONG THANH TIEN')
            
        data = [
            row0, 
            row1,
            ['990', 'TOÀN TRUNG TÂM'] + [0.0]*38,
            ['990.1', 'KHO 194 TRẦN QUANG KHẢI'] + [0.0]*38,
            ['990.3', 'KHO CƠ SỞ 2'] + [0.0]*38,
            ['990.3.GDV', 'GDV CƠ SỞ 2'] + [0.0]*38,
            ['990.5', 'KHO CƠ SỞ HẢI PHÒNG'] + [0.0]*38,
            ['990.5.GDV', 'GDV CƠ SỞ HẢI PHÒNG'] + [0.0]*38,
            ['990.6', 'KHO CƠ SỞ HƯNG YÊN'] + [0.0]*38,
            ['990.6.GDV', 'GDV CƠ SỞ HƯNG YÊN'] + [0.0]*38,
            ['990.7', 'KHO HUB NVC'] + [0.0]*38,
            ['990.7.GDV', 'GDV HUB NVC'] + [0.0]*38
        ]
        df_init = pd.DataFrame(data)
        df_init.to_excel(path_gl, index=False, header=False)

def get_list_accounts_from_gl(path_gl):
    """Lấy danh sách mã đơn vị / khách hàng từ file gl_tienmo.xlsx để cho phép người dùng chọn"""
    accounts = []
    if os.path.exists(path_gl):
        try:
            wb = openpyxl.load_workbook(path_gl)
            ws = wb.active
            for r in range(3, ws.max_row + 1):
                ma_dv = ws.cell(row=r, column=1).value
                ten_dv = ws.cell(row=r, column=2).value
                if ma_dv is not None and str(ma_dv).strip() != '':
                    accounts.append(f"{str(ma_dv).strip()} - {str(ten_dv).strip() if ten_dv else ''}")
        except:
            pass
    return accounts

def get_donvi_denominations_from_gl(path_gl, kho_code):
    """Đọc số tờ theo từng mệnh giá ở khối DON VI của mã kho/khách hàng trong gl_tienmo.xlsx"""
    result = {mg: 0 for mg in DENOMINATIONS}
    if not kho_code or str(kho_code).strip() == '' or not os.path.exists(path_gl):
        return result

    try:
        wb = openpyxl.load_workbook(path_gl, data_only=True)
        ws = wb.active
        header_row = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]

        dv_to_cols = {}
        for col_idx_1based in range(22, min(40, ws.max_column + 1)):
            val = header_row[col_idx_1based - 1]
            try:
                mg_val = int(float(val))
                if mg_val in DENOMINATIONS:
                    dv_to_cols[mg_val] = col_idx_1based
            except:
                pass

        target_str = str(kho_code).strip().lower()
        target_row = None
        for r in range(3, ws.max_row + 1):
            cell_val = ws.cell(row=r, column=1).value
            if cell_val is not None:
                cell_str = str(cell_val).strip().lower()
                if cell_str == target_str or cell_str == target_str.replace('.0', ''):
                    target_row = r
                    break

        if target_row is not None:
            for mg in DENOMINATIONS:
                if mg in dv_to_cols:
                    val_to = ws.cell(row=target_row, column=dv_to_cols[mg]).value
                    try:
                        so_to = int(float(val_to)) if val_to is not None and str(val_to).strip() != '' else 0
                    except:
                        so_to = 0
                    result[mg] = max(0, so_to)
    except:
        pass

    return result

def apply_so_to_to_session_inputs(so_to_map, unit_mode="Số tờ"):
    """Đổ bản đồ số tờ theo mệnh giá vào các widget nhập liệu theo đơn vị tính hiện hành"""
    for mg in DENOMINATIONS:
        so_to = int(so_to_map.get(mg, 0) or 0)
        if unit_mode.startswith("Số tờ"):
            st.session_state[f"to_{mg}"] = f"{so_to:,}"
        elif unit_mode.startswith("Theo Bó"):
            st.session_state[f"bo_{mg}"] = round(so_to / 1000.0, 2)
        else:
            st.session_state[f"bao_{mg}"] = round(so_to / 20000.0, 2)

def get_list_accounts_from_coso(path_coso):
    """Lấy danh sách đơn vị từ file Thamso_coso.xlsx phục vụ cho nghiệp vụ Nội bộ"""
    coso_accounts = []
    if os.path.exists(path_coso):
        try:
            df_cs = pd.read_excel(path_coso)
            df_cs.columns = [str(c).strip() for c in df_cs.columns]
            col_kho = next((c for c in df_cs.columns if 'kho' in c.lower()), None)
            col_ten = next((c for c in df_cs.columns if 'tên' in c.lower() or 'cơ sở' in c.lower() or 'ten' in c.lower()), None)
            
            if col_kho:
                for _, row in df_cs.iterrows():
                    m_kho = str(row[col_kho]).strip()
                    t_kho = str(row[col_ten]).strip() if col_ten and pd.notna(row[col_ten]) else ''
                    if m_kho and m_kho.lower() != 'nan':
                        coso_accounts.append(f"{m_kho} - {t_kho}")
        except:
            pass
    return coso_accounts

def get_ld_hub_coso_from_canbo(param_dir, current_user):
    """Đọc CHÍNH XÁC giá trị từ cột 'CoSo' trong file Thamso_canbo.xlsx bằng cách quét cả MaNV và TenNV"""
    path_canbo = os.path.join(param_dir, "Thamso_canbo.xlsx")
    if not os.path.exists(path_canbo):
        return ""
    try:
        df_cb = pd.read_excel(path_canbo)
        df_cb.columns = [str(c).strip() for c in df_cb.columns]
        
        col_coso = next((c for c in df_cb.columns if c.lower() == 'coso'), None)
        if not col_coso:
            return ""
            
        u_clean = str(current_user).strip().lower()
        
        for _, row in df_cb.iterrows():
            ma_val = str(row.get('MaNV', '')).strip().lower()
            ten_val = str(row.get('TenNV', '')).strip().lower()
            coso_val = str(row.get(col_coso, '')).strip()
            
            if u_clean == ma_val or u_clean == ten_val or u_clean in ten_val or ten_val in u_clean:
                if coso_val and coso_val.lower() != 'nan':
                    return coso_val
    except:
        pass
    return ""

def ensure_row_exists_in_gl(ws, kho_code, ten_kho=""):
    """Tự động thêm dòng mới vào file Sổ cái nếu mã đơn vị/kho chưa tồn tại (giữ nguyên tuyệt đối mã định danh)"""
    if not kho_code or str(kho_code).strip() == '' or str(kho_code).strip().lower() == 'nan':
        return None
    
    target_str = str(kho_code).strip().lower()
    for r in range(3, ws.max_row + 1):
        cell_val = ws.cell(row=r, column=1).value
        if cell_val is not None:
            cell_str = str(cell_val).strip().lower()
            if cell_str == target_str or cell_str == target_str.replace('.0', ''):
                return r
                
    new_row_idx = ws.max_row + 1
    ws.cell(row=new_row_idx, column=1, value=str(kho_code).strip())
    ws.cell(row=new_row_idx, column=2, value=str(ten_kho).strip() if ten_kho else f"ĐƠN VỊ {str(kho_code).strip()}")
    for c in range(3, ws.max_column + 1):
        ws.cell(row=new_row_idx, column=c, value=0.0)
    return new_row_idx

def _norm_code(code_val):
    s = str(code_val or '').strip().lower()
    return s.replace('.0', '') if s.endswith('.0') else s

def _extract_row_ccmg(row_series):
    ccmg = {}
    for mg in DENOMINATIONS:
        val_to = 0.0
        if mg in row_series and pd.notna(row_series[mg]):
            val_to = row_series[mg]
        elif str(mg) in row_series and pd.notna(row_series[str(mg)]):
            val_to = row_series[str(mg)]
        try:
            ccmg[mg] = float(val_to)
        except:
            ccmg[mg] = 0.0
    return ccmg

def _get_outflow_targets_for_tx(thu_chi_val, loai_nv_val, don_vi_xl_val, ma_kh_val):
    """
    Trả về danh sách các tuple: (ma_tai_khoan_norm, khoi_key: 'TK' | 'DV')
    đại diện cho các vị trí bị xuất/giảm số tờ khi thực hiện giao dịch.
    """
    tc = str(thu_chi_val or '').strip().upper()
    lnv = str(loai_nv_val or '').strip()
    dvxl = str(don_vi_xl_val or '').strip()
    makh = str(ma_kh_val or '').strip()
    is_gdv = ".gdv" in dvxl.lower()

    targets = []
    if lnv == "Điều chuyển CN":
        if dvxl:
            targets.append((_norm_code(dvxl), "DV"))
        targets.append((_norm_code("990"), "DV"))
        return targets

    if tc == "CHI":
        if lnv == "Chi nhánh":
            if dvxl:
                targets.append((_norm_code(dvxl), "TK"))
            if makh:
                targets.append((_norm_code(makh), "DV"))
        elif lnv == "Cá Nhân":
            if is_gdv:
                targets.append((_norm_code(dvxl), "DV"))
            else:
                targets.append((_norm_code(dvxl), "TK"))
                targets.append((_norm_code(dvxl), "DV"))
                targets.append((_norm_code("990"), "TK"))
                targets.append((_norm_code("990"), "DV"))
        else:
            if is_gdv:
                targets.append((_norm_code(dvxl), "DV"))
            else:
                targets.append((_norm_code(dvxl), "TK"))
                targets.append((_norm_code(dvxl), "DV"))

            if lnv in ["Tạm ứng"]:
                targets.append((_norm_code("990"), "TK"))
    elif tc == "THU" and lnv == "Hoàn ứng":
        if makh and lnv != "Cá Nhân":
            targets.append((_norm_code(makh), "DV"))
    return targets

def compute_pending_outflow_by_status(df_tx, exclude_seq=None):
    """
    Tính tổng số tờ đang chiếm chỗ/xuất giảm cho từng (ma_code_norm, block_key, mg)
    tách riêng theo 2 trạng thái: 'Duyệt' và 'Chờ Phê Duyệt'.
    """
    pending_duyet = {}
    pending_cho = {}

    if df_tx is None or df_tx.empty or 'TRANG THAI' not in df_tx.columns:
        return pending_duyet, pending_cho

    for _, r in df_tx.iterrows():
        stt = str(r.get('TRANG THAI', '')).strip()
        if stt not in ['Duyệt', 'Chờ Phê Duyệt']:
            continue
        if exclude_seq is not None and pd.notna(r.get('SO SEQ')):
            try:
                if float(r.get('SO SEQ')) == float(exclude_seq):
                    continue
            except:
                pass

        ccmg = _extract_row_ccmg(r)
        out_targets = _get_outflow_targets_for_tx(
            r.get('THU CHI', ''),
            r.get('LOAI NV', ''),
            r.get('DON VI XL', ''),
            r.get('MA KH CCCD', '')
        )
        bucket = pending_duyet if stt == 'Duyệt' else pending_cho
        for code_norm, blk in out_targets:
            if not code_norm:
                continue
            for mg in DENOMINATIONS:
                so_to = float(ccmg.get(mg, 0) or 0)
                if so_to > 0:
                    k = (code_norm, blk, mg)
                    bucket[k] = bucket.get(k, 0.0) + so_to

    return pending_duyet, pending_cho

def validate_chi_kha_dung(path_gl, thu_chi_val, loai_nv_val, don_vi_xl_val, ma_kh_val, ccmg_data, df_all_tx=None, exclude_seq=None):
    """
    Kiểm tra tồn khả dụng cho các nghiệp vụ CHI, Điều chuyển CN (và THU Hoàn ứng) ngay tại thời điểm
    Lập/Lưu đẩy duyệt, Sửa cơ cấu, Phê duyệt hoặc Hạch toán XL.
    Trả về (is_valid: bool, error_message: str).
    """
    tc = str(thu_chi_val or '').strip().upper()
    lnv = str(loai_nv_val or '').strip()
    dvxl = str(don_vi_xl_val or '').strip()
    makh = str(ma_kh_val or '').strip()

    if not os.path.exists(path_gl):
        return True, ""

    try:
        wb = openpyxl.load_workbook(path_gl, data_only=True)
        ws = wb.active
    except Exception as e:
        return False, f"Không thể đọc file Sổ cái gl_tienmo.xlsx để kiểm tra tồn: {e}"

    header_row = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]

    def get_denomination_cols_wb(start_col, end_col):
        cols_map = {}
        for col_idx_1based in range(start_col, end_col + 1):
            val = header_row[col_idx_1based - 1]
            try:
                mg_val = int(float(val))
                if mg_val in DENOMINATIONS:
                    cols_map[mg_val] = {
                        'to_col': col_idx_1based,
                        'tt_col': col_idx_1based + 1
                    }
            except:
                pass
        return cols_map

    tk_map = get_denomination_cols_wb(3, 20)
    dv_map = get_denomination_cols_wb(22, 39)

    def get_row_index_wb(kho_code):
        if not kho_code or str(kho_code).strip() == '' or str(kho_code).strip().lower() == 'nan':
            return None
        target_str = str(kho_code).strip().lower()
        for r in range(3, ws.max_row + 1):
            cell_val = ws.cell(row=r, column=1).value
            if cell_val is not None:
                cell_str = str(cell_val).strip().lower()
                if cell_str == target_str or cell_str == target_str.replace('.0', ''):
                    return r
        return None

    pending_duyet, pending_cho = compute_pending_outflow_by_status(df_all_tx, exclude_seq=exclude_seq)

    def check_block(target_row, kho_block_map, blk_key, code_raw, block_label):
        c_norm = _norm_code(code_raw)
        for mg in DENOMINATIONS:
            so_to_yc = float(ccmg_data.get(mg, 0) or 0)
            if so_to_yc > 0:
                if target_row is None:
                    return False, (
                        f"❌ Chặn giao dịch tại '{code_raw}' ({block_label}): Không tìm thấy tài khoản trên Sổ cái "
                        f"(Mệnh giá {mg:,} VNĐ: Khả dụng 0 tờ, yêu cầu {int(so_to_yc):,} tờ, thiếu {int(so_to_yc):,} tờ)."
                    )
                if mg in kho_block_map:
                    to_col = kho_block_map[mg]['to_col']
                    val = ws.cell(row=target_row, column=to_col).value
                    try:
                        hien_co = float(val) if val is not None and str(val).strip() != '' else 0.0
                    except:
                        hien_co = 0.0

                    da_duyet = float(pending_duyet.get((c_norm, blk_key, mg), 0.0))
                    da_cho = float(pending_cho.get((c_norm, blk_key, mg), 0.0))
                    kha_dung = hien_co - (da_duyet + da_cho)

                    if so_to_yc > kha_dung:
                        thieu = so_to_yc - max(0.0, kha_dung)
                        return False, (
                            f"❌ Không đủ số tờ khả dụng tại '{code_raw}' - {block_label} "
                            f"(Mệnh giá {mg:,} VNĐ: Tồn GL {int(hien_co):,} tờ; GD 'Duyệt' giữ {int(da_duyet):,} tờ + "
                            f"GD 'Chờ Phê Duyệt' giữ {int(da_cho):,} tờ => Khả dụng {int(kha_dung):,} tờ, "
                            f"yêu cầu {int(so_to_yc):,} tờ, thiếu {int(thieu):,} tờ)."
                        )
        return True, ""

    is_gdv = ".gdv" in dvxl.lower()
    idx_dvxl = get_row_index_wb(dvxl)
    idx_990 = get_row_index_wb("990")
    idx_makh = get_row_index_wb(makh) if (makh and lnv != "Cá Nhân") else None

    if lnv == "Điều chuyển CN":
        ok_dvxl_dv, msg_dvxl_dv = check_block(
            idx_dvxl, dv_map, "DV", dvxl, "Khối DON VI của DON VI XL (Điều chuyển CN)"
        )
        if not ok_dvxl_dv:
            return False, msg_dvxl_dv

        ok_990_dv, msg_990_dv = check_block(
            idx_990, dv_map, "DV", "990", "Kho Tổng 990 - Khối DON VI (Điều chuyển CN)"
        )
        if not ok_990_dv:
            return False, msg_990_dv

    elif tc == "CHI":
        if lnv == "Chi nhánh":
            ok_dvxl_tk, msg_dvxl_tk = check_block(
                idx_dvxl, tk_map, "TK", dvxl, "Khối TRONG KHO của DON VI XL (Chi Chi Nhánh)"
            )
            if not ok_dvxl_tk:
                return False, msg_dvxl_tk

            ok_makh_dv, msg_makh_dv = check_block(
                idx_makh, dv_map, "DV", makh, f"Khối DON VI của MA KH CCCD '{makh}' (Chi Chi Nhánh)"
            )
            if not ok_makh_dv:
                return False, msg_makh_dv
        elif lnv == "Cá Nhân":
            if is_gdv:
                ok, msg = check_block(idx_dvxl, dv_map, "DV", dvxl, "Khối DON VI (GDV Chi Cá Nhân)")
                if not ok:
                    return False, msg
            else:
                ok_tk, msg_tk = check_block(idx_dvxl, tk_map, "TK", dvxl, "Khối TRONG KHO (Kho Chi Cá Nhân)")
                if not ok_tk:
                    return False, msg_tk
                ok_dv, msg_dv = check_block(idx_dvxl, dv_map, "DV", dvxl, "Khối DON VI (Kho Chi Cá Nhân)")
                if not ok_dv:
                    return False, msg_dv
                ok_990_tk, msg_990_tk = check_block(idx_990, tk_map, "TK", "990", "Kho Tổng 990 - Khối TRONG KHO (Chi Cá Nhân)")
                if not ok_990_tk:
                    return False, msg_990_tk
                ok_990_dv, msg_990_dv = check_block(idx_990, dv_map, "DV", "990", "Kho Tổng 990 - Khối DON VI (Chi Cá Nhân)")
                if not ok_990_dv:
                    return False, msg_990_dv
        else:
            if is_gdv:
                ok_dv, msg_dv = check_block(idx_dvxl, dv_map, "DV", dvxl, "Khối DON VI")
                if not ok_dv:
                    return False, msg_dv
            else:
                ok_tk, msg_tk = check_block(idx_dvxl, tk_map, "TK", dvxl, "Khối TRONG KHO")
                if not ok_tk:
                    return False, msg_tk
                ok_dv, msg_dv = check_block(idx_dvxl, dv_map, "DV", dvxl, "Khối DON VI")
                if not ok_dv:
                    return False, msg_dv

            if lnv in ["Tạm ứng"]:
                ok_990, msg_990 = check_block(idx_990, tk_map, "TK", "990", "Kho Tổng 990 (TRONG KHO)")
                if not ok_990:
                    return False, msg_990

    elif tc == "THU" and lnv == "Hoàn ứng":
        if makh and lnv != "Cá Nhân":
            ok_makh, msg_makh = check_block(idx_makh, dv_map, "DV", makh, f"Khối DON VI của MA KH CCCD '{makh}'")
            if not ok_makh:
                return False, msg_makh

    return True, ""

def execute_hach_toan_kho(path_gl, row_item, ccmg_data, df_all_tx=None, current_seq=None):
    """
    Hạch toán chuẩn xác vào file gl_tienmo.xlsx đúng theo các nghiệp vụ.
    Kiểm tra chặt chẽ tồn khả dụng sau khi trừ cả 2 loại giao dịch chưa XL:
    'Duyệt' và 'Chờ Phê Duyệt' (nếu vượt sẽ chặn, không cho thực hiện để GDV sửa).
    """
    if not os.path.exists(path_gl):
        return False, "File Sổ cái gl_tienmo.xlsx không tồn tại!"
        
    try:
        wb = openpyxl.load_workbook(path_gl)
        ws = wb.active
    except Exception as e:
        return False, f"Không thể mở file gl_tienmo.xlsx: {e}"
        
    for r in range(3, ws.max_row + 1):
        for c in range(3, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is None or str(cell.value).strip() == '' or str(cell.value).strip().lower() == 'nan':
                cell.value = 0.0

    thu_chi = str(row_item.get('THU CHI', '')).strip().upper()
    loai_nv = str(row_item.get('LOAI NV', '')).strip()
    don_vi_xl = str(row_item.get('DON VI XL', '')).strip()
    ma_kh = str(row_item.get('MA KH CCCD', '')).strip()

    ok_check, msg_check = validate_chi_kha_dung(
        path_gl,
        thu_chi_val=thu_chi,
        loai_nv_val=loai_nv,
        don_vi_xl_val=don_vi_xl,
        ma_kh_val=ma_kh,
        ccmg_data=ccmg_data,
        df_all_tx=df_all_tx,
        exclude_seq=current_seq
    )
    if not ok_check:
        return False, msg_check
    
    header_row = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    
    def get_denomination_cols_wb(start_col, end_col):
        cols_map = {}
        for col_idx_1based in range(start_col, end_col + 1):
            val = header_row[col_idx_1based - 1]
            try:
                mg_val = int(float(val))
                if mg_val in DENOMINATIONS:
                    cols_map[mg_val] = {
                        'to_col': col_idx_1based,
                        'tt_col': col_idx_1based + 1
                    }
            except:
                pass
        return cols_map

    tk_map = get_denomination_cols_wb(3, 20)   # Khối TRONG KHO
    dv_map = get_denomination_cols_wb(22, 39) # Khối DON VI

    def get_row_index_wb(kho_code):
        if not kho_code or str(kho_code).strip() == '' or str(kho_code).strip().lower() == 'nan':
            return None
        target_str = str(kho_code).strip().lower()
        for r in range(3, ws.max_row + 1):
            cell_val = ws.cell(row=r, column=1).value
            if cell_val is not None:
                cell_str = str(cell_val).strip().lower()
                if cell_str == target_str or cell_str == target_str.replace('.0', ''):
                    return r
        return None

    idx_990 = get_row_index_wb('990')
    idx_dvxl = ensure_row_exists_in_gl(ws, don_vi_xl)
    idx_makh = get_row_index_wb(ma_kh) if loai_nv != "Cá Nhân" else None

    if idx_dvxl is None:
        return False, f"Lỗi: Không tìm thấy đơn vị xử lý '{don_vi_xl}' trong Sổ cái!"

    is_gdv_account = ".gdv" in don_vi_xl.lower()

    def apply_delta_block_wb(target_row, is_tang, delta_dict, kho_block_map, sum_col_idx):
        if target_row is None:
            return
        total_tt_delta = 0
        for mg in DENOMINATIONS:
            so_to = float(delta_dict.get(mg, 0) or 0)
            if so_to == 0:
                continue
            thanh_tien = so_to * mg
            total_tt_delta += thanh_tien if is_tang else -thanh_tien
            
            if mg in kho_block_map:
                to_col = kho_block_map[mg]['to_col']
                tt_col = kho_block_map[mg]['tt_col']
                
                cur_to_val = ws.cell(row=target_row, column=to_col).value
                cur_to = float(cur_to_val) if cur_to_val is not None and str(cur_to_val).strip() != '' else 0.0
                new_to = (cur_to + so_to) if is_tang else max(0.0, cur_to - so_to)
                cell_to = ws.cell(row=target_row, column=to_col, value=new_to)
                cell_to.number_format = '#,##0'
                
                cur_tt_val = ws.cell(row=target_row, column=tt_col).value
                cur_tt = float(cur_tt_val) if cur_tt_val is not None and str(cur_tt_val).strip() != '' else 0.0
                new_tt = (cur_tt + thanh_tien) if is_tang else max(0.0, cur_tt - thanh_tien)
                cell_tt = ws.cell(row=target_row, column=tt_col, value=new_tt)
                cell_tt.number_format = '#,##0'
                
        if sum_col_idx <= ws.max_column:
            cur_sum_val = ws.cell(row=target_row, column=sum_col_idx).value
            cur_sum = float(cur_sum_val) if cur_sum_val is not None and str(cur_sum_val).strip() != '' else 0.0
            new_sum = max(0.0, cur_sum + total_tt_delta)
            cell_sum = ws.cell(row=target_row, column=sum_col_idx, value=new_sum)
            cell_sum.number_format = '#,##0'

    if thu_chi == "THU" and loai_nv == "Chi nhánh":
        if idx_makh is None:
            return False, f"Lỗi: Mã chi nhánh '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        apply_delta_block_wb(idx_990, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
            
    elif thu_chi == "THU" and loai_nv == "NHNN":
        apply_delta_block_wb(idx_990, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, True, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, dv_map, 40)
        
    elif thu_chi == "CHI" and loai_nv == "Nội bộ":
        if idx_makh is None:
            return False, f"Lỗi: Đơn vị nhận nội bộ '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_makh, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
            
    elif thu_chi == "CHI" and loai_nv == "Chi nhánh":
        if idx_makh is None:
            return False, f"Lỗi: Mã chi nhánh '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, False, ccmg_data, dv_map, 40)
        
    elif thu_chi == "CHI" and loai_nv == "Cá Nhân":
        if is_gdv_account:
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        else:
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
            apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_990, False, ccmg_data, dv_map, 40)
            
    elif thu_chi == "CHI" and loai_nv == "Tạm ứng":
        apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        if idx_makh:
            apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
            
    elif thu_chi == "THU" and loai_nv == "Hoàn ứng":
        apply_delta_block_wb(idx_990, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, True, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, dv_map, 40)
        if idx_makh:
            apply_delta_block_wb(idx_makh, False, ccmg_data, dv_map, 40)
        
    elif loai_nv == "Điều chuyển CN":
        if idx_makh is None:
            return False, f"Lỗi: Mã chi nhánh nhận điều chuyển '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_990, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
        
    elif loai_nv == "Điều chỉnh kiểm kê":
        if thu_chi == "THU":
            apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, True, ccmg_data, dv_map, 40)
        else:
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        
    else:
        if is_gdv_account:
            apply_delta_block_wb(idx_dvxl, thu_chi == "THU", ccmg_data, dv_map, 40)
        else:
            apply_delta_block_wb(idx_dvxl, thu_chi == "THU", ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, thu_chi == "THU", ccmg_data, dv_map, 40)

    wb.save(path_gl)
    return True, "Hạch toán kho thành công!"

def show_nhap_tien_moi_page(base_dir, shared_data_dir, current_user="System", user_role="NV"):
    # Cấu hình Custom CSS hiện đại & tối ưu đồng bộ Desktop + Mobile (iPhone/iPad)
    st.markdown("""
        <style>
            .app-header-banner {
                background: linear-gradient(135deg, #0f294a 0%, #1d4ed8 100%);
                color: #ffffff;
                padding: 16px 20px;
                border-radius: 12px;
                margin-bottom: 14px;
                box-shadow: 0 4px 12px rgba(15, 41, 74, 0.14);
            }
            .app-header-title {
                font-size: 1.25rem;
                font-weight: 700;
                margin: 0 0 6px 0;
                letter-spacing: 0.2px;
            }
            .app-header-meta {
                font-size: 0.9rem;
                opacity: 0.95;
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
                align-items: center;
            }
            .role-pill {
                background: rgba(255, 255, 255, 0.18);
                padding: 2px 10px;
                border-radius: 999px;
                font-weight: 600;
                font-size: 0.82rem;
            }
            .mobile-selection-bar {
                background: #f0f7ff;
                border: 1px solid #bfdbfe;
                border-left: 4px solid #2563eb;
                padding: 10px 14px;
                border-radius: 8px;
                margin: 10px 0;
                font-size: 0.9rem;
            }
            div.stButton > button:first-child {
                border-radius: 8px;
                font-weight: 600;
                min-height: 42px;
            }
            @media (max-width: 768px) {
                .block-container {
                    padding-left: 0.75rem !important;
                    padding-right: 0.75rem !important;
                    padding-top: 1rem !important;
                }
                .app-header-banner {
                    padding: 12px 14px;
                    border-radius: 10px;
                }
                .app-header-title {
                    font-size: 1.05rem;
                }
                .stRadio > div[role="radiogroup"] {
                    display: flex !important;
                    flex-direction: row !important;
                    flex-wrap: nowrap !important;
                    overflow-x: auto !important;
                    -webkit-overflow-scrolling: touch;
                    gap: 8px !important;
                    padding-bottom: 6px;
                }
                .stRadio > div[role="radiogroup"] > label {
                    flex: 0 0 auto !important;
                    background: #f8fafc;
                    border: 1px solid #cbd5e1;
                    border-radius: 8px;
                    padding: 6px 10px !important;
                    font-size: 0.82rem !important;
                    white-space: nowrap;
                }
                [data-testid="stMetricValue"] {
                    font-size: 1.15rem !important;
                }
                div.stButton > button:first-child {
                    min-height: 46px;
                    font-size: 0.9rem;
                }
            }
        </style>
    """, unsafe_allow_html=True)

    param_dir = os.path.join(base_dir, "parameters")
    os.makedirs(shared_data_dir, exist_ok=True)
    os.makedirs(param_dir, exist_ok=True)
    
    path_ds = os.path.join(shared_data_dir, "Ds_giaodichtienmoi.xlsx")
    path_gl = os.path.join(shared_data_dir, "gl_tienmo.xlsx")
    path_coso = os.path.join(param_dir, "Thamso_coso.xlsx")
    path_kh = os.path.join(param_dir, "Thamso_KH.xlsx")
    
    init_default_gl(path_gl)
    
    if os.path.exists(path_kh):
        try:
            df_kh_check = pd.read_excel(path_kh)
            df_kh_check.columns = [str(c).strip().upper() for c in df_kh_check.columns]
            needed_cols = ['MA CCCD', 'TEN KH', 'DIACHI KH']
            changed = False
            for col in needed_cols:
                if col not in df_kh_check.columns:
                    found = False
                    for ec in df_kh_check.columns:
                        if col.replace(" ", "") in ec.replace(" ", ""):
                            df_kh_check.rename(columns={ec: col}, inplace=True)
                            found = True
                            break
                    if not found:
                        df_kh_check[col] = ''
                    changed = True
            if changed:
                df_kh_check.to_excel(path_kh, index=False)
        except:
            df_kh_init = pd.DataFrame(columns=['MA CCCD', 'TEN KH', 'DIACHI KH'])
            df_kh_init.to_excel(path_kh, index=False)
    else:
        df_kh_init = pd.DataFrame(columns=['MA CCCD', 'TEN KH', 'DIACHI KH'])
        df_kh_init.to_excel(path_kh, index=False)
    
    role_upper = str(user_role).strip().upper()
    is_giam_doc = role_upper == "GIÁM ĐỐC"
    is_ld_hub = role_upper == "LĐ_HUB"

    assigned_coso_ld_hub = get_ld_hub_coso_from_canbo(param_dir, current_user) if is_ld_hub else ""

    coso_options = ["CS1 - Cơ sở 1", "CS2 - Cơ sở 2", "990 - Kho Tổng"]
    if os.path.exists(path_coso):
        try:
            df_cs = pd.read_excel(path_coso)
            df_cs.columns = [str(c).strip() for c in df_cs.columns]
            col_kho = next((c for c in df_cs.columns if 'kho' in c.lower()), None)
            col_ten = next((c for c in df_cs.columns if 'tên' in c.lower() or 'cơ sở' in c.lower() or 'ten' in c.lower()), None)
            if col_kho:
                coso_options = []
                for _, row in df_cs.iterrows():
                    m_kho = str(row[col_kho]).strip()
                    t_kho = str(row[col_ten]).strip() if col_ten and pd.notna(row[col_ten]) else ''
                    full_str = f"{m_kho} - {t_kho}"
                    coso_options.append(full_str)
        except:
            pass

    gl_accounts = get_list_accounts_from_gl(path_gl)
    coso_accounts = get_list_accounts_from_coso(path_coso)

    # --- KHỞI TẠO TRẠNG THÁI SESSION_STATE ---
    if 'bind_cccd' not in st.session_state: st.session_state['bind_cccd'] = ""
    if 'bind_ten' not in st.session_state: st.session_state['bind_ten'] = ""
    if 'bind_dc' not in st.session_state: st.session_state['bind_dc'] = ""
    if 'is_nhap_moi_active' not in st.session_state: st.session_state['is_nhap_moi_active'] = False
    if 'last_hoan_ung_auto_key' not in st.session_state: st.session_state['last_hoan_ung_auto_key'] = ""
    
    if 'persistent_msg' not in st.session_state: st.session_state['persistent_msg'] = None
    if 'persistent_type' not in st.session_state: st.session_state['persistent_type'] = None

    for mg in DENOMINATIONS:
        if f"to_{mg}" not in st.session_state or isinstance(st.session_state[f"to_{mg}"], (int, float)):
            st.session_state[f"to_{mg}"] = f"{_parse_int_formatted(st.session_state.get(f'to_{mg}', 0)):,}"
        if f"bo_{mg}" not in st.session_state: st.session_state[f"bo_{mg}"] = 0.0
        if f"bao_{mg}" not in st.session_state: st.session_state[f"bao_{mg}"] = 0.0

    # Header trực quan đồng bộ desktop & mobile
    st.markdown(f"""
        <div class="app-header-banner">
            <div class="app-header-title">🏦 QUẢN LÝ QUỸ TIỀN MỚI & HẠCH TOÁN KHO</div>
            <div class="app-header-meta">
                <span>👤 Cán bộ: <strong>{current_user}</strong></span>
                <span class="role-pill">🏷️ {user_role}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)

    tabs_list = [
        "📥 Giao Dịch Thu/Chi", "🔄 Ứng / Hoàn Tiền", "🛡️ Phê Duyệt (LĐ)", "🔍 Tìm Kiếm & Xử Lý", "📊 Báo Cáo"
    ]
    default_tab_idx = 2 if (is_giam_doc or is_ld_hub) else 0
    
    current_tab = st.radio(
        "Chuyển Tab Chức Năng",
        tabs_list,
        index=default_tab_idx,
        horizontal=True,
        label_visibility="collapsed"
    )
    st.markdown("---")

    if st.session_state['persistent_msg']:
        p_type = st.session_state['persistent_type']
        p_msg = st.session_state['persistent_msg']
        if p_type == 'success':
            st.success(p_msg)
        elif p_type == 'error':
            st.error(p_msg)
        elif p_type == 'warning':
            st.warning(p_msg)
        
        if st.button("✖️ Đóng thông báo này", key="btn_close_persistent_msg"):
            st.session_state['persistent_msg'] = None
            st.session_state['persistent_type'] = None
            st.rerun()
        st.markdown("---")

    if current_tab == "📥 Giao Dịch Thu/Chi":
        if is_giam_doc or is_ld_hub or role_upper in ["LĐTT", "ADMIN"]:
            st.warning("⚠️ Tài khoản nhóm quản lý cấp cao / Lãnh đạo không có quyền lập hoặc nhập liệu giao dịch.")
        else:
            col_title, col_btn_nhapmoi = st.columns([3, 1])
            with col_title:
                st.markdown("#### 📝 Lập Giao Dịch Thu / Chi Tiền Mới")
            with col_btn_nhapmoi:
                if not st.session_state['is_nhap_moi_active']:
                    if st.button("➕ Nhập Mới", type="primary", use_container_width=True):
                        st.session_state['is_nhap_moi_active'] = True
                        st.session_state['bind_cccd'] = ""
                        st.session_state['bind_ten'] = ""
                        st.session_state['bind_dc'] = ""
                        st.session_state['last_hoan_ung_auto_key'] = ""
                        st.session_state['persistent_msg'] = None
                        if 'gd_select_kh_chung' in st.session_state:
                            st.session_state['gd_select_kh_chung'] = "-- Chọn khách hàng từ danh sách hoặc nhập mới --"
                        for mg in DENOMINATIONS:
                            st.session_state[f"to_{mg}"] = "0"
                            st.session_state[f"bo_{mg}"] = 0.0
                            st.session_state[f"bao_{mg}"] = 0.0
                        st.rerun()
                else:
                    if st.button("🔄 Làm Mới Form", type="secondary", use_container_width=True):
                        st.session_state['bind_cccd'] = ""
                        st.session_state['bind_ten'] = ""
                        st.session_state['bind_dc'] = ""
                        st.session_state['last_hoan_ung_auto_key'] = ""
                        st.session_state['persistent_msg'] = "✨ Đã làm mới màn hình nhập liệu thành công!"
                        st.session_state['persistent_type'] = "success"
                        if 'gd_select_kh_chung' in st.session_state:
                            st.session_state['gd_select_kh_chung'] = "-- Chọn khách hàng từ danh sách hoặc nhập mới --"
                        for mg in DENOMINATIONS:
                            st.session_state[f"to_{mg}"] = "0"
                            st.session_state[f"bo_{mg}"] = 0.0
                            st.session_state[f"bao_{mg}"] = 0.0
                        st.rerun()

            if not st.session_state['is_nhap_moi_active']:
                st.info("ℹ️ Vui lòng bấm vào nút **'➕ Nhập Mới'** ở góc trên bên phải để bắt đầu tạo giao dịch.")
            else:
                with st.container(border=True):
                    st.markdown("##### ⚙️ Thông Tin Cơ Bản Giao Dịch")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        ngay_hieu_luc = st.date_input("📅 Ngày Hiệu Lực", value=date.today(), key="gd_ngay")
                        loai_nv_options = ["Cá Nhân", "Chi nhánh", "NHNN", "Nội bộ", "Tạm ứng", "Hoàn ứng", "Điều chuyển CN", "Điều chỉnh kiểm kê"]
                        loai_nv = st.selectbox("📌 Loại Nghiệp Vụ", loai_nv_options, index=0, key="gd_loainv")
                    with col2:
                        chon_coso = st.selectbox("🏢 Đơn Vị Xử Lý (Kho)", coso_options, key="gd_donvixl")
                        don_vi_xl = chon_coso.split(" - ")[0].strip()
                        
                        if loai_nv in ["Tạm ứng", "Cá Nhân", "Nội bộ"]:
                            thu_chi = "CHI"
                            st.selectbox("🔀 Loại Thu / Chi", ["CHI"], disabled=True, key="gd_thuchi_locked")
                        elif loai_nv == "Hoàn ứng":
                            thu_chi = "THU"
                            st.selectbox("🔀 Loại Thu / Chi", ["THU"], disabled=True, key="gd_thuchi_locked")
                        else:
                            thu_chi = st.selectbox("🔀 Loại Thu / Chi", ["THU", "CHI"], key="gd_thuchi")
                    with col3:
                        if loai_nv == "Nội bộ":
                            if coso_accounts:
                                sel_acc_cs = st.selectbox("🔍 Chọn Đơn Vị Nhận", ["-- Chọn đơn vị --"] + coso_accounts, key="gd_sel_acc_coso")
                                if sel_acc_cs != "-- Chọn đơn vị --":
                                    ma_kh = sel_acc_cs.split(" - ")[0].strip()
                                    default_ten_kh = " - ".join(sel_acc_cs.split(" - ")[1:])
                                else:
                                    ma_kh = st.text_input("Mã Đơn Vị Nhận", key="gd_makh_coso_free")
                                    default_ten_kh = ""
                            else:
                                ma_kh = st.text_input("Mã Đơn Vị Nhận", key="gd_makh_coso_free")
                                default_ten_kh = ""
                            diachi_kh = st.text_input("📍 Địa Chỉ (Không bắt buộc)", key="gd_diachi_coso")
                            ten_kh = st.text_input("🏷️ Tên Đơn Vị Nhận", value=default_ten_kh.upper(), max_chars=30, key="gd_tenkh_coso")
                            
                        elif loai_nv in ["Chi nhánh", "Hoàn ứng", "Tạm ứng", "Điều chuyển CN"]:
                            if gl_accounts:
                                sel_acc = st.selectbox("🔍 Chọn Mã Đơn Vị / KH từ Sổ Cái", ["-- Nhập tự do hoặc chọn --"] + gl_accounts, key="gd_sel_acc")
                                if sel_acc != "-- Nhập tự do hoặc chọn --":
                                    ma_kh = sel_acc.split(" - ")[0].strip()
                                    default_ten_kh = " - ".join(sel_acc.split(" - ")[1:])
                                else:
                                    ma_kh = st.text_input("Mã KH / Mã Đơn Vị", key="gd_makh_free")
                                    default_ten_kh = ""
                            else:
                                ma_kh = st.text_input("Mã KH / Mã Đơn Vị", key="gd_makh_free")
                                default_ten_kh = ""
                            diachi_kh = st.text_input("📍 Địa Chỉ KH (Không bắt buộc)", key="gd_diachi_other")
                            ten_kh = st.text_input("🏷️ Tên Khách Hàng / Đơn Vị", value=default_ten_kh.upper(), max_chars=30, key="gd_tenkh")
                        else:
                            df_kh_all = pd.DataFrame(columns=['MA CCCD', 'TEN KH', 'DIACHI KH'])
                            if os.path.exists(path_kh):
                                try:
                                    df_kh_all = pd.read_excel(path_kh)
                                    df_kh_all.columns = [str(c).strip().upper() for c in df_kh_all.columns]
                                except:
                                    pass
                                    
                            kh_options_dict = {"-- Chọn khách hàng từ danh sách hoặc nhập mới --": {"cccd": "", "ten": "", "dc": ""}}
                            if not df_kh_all.empty:
                                for _, row in df_kh_all.iterrows():
                                    m_cccd = str(row['MA CCCD']).strip() if 'MA CCCD' in df_kh_all.columns and pd.notna(row['MA CCCD']) else ''
                                    t_kh = str(row['TEN KH']).strip() if 'TEN KH' in df_kh_all.columns and pd.notna(row['TEN KH']) else ''
                                    d_kh = str(row['DIACHI KH']).strip() if 'DIACHI KH' in df_kh_all.columns and pd.notna(row['DIACHI KH']) else ''
                                    if m_cccd.lower() != 'nan' and m_cccd or t_kh.lower() != 'nan' and t_kh:
                                        label = f"{t_kh} - CCCD: {m_cccd} ({d_kh})"
                                        kh_options_dict[label] = {"cccd": m_cccd, "ten": t_kh, "dc": d_kh}

                            def on_change_chon_kh_v2():
                                sel_label = st.session_state.get('gd_select_kh_chung', "-- Chọn khách hàng từ danh sách hoặc nhập mới --")
                                info = kh_options_dict.get(sel_label, {"cccd": "", "ten": "", "dc": ""})
                                st.session_state['bind_cccd'] = info["cccd"]
                                st.session_state['bind_ten'] = info["ten"].upper()
                                st.session_state['bind_dc'] = info["dc"]

                            st.selectbox(
                                "🔍 Tìm và Chọn Khách Hàng",
                                list(kh_options_dict.keys()),
                                key="gd_select_kh_chung",
                                on_change=on_change_chon_kh_v2
                            )

                            ma_kh = st.text_input("🆔 Mã CCCD / Mã KH", key="bind_cccd")
                            ten_kh = st.text_input("🏷️ Tên Khách Hàng (In hoa)", max_chars=30, key="bind_ten")
                            diachi_kh = st.text_input("📍 Địa Chỉ KH", key="bind_dc")

                with st.container(border=True):
                    col_ccmg_title, col_ccmg_btn = st.columns([3, 1])
                    with col_ccmg_title:
                        st.markdown("##### 🔢 Cơ Cấu Mệnh Giá Chi Tiết")
                    unit_mode = st.radio("Đơn vị tính nhập liệu:", ["Số tờ (Tự do, ngăn hàng ngàn)", "Theo Bó (1 Bó = 1,000 tờ)", "Theo Bao (1 Bao = 20,000 tờ)"], horizontal=True, key="gd_unit_mode")

                    if loai_nv == "Hoàn ứng" and str(ma_kh).strip():
                        auto_key = f"{str(ma_kh).strip().lower()}|{unit_mode}"
                        if st.session_state.get('last_hoan_ung_auto_key') != auto_key:
                            gl_to_map = get_donvi_denominations_from_gl(path_gl, str(ma_kh).strip())
                            apply_so_to_to_session_inputs(gl_to_map, unit_mode)
                            st.session_state['last_hoan_ung_auto_key'] = auto_key

                    with col_ccmg_btn:
                        if loai_nv == "Hoàn ứng" and str(ma_kh).strip():
                            if st.button("📥 Lấy số tờ từ GL (DON VI)", key="btn_reload_gl_hoan_ung", use_container_width=True):
                                gl_to_map = get_donvi_denominations_from_gl(path_gl, str(ma_kh).strip())
                                apply_so_to_to_session_inputs(gl_to_map, unit_mode)
                                st.session_state['last_hoan_ung_auto_key'] = f"{str(ma_kh).strip().lower()}|{unit_mode}"
                                st.rerun()

                    if loai_nv == "Hoàn ứng" and str(ma_kh).strip():
                        st.caption(f"💡 Đã tự động tải số tờ từ khối **DON VI** của mã `{str(ma_kh).strip()}` trên `gl_tienmo.xlsx`.")
                    
                    mg_cols = st.columns(5)
                    ccmg_data = {}
                    tong_thanh_tien = 0

                    def make_callback_nhap_moi_to(denomination):
                        def cb():
                            num_val = _parse_int_formatted(st.session_state.get(f"to_{denomination}", "0"))
                            st.session_state[f"to_{denomination}"] = f"{num_val:,}"
                        return cb
                    
                    for idx, mg in enumerate(DENOMINATIONS):
                        with mg_cols[idx % 5]:
                            if unit_mode.startswith("Số tờ"):
                                key_to = f"to_{mg}"
                                st.session_state[key_to] = f"{_parse_int_formatted(st.session_state.get(key_to, '0')):,}"
                                st.text_input(
                                    f"{mg:,} đ",
                                    key=key_to,
                                    on_change=make_callback_nhap_moi_to(mg)
                                )
                                so_to = _parse_int_formatted(st.session_state.get(key_to, "0"))
                            elif unit_mode.startswith("Theo Bó"):
                                so_bo = st.number_input(f"{mg:,} (Bó)", min_value=0.0, step=1.0, format="%.2f", key=f"bo_{mg}")
                                so_to = int(so_bo * 1000)
                            else:
                                so_bao = st.number_input(f"{mg:,} (Bao)", min_value=0.0, step=0.5, format="%.2f", key=f"bao_{mg}")
                                so_to = int(so_bao * 20000)
                                
                            ccmg_data[mg] = so_to
                            tong_thanh_tien += so_to * mg

                    st.markdown("---")
                    st.metric(label="💰 Tổng Giá Trị Giao Dịch Khởi Tạo", value=f"{tong_thanh_tien:,.0f} VNĐ")
                
                if st.button("💾 Lưu và Gửi Bản Ghi (Chờ Phê Duyệt)", type="primary", use_container_width=True):
                    df_ds_existing = pd.read_excel(path_ds) if os.path.exists(path_ds) else pd.DataFrame()
                    effective_ma_kh = ma_kh.strip() if ma_kh.strip() else don_vi_xl

                    ok_chi, err_chi = validate_chi_kha_dung(
                        path_gl,
                        thu_chi_val=thu_chi,
                        loai_nv_val=loai_nv,
                        don_vi_xl_val=don_vi_xl,
                        ma_kh_val=effective_ma_kh,
                        ccmg_data=ccmg_data,
                        df_all_tx=df_ds_existing,
                        exclude_seq=None
                    )

                    if not ma_kh.strip() and loai_nv != "Điều chỉnh kiểm kê":
                        st.session_state['persistent_msg'] = "❌ Mã CCCD / Mã KH / Mã Đơn Vị nhận không được để trống!"
                        st.session_state['persistent_type'] = "error"
                        st.rerun()
                    elif not ten_kh.strip():
                        st.session_state['persistent_msg'] = "❌ Tên khách hàng / Đơn vị nhận không được để trống!"
                        st.session_state['persistent_type'] = "error"
                        st.rerun()
                    elif tong_thanh_tien <= 0:
                        st.session_state['persistent_msg'] = "❌ Vui lòng nhập số lượng tiền giao dịch lớn hơn 0!"
                        st.session_state['persistent_type'] = "error"
                        st.rerun()
                    elif not ok_chi:
                        st.session_state['persistent_msg'] = f"{err_chi} — Chặn lưu đẩy duyệt, vui lòng điều chỉnh lại số tờ theo tồn khả dụng!"
                        st.session_state['persistent_type'] = "error"
                        st.rerun()
                    else:
                        try:
                            if loai_nv == "Cá Nhân" and ma_kh.strip():
                                if os.path.exists(path_kh):
                                    df_kh_file = pd.read_excel(path_kh)
                                    df_kh_file.columns = [str(c).strip().upper() for c in df_kh_file.columns]
                                else:
                                    df_kh_file = pd.DataFrame(columns=['MA CCCD', 'TEN KH', 'DIACHI KH'])
                                    
                                for col in ['MA CCCD', 'TEN KH', 'DIACHI KH']:
                                    if col not in df_kh_file.columns:
                                        df_kh_file[col] = ''
                                
                                m_cccd_str = str(ma_kh).strip()
                                exist_mask = df_kh_file['MA CCCD'].astype(str).str.strip() == m_cccd_str
                                
                                if not exist_mask.any():
                                    new_kh_row = pd.DataFrame([{'MA CCCD': m_cccd_str, 'TEN KH': ten_kh.strip().upper(), 'DIACHI KH': diachi_kh.strip()}])
                                    df_kh_file = pd.concat([df_kh_file, new_kh_row], ignore_index=True)
                                    df_kh_file.to_excel(path_kh, index=False)

                            if os.path.exists(path_ds):
                                df_ds = pd.read_excel(path_ds)
                            else:
                                cols = ['MA KH CCCD', 'TEN KH', 'DIACHI KH', 'DON VI XL', 'LOAI NV', 'SO SEQ', 'NGAY HIEU LUC', 'THU CHI', 'TRANG THAI']
                                for d in DENOMINATIONS:
                                    cols.extend([str(d), f"{d}_THANHTIEN"])
                                cols.extend(['TONG THANH TIEN', 'GDV TM', 'NGUOI XL', 'NGUOI PHE DUYET'])
                                df_ds = pd.DataFrame(columns=cols)
                            
                            if 'SO SEQ' in df_ds.columns and not df_ds.empty:
                                valid_seqs = pd.to_numeric(df_ds['SO SEQ'], errors='coerce')
                                next_seq = int(valid_seqs.max() + 1) if not valid_seqs.isna().all() else 1
                            else:
                                next_seq = 1
                                
                            new_row = {
                                'MA KH CCCD': effective_ma_kh,
                                'TEN KH': ten_kh.strip().upper(),
                                'DIACHI KH': diachi_kh.strip(),
                                'DON VI XL': don_vi_xl,
                                'LOAI NV': loai_nv,
                                'SO SEQ': float(next_seq),
                                'NGAY HIEU LUC': ngay_hieu_luc.strftime('%Y-%m-%d'),
                                'THU CHI': thu_chi,
                                'TRANG THAI': 'Chờ Phê Duyệt',
                                'GDV TM': str(current_user),
                                'NGUOI XL': '',
                                'NGUOI PHE DUYET': '',
                                'TONG THANH TIEN': float(tong_thanh_tien)
                            }
                            
                            for idx_mg, mg in enumerate(DENOMINATIONS):
                                so_to_val = float(ccmg_data[mg])
                                thanh_tien_val = so_to_val * mg
                                
                                if mg in df_ds.columns:
                                    new_row[mg] = so_to_val
                                if str(mg) in df_ds.columns:
                                    new_row[str(mg)] = so_to_val
                                    
                                if idx_mg == 0 and 'THANH TIEN' in df_ds.columns:
                                    new_row['THANH TIEN'] = thanh_tien_val
                                else:
                                    tt_col_name = f"THANH TIEN.{idx_mg}"
                                    if tt_col_name in df_ds.columns:
                                        new_row[tt_col_name] = thanh_tien_val
                                    alt_tt_name = f"{mg}_THANHTIEN"
                                    if alt_tt_name in df_ds.columns:
                                        new_row[alt_tt_name] = thanh_tien_val

                            final_row = {}
                            for col in df_ds.columns:
                                val_item = new_row.get(col, None)
                                is_numeric_col = (col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ')
                                
                                if val_item is None or str(val_item).strip() == '' or str(val_item).lower() == 'nan':
                                    final_row[col] = 0.0 if is_numeric_col else ''
                                else:
                                    if is_numeric_col:
                                        try:
                                            final_row[col] = float(val_item)
                                        except:
                                            final_row[col] = 0.0
                                    else:
                                        final_row[col] = str(val_item)

                            df_ds = pd.concat([df_ds, pd.DataFrame([final_row])], ignore_index=True)
                            
                            for col in df_ds.columns:
                                if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                    df_ds[col] = pd.to_numeric(df_ds[col], errors='coerce').fillna(0.0)

                            df_ds.to_excel(path_ds, index=False)
                            
                            st.session_state['is_nhap_moi_active'] = False
                            st.session_state['last_hoan_ung_auto_key'] = ""
                            st.session_state['persistent_msg'] = f"🎉 Lập bản ghi giao dịch thành công (Mã SEQ: {next_seq})!"
                            st.session_state['persistent_type'] = "success"
                            st.rerun()
                        except Exception as e:
                            st.session_state['persistent_msg'] = f"❌ Lỗi lưu giao dịch: {e}"
                            st.session_state['persistent_type'] = "error"
                            st.rerun()

    elif current_tab == "🔄 Ứng / Hoàn Tiền":
        st.markdown("#### 🔄 Giao Dịch Tạm Ứng / Hoàn Ứng Nhanh")
        st.info("Chức năng đồng bộ tương tự tab Giao dịch với nghiệp vụ Tạm ứng hoặc Hoàn ứng tự động.")

    elif current_tab == "🛡️ Phê Duyệt (LĐ)":
        st.markdown("#### 🛡️ Phê Duyệt Giao Dịch & Sửa Cơ Cấu Mệnh Giá")
        st.caption("Giao diện tối ưu đồng nhất cho cả **Máy tính** và **Điện thoại / Máy tính bảng**.")
        
        if not is_giam_doc and not is_ld_hub:
            st.warning("⚠️ Tài khoản của bạn không có thẩm quyền phê duyệt. Chức năng này chỉ dành riêng cho nhóm quyền **GIÁM ĐỐC** và **LĐ_HUB**.")
        else:
            if os.path.exists(path_ds):
                df_app = pd.read_excel(path_ds)
                if 'TRANG THAI' not in df_app.columns:
                    df_app['TRANG THAI'] = 'Chờ Phê Duyệt'
                    
                df_cho = df_app[df_app['TRANG THAI'].astype(str).str.strip() == 'Chờ Phê Duyệt'].copy()
                
                allowed_codes = []
                
                if is_giam_doc:
                    df_cho_filtered = df_cho
                    allowed_codes = ["Toàn bộ kho (Quyền Giám Đốc)"]
                elif is_ld_hub:
                    if not assigned_coso_ld_hub:
                        st.error(f"❌ Không tìm thấy giá trị `CoSo` trong file `Thamso_canbo.xlsx` cho tài khoản '{current_user}'!")
                        df_cho_filtered = pd.DataFrame(columns=df_cho.columns)
                    else:
                        c_lower = str(assigned_coso_ld_hub).strip().lower()
                        
                        if 'cơ sở 2' in c_lower or 'co so 2' in c_lower:
                            allowed_codes = ['990.3', '990.3.gdv']
                        elif 'cơ sở 1' in c_lower or 'co so 1' in c_lower:
                            allowed_codes = ['990.1', '990.1.gdv']
                        elif 'hđ' in c_lower or 'hd' in c_lower:
                            allowed_codes = ['990.6', '990.6.gdv']
                        elif 'hp' in c_lower:
                            allowed_codes = ['990.5', '990.5.gdv']
                        elif 'nvc' in c_lower:
                            allowed_codes = ['990.7', '990.7.gdv']
                        else:
                            base_m = c_lower.split(" - ")[0].strip()
                            allowed_codes = [base_m, f"{base_m}.gdv"]

                        def match_mapped_hub(val):
                            v_str = str(val).strip().lower()
                            if v_str in ['990.1', '990.1.gdv']:
                                return False
                            return v_str in allowed_codes

                        mask_hub = df_cho['DON VI XL'].apply(match_mapped_hub)
                        df_cho_filtered = df_cho[mask_hub]
                else:
                    df_cho_filtered = pd.DataFrame(columns=df_cho.columns)

                if df_cho_filtered.empty:
                    st.info(f"ℹ️ Hiện không có bản ghi nào chờ phê duyệt thuộc thẩm quyền của bạn (Cơ sở đọc từ file: `{assigned_coso_ld_hub if is_ld_hub else 'GIÁM ĐỐC'}`).")
                else:
                    total_cho_val = pd.to_numeric(df_cho_filtered.get('TONG THANH TIEN', 0), errors='coerce').fillna(0).sum()
                    m_col1, m_col2, m_col3 = st.columns(3)
                    with m_col1:
                        st.metric("📌 Bản Ghi Chờ Duyệt", f"{len(df_cho_filtered):,}")
                    with m_col2:
                        st.metric("💰 Tổng Giá Trị Chờ Duyệt", f"{total_cho_val:,.0f} đ")
                    with m_col3:
                        st.metric("🏢 Phạm Vi Quản Lý", f"{assigned_coso_ld_hub if is_ld_hub else 'Toàn hệ thống'}")

                    with st.container(border=True):
                        f_app1, f_app2 = st.columns([2, 1])
                        with f_app1:
                            kw_app = st.text_input("🔎 Lọc nhanh theo Mã SEQ / Tên KH / Đơn vị XL", key="filter_app_kw", placeholder="Nhập từ khóa...")
                        with f_app2:
                            tc_opts_app = ["TẤT CẢ"] + sorted(list(df_cho_filtered.get('THU CHI', pd.Series(dtype=str)).dropna().astype(str).unique()))
                            tc_filter_app = st.selectbox("🔀 Lọc Thu / Chi", tc_opts_app, key="filter_app_tc")

                    df_cho_view_src = df_cho_filtered.copy()
                    if kw_app.strip():
                        k_low = kw_app.strip().lower()
                        mask_kw = (
                            df_cho_view_src['SO SEQ'].astype(str).str.lower().str.contains(k_low, na=False) |
                            df_cho_view_src['TEN KH'].astype(str).str.lower().str.contains(k_low, na=False) |
                            df_cho_view_src['DON VI XL'].astype(str).str.lower().str.contains(k_low, na=False)
                        )
                        df_cho_view_src = df_cho_view_src[mask_kw]
                    if tc_filter_app != "TẤT CẢ" and 'THU CHI' in df_cho_view_src.columns:
                        df_cho_view_src = df_cho_view_src[df_cho_view_src['THU CHI'].astype(str) == tc_filter_app]

                    st.markdown(f"##### 📋 Danh sách chờ duyệt ({len(df_cho_view_src)} giao dịch):")
                    
                    cols_to_show_app = ['SO SEQ', 'TEN KH', 'LOAI NV', 'THU CHI', 'DON VI XL', 'TRANG THAI']
                    for mg in DENOMINATIONS:
                        if mg in df_cho_view_src.columns:
                            cols_to_show_app.append(mg)
                        elif str(mg) in df_cho_view_src.columns:
                            cols_to_show_app.append(str(mg))
                    if 'TONG THANH TIEN' in df_cho_view_src.columns:
                        cols_to_show_app.append('TONG THANH TIEN')
                        
                    valid_view_cols_app = [c for c in cols_to_show_app if c in df_cho_view_src.columns]
                    df_cho_display = df_cho_view_src[valid_view_cols_app].copy()
                    
                    for col in df_cho_display.columns:
                        try:
                            val_int_check = int(float(col))
                            if val_int_check in DENOMINATIONS or col == 'TONG THANH TIEN':
                                df_cho_display[col] = pd.to_numeric(df_cho_display[col], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,}")
                        except:
                            if col == 'TONG THANH TIEN':
                                df_cho_display[col] = pd.to_numeric(df_cho_display[col], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,}")

                    rename_map_app = {}
                    for col in df_cho_display.columns:
                        try:
                            val_int = int(float(col))
                            if val_int in DENOMINATIONS:
                                rename_map_app[col] = f"{val_int:,}"
                        except:
                            pass
                    df_cho_display.rename(columns=rename_map_app, inplace=True)

                    df_cho_display.insert(0, "Chọn", False)

                    edited_list_app = st.data_editor(
                        df_cho_display,
                        hide_index=True,
                        use_container_width=True,
                        key="editor_pheduyet_list_v103"
                    )
                    
                    selected_rows_app = edited_list_app[edited_list_app['Chọn'] == True]
                    selected_seqs_app = selected_rows_app['SO SEQ'].tolist()

                    if selected_seqs_app:
                        sel_src_df = df_cho_view_src[df_cho_view_src['SO SEQ'].isin(selected_seqs_app)]
                        sel_sum_amt = pd.to_numeric(sel_src_df.get('TONG THANH TIEN', 0), errors='coerce').fillna(0).sum()
                        st.markdown(f"""
                            <div class="mobile-selection-bar">
                                ☑️ Đã chọn <strong>{len(selected_seqs_app)}</strong> giao dịch | Tổng giá trị đã chọn: <strong>{sel_sum_amt:,.0f} VNĐ</strong>
                            </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.caption("💡 Tích chọn vào ô **Chọn** ở dòng giao dịch cần thao tác, sau đó bấm nút bên dưới.")
                    
                    st.markdown("---")
                    col_b1, col_b2, col_b3, col_b4 = st.columns(4)
                    with col_b1:
                        if st.button("✅ Phê Duyệt Nhanh", type="primary", key="btn_ld_duyet_nhanh", use_container_width=True):
                            if not selected_seqs_app:
                                st.warning("⚠️ Vui lòng tích chọn ít nhất một giao dịch!")
                            else:
                                try:
                                    blocked_msgs = []
                                    valid_approve_seqs = []
                                    for s_item in selected_seqs_app:
                                        r_match = df_app[df_app['SO SEQ'] == s_item]
                                        if not r_match.empty:
                                            r_one = r_match.iloc[0]
                                            ccmg_one = _extract_row_ccmg(r_one)
                                            ok_v, msg_v = validate_chi_kha_dung(
                                                path_gl,
                                                thu_chi_val=r_one.get('THU CHI', ''),
                                                loai_nv_val=r_one.get('LOAI NV', ''),
                                                don_vi_xl_val=r_one.get('DON VI XL', ''),
                                                ma_kh_val=r_one.get('MA KH CCCD', ''),
                                                ccmg_data=ccmg_one,
                                                df_all_tx=df_app,
                                                exclude_seq=s_item
                                            )
                                            if ok_v:
                                                valid_approve_seqs.append(s_item)
                                            else:
                                                blocked_msgs.append(f"SEQ {s_item}: {msg_v}")

                                    if valid_approve_seqs:
                                        df_app['TRANG THAI'] = df_app['TRANG THAI'].astype(str)
                                        df_app['NGUOI PHE DUYET'] = df_app['NGUOI PHE DUYET'].astype(str)
                                        df_app.loc[df_app['SO SEQ'].isin(valid_approve_seqs), 'TRANG THAI'] = 'Duyệt'
                                        df_app.loc[df_app['SO SEQ'].isin(valid_approve_seqs), 'NGUOI PHE DUYET'] = str(current_user)
                                        
                                        for col in df_app.columns:
                                            if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                                df_app[col] = pd.to_numeric(df_app[col], errors='coerce').fillna(0.0)

                                        df_app.to_excel(path_ds, index=False)

                                    if blocked_msgs:
                                        prefix = f"Đã duyệt {len(valid_approve_seqs)} GD. " if valid_approve_seqs else ""
                                        st.session_state['persistent_msg'] = prefix + "\n\n".join(blocked_msgs)
                                        st.session_state['persistent_type'] = "error"
                                    else:
                                        st.session_state['persistent_msg'] = f"🎉 Đã phê duyệt thành công {len(valid_approve_seqs)} giao dịch!"
                                        st.session_state['persistent_type'] = "success"
                                    st.rerun()
                                except Exception as ex:
                                    st.session_state['persistent_msg'] = f"❌ Lỗi phê duyệt: {ex}"
                                    st.session_state['persistent_type'] = "error"
                                    st.rerun()
                    with col_b2:
                        if st.button("✏️ Sửa Cơ Cấu", key="btn_ld_mo_sua_ccmg", use_container_width=True):
                            if len(selected_seqs_app) != 1:
                                st.warning("⚠️ Vui lòng tích chọn **chính xác 1 dòng duy nhất** để mở form sửa!")
                            else:
                                target_s = selected_seqs_app[0]
                                st.session_state['sua_seq_dang_chon_ld'] = target_s
                                r_init = df_cho_filtered[df_cho_filtered['SO SEQ'] == target_s].iloc[0]
                                for mg in DENOMINATIONS:
                                    val_mg = 0
                                    if mg in r_init and pd.notna(r_init[mg]):
                                        val_mg = int(r_init[mg])
                                    elif str(mg) in r_init and pd.notna(r_init[str(mg)]):
                                        val_mg = int(r_init[str(mg)])
                                    st.session_state[f"input_ld_{target_s}_{mg}"] = f"{val_mg:,}"
                                st.rerun()
                    with col_b3:
                        if st.button("🗑️ Xóa Bản Ghi", key="btn_ld_xoa", use_container_width=True):
                            if not selected_seqs_app:
                                st.warning("⚠️ Vui lòng tích chọn bản ghi cần xóa!")
                            else:
                                try:
                                    df_app.loc[df_app['SO SEQ'].isin(selected_seqs_app), 'TRANG THAI'] = 'Xóa'
                                    for col in df_app.columns:
                                        if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                            df_app[col] = pd.to_numeric(df_app[col], errors='coerce').fillna(0.0)
                                    df_app.to_excel(path_ds, index=False)
                                    st.session_state['persistent_msg'] = f"⚠️ Đã xóa {len(selected_seqs_app)} bản ghi."
                                    st.session_state['persistent_type'] = "warning"
                                    st.rerun()
                                except Exception as ex:
                                    st.session_state['persistent_msg'] = f"❌ Lỗi xóa: {ex}"
                                    st.session_state['persistent_type'] = "error"
                                    st.rerun()
                    with col_b4:
                        if st.button("❌ Từ Chối", key="btn_ld_tuchoi", use_container_width=True):
                            if not selected_seqs_app:
                                st.warning("⚠️ Vui lòng tích chọn ít nhất một giao dịch!")
                            else:
                                try:
                                    df_app['TRANG THAI'] = df_app['TRANG THAI'].astype(str)
                                    df_app['NGUOI PHE DUYET'] = df_app['NGUOI PHE DUYET'].astype(str)
                                    df_app.loc[df_app['SO SEQ'].isin(selected_seqs_app), 'TRANG THAI'] = 'Từ Chối Phê Duyệt'
                                    df_app.loc[df_app['SO SEQ'].isin(selected_seqs_app), 'NGUOI PHE DUYET'] = str(current_user)
                                    for col in df_app.columns:
                                        if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                            df_app[col] = pd.to_numeric(df_app[col], errors='coerce').fillna(0.0)
                                    df_app.to_excel(path_ds, index=False)
                                    st.session_state['persistent_msg'] = f"⚠️ Đã từ chối {len(selected_seqs_app)} giao dịch."
                                    st.session_state['persistent_type'] = "warning"
                                    st.rerun()
                                except Exception as ex:
                                    st.session_state['persistent_msg'] = f"❌ Lỗi từ chối: {ex}"
                                    st.session_state['persistent_type'] = "error"
                                    st.rerun()

                    if 'sua_seq_dang_chon_ld' in st.session_state and st.session_state['sua_seq_dang_chon_ld'] is not None:
                        target_seq = st.session_state['sua_seq_dang_chon_ld']
                        row_target_df = df_app[df_app['SO SEQ'] == target_seq]
                        
                        if not row_target_df.empty:
                            r_val = row_target_df.iloc[0]
                            with st.container(border=True):
                                st.markdown(f"#### ✍️ [LÃNH ĐẠO] Điều Chỉnh Cơ Cấu Mệnh Giá")
                                st.caption(f"Mã SEQ: **`{target_seq}`** | KH: **`{r_val['TEN KH']}`** | Thu/Chi: **`{r_val.get('THU CHI', '')}`** | Kho XL: **`{r_val.get('DON VI XL', '')}`**")
                                
                                edit_mg_cols = st.columns(5)
                                new_ccmg_edited = {}
                                new_total_val = 0
                                
                                def make_callback_ld(s_id, denomination):
                                    def cb():
                                        raw_text = st.session_state.get(f"input_ld_{s_id}_{denomination}", "0")
                                        try:
                                            clean_str = str(raw_text).replace(",", "").strip()
                                            num_val = int(float(clean_str)) if clean_str else 0
                                            st.session_state[f"input_ld_{s_id}_{denomination}"] = f"{num_val:,}"
                                        except:
                                            st.session_state[f"input_ld_{s_id}_{denomination}"] = "0"
                                    return cb

                                for idx, mg in enumerate(DENOMINATIONS):
                                    with edit_mg_cols[idx % 5]:
                                        key_name = f"input_ld_{target_seq}_{mg}"
                                        if key_name not in st.session_state:
                                            old_to = 0
                                            if mg in r_val and pd.notna(r_val[mg]):
                                                old_to = int(r_val[mg])
                                            elif str(mg) in r_val and pd.notna(r_val[str(mg)]):
                                                old_to = int(r_val[str(mg)])
                                            st.session_state[key_name] = f"{old_to:,}"

                                        st.text_input(
                                            f"{mg:,} đ", 
                                            key=key_name,
                                            on_change=make_callback_ld(target_seq, mg)
                                        )
                                        
                                        try:
                                            c_str = str(st.session_state[key_name]).replace(",", "").strip()
                                            val_sua = int(float(c_str)) if c_str else 0
                                        except:
                                            val_sua = 0
                                            
                                        new_ccmg_edited[mg] = val_sua
                                        new_total_val += val_sua * mg

                                st.metric(label="💰 Tổng Giá Trị Sau Khi Sửa", value=f"{new_total_val:,.0f} VNĐ")
                                
                                col_f1, col_f2 = st.columns(2)
                                with col_f1:
                                    if st.button("💾 Lưu Sửa Đổi & Phê Duyệt Ngay", type="primary", key="btn_ld_save_approve", use_container_width=True):
                                        ok_v_ld, msg_v_ld = validate_chi_kha_dung(
                                            path_gl,
                                            thu_chi_val=r_val.get('THU CHI', ''),
                                            loai_nv_val=r_val.get('LOAI NV', ''),
                                            don_vi_xl_val=r_val.get('DON VI XL', ''),
                                            ma_kh_val=r_val.get('MA KH CCCD', ''),
                                            ccmg_data=new_ccmg_edited,
                                            df_all_tx=df_app,
                                            exclude_seq=target_seq
                                        )
                                        if not ok_v_ld:
                                            st.session_state['persistent_msg'] = f"{msg_v_ld} — Không thể phê duyệt vượt số tờ tồn khả dụng!"
                                            st.session_state['persistent_type'] = "error"
                                            st.rerun()
                                        else:
                                            try:
                                                for mg in DENOMINATIONS:
                                                    so_to_moi = int(new_ccmg_edited[mg])
                                                    col_key = mg if mg in df_app.columns else (str(mg) if str(mg) in df_app.columns else None)
                                                    if col_key:
                                                        df_app.loc[df_app['SO SEQ'] == target_seq, col_key] = float(so_to_moi)
                                                    alt_t = f"{mg}_THANHTIEN"
                                                    if alt_t in df_app.columns:
                                                        df_app.loc[df_app['SO SEQ'] == target_seq, alt_t] = float(so_to_moi * mg)
                                                        
                                                df_app.loc[df_app['SO SEQ'] == target_seq, 'TONG THANH TIEN'] = float(new_total_val)
                                                df_app.loc[df_app['SO SEQ'] == target_seq, 'TRANG THAI'] = 'Duyệt'
                                                df_app.loc[df_app['SO SEQ'] == target_seq, 'NGUOI PHE DUYET'] = str(current_user)
                                                
                                                for col in df_app.columns:
                                                    if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                                        df_app[col] = pd.to_numeric(df_app[col], errors='coerce').fillna(0.0)

                                                df_app.to_excel(path_ds, index=False)
                                                del st.session_state['sua_seq_dang_chon_ld']
                                                st.session_state['persistent_msg'] = f"🎉 Đã cập nhật cơ cấu mệnh giá và phê duyệt thành công giao dịch SEQ {target_seq}!"
                                                st.session_state['persistent_type'] = "success"
                                                st.rerun()
                                            except Exception as ex:
                                                st.session_state['persistent_msg'] = f"❌ Lỗi: {ex}"
                                                st.session_state['persistent_type'] = "error"
                                                st.rerun()
                                with col_f2:
                                    if st.button("↩️ Hủy Bỏ", key="btn_ld_cancel_ccmg", use_container_width=True):
                                        del st.session_state['sua_seq_dang_chon_ld']
                                        st.rerun()

    elif current_tab == "🔍 Tìm Kiếm & Xử Lý":
        st.markdown("#### 🔍 Xử Lý Hạch Toán Kho (XL) & Quản Lý Bản Ghi (Nhân Viên)")
        
        if is_giam_doc or is_ld_hub:
            st.warning("⚠️ Tài khoản Lãnh đạo (Giám đốc / LĐ_HUB) tập trung thực hiện phê duyệt tại tab Phê Duyệt.")
        else:
            if os.path.exists(path_ds):
                df_mgr = pd.read_excel(path_ds)
                if 'TRANG THAI' not in df_mgr.columns:
                    df_mgr['TRANG THAI'] = 'Chờ Phê Duyệt'
                
                df_editable_list = df_mgr[df_mgr['TRANG THAI'].astype(str).str.strip().isin(['Chờ Phê Duyệt', 'Từ Chối Phê Duyệt', 'Duyệt', 'Xóa'])].copy()
                
                if df_editable_list.empty:
                    st.info("ℹ️ Hiện không có bản ghi giao dịch nào cần xử lý.")
                else:
                    st.markdown("##### ✏️ Quản Lý Xử Lý Hạch Toán Kho:")
                    
                    cols_to_show_base = ['SO SEQ', 'TEN KH', 'LOAI NV', 'THU CHI', 'DON VI XL', 'TRANG THAI']
                    for mg in DENOMINATIONS:
                        if mg in df_editable_list.columns:
                            cols_to_show_base.append(mg)
                        elif str(mg) in df_editable_list.columns:
                            cols_to_show_base.append(str(mg))
                    if 'TONG THANH TIEN' in df_editable_list.columns:
                        cols_to_show_base.append('TONG THANH TIEN')
                        
                    valid_view_cols_mgr = [c for c in cols_to_show_base if c in df_editable_list.columns]
                    df_mgr_display = df_editable_list[valid_view_cols_mgr].copy()
                    
                    for col in df_mgr_display.columns:
                        try:
                            val_int_check = int(float(col))
                            if val_int_check in DENOMINATIONS or col == 'TONG THANH TIEN':
                                df_mgr_display[col] = pd.to_numeric(df_mgr_display[col], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,}")
                        except:
                            if col == 'TONG THANH TIEN':
                                df_mgr_display[col] = pd.to_numeric(df_mgr_display[col], errors='coerce').fillna(0).apply(lambda x: f"{int(x):,}")

                    rename_map_mgr = {}
                    for col in df_mgr_display.columns:
                        try:
                            val_int = int(float(col))
                            if val_int in DENOMINATIONS:
                                rename_map_mgr[col] = f"{val_int:,}"
                        except:
                            pass
                    df_mgr_display.rename(columns=rename_map_mgr, inplace=True)

                    df_mgr_display.insert(0, "Chọn", False)

                    edited_df_mgr = st.data_editor(
                        df_mgr_display,
                        hide_index=True,
                        use_container_width=True,
                        key="editor_nhanvien_suaxoa_v103"
                    )
                    
                    selected_rows_mgr = edited_df_mgr[edited_df_mgr['Chọn'] == True]
                    selected_seqs_xl = selected_rows_mgr['SO SEQ'].tolist()
                    
                    st.markdown("---")
                    col_act1, col_act2, col_act3 = st.columns(3)
                    with col_act1:
                        if st.button("✏️ Sửa Cơ Cấu Mệnh Giá", key="btn_gdv_mo_sua_ccmg", use_container_width=True):
                            if len(selected_seqs_xl) != 1:
                                st.warning("⚠️ Vui lòng tích chọn **chính xác 1 dòng duy nhất** để mở form sửa!")
                            else:
                                target_seq_gdv = selected_seqs_xl[0]
                                r_check = df_mgr.loc[df_mgr['SO SEQ'] == target_seq_gdv, 'TRANG THAI'].values[0]
                                if str(r_check).strip() in ['Chờ Phê Duyệt', 'Từ Chối Phê Duyệt']:
                                    st.session_state['sua_seq_dang_chon_gdv'] = target_seq_gdv
                                    r_init_g = df_editable_list[df_editable_list['SO SEQ'] == target_seq_gdv].iloc[0]
                                    for mg in DENOMINATIONS:
                                        val_mg_g = 0
                                        if mg in r_init_g and pd.notna(r_init_g[mg]):
                                            val_mg_g = int(r_init_g[mg])
                                        elif str(mg) in r_init_g and pd.notna(r_init_g[str(mg)]):
                                            val_mg_g = int(r_init_g[str(mg)])
                                        st.session_state[f"input_gdv_{target_seq_gdv}_{mg}"] = f"{val_mg_g:,}"
                                    st.rerun()
                                else:
                                    st.error("❌ Bản ghi đã ở trạng thái 'Duyệt' hoặc 'XL', GDV không có quyền sửa!")
                    with col_act2:
                        if st.button("🗑️ Xóa Bản Ghi", key="btn_xoa_nv", use_container_width=True):
                            if not selected_seqs_xl:
                                st.warning("⚠️ Vui lòng tích chọn bản ghi cần xóa!")
                            else:
                                try:
                                    err_cnt = 0
                                    valid_seqs_to_delete = []
                                    for seq in selected_seqs_xl:
                                        r_curr_stt = df_mgr.loc[df_mgr['SO SEQ'] == seq, 'TRANG THAI'].values[0]
                                        if str(r_curr_stt).strip() in ['Chờ Phê Duyệt', 'Từ Chối Phê Duyệt']:
                                            valid_seqs_to_delete.append(seq)
                                        else:
                                            err_cnt += 1
                                            
                                    if valid_seqs_to_delete:
                                        df_mgr.loc[df_mgr['SO SEQ'].isin(valid_seqs_to_delete), 'TRANG THAI'] = 'Xóa'
                                        for col in df_mgr.columns:
                                            if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                                df_mgr[col] = pd.to_numeric(df_mgr[col], errors='coerce').fillna(0.0)
                                        df_mgr.to_excel(path_ds, index=False)
                                        
                                    if err_cnt > 0:
                                        st.warning(f"⚠️ Đã xóa các dòng hợp lệ. Có {err_cnt} bản ghi đã được **Duyệt** nên GDV không có quyền xóa!")
                                    else:
                                        st.warning(f"⚠️ Đã chuyển các bản ghi sang trạng thái **Xóa** thành công.")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Lỗi xóa: {ex}")
                    with col_act3:
                        if st.button("⚡ Thực Thi Hạch Toán Kho (XL)", type="primary", key="btn_xl_compact", use_container_width=True):
                            if not selected_seqs_xl:
                                st.warning("⚠️ Vui lòng tích chọn ít nhất một giao dịch trên bảng!")
                            else:
                                try:
                                    df_mgr['TRANG THAI'] = df_mgr['TRANG THAI'].astype(str)
                                    df_mgr['NGUOI XL'] = df_mgr['NGUOI XL'].astype(str)
                                    
                                    count_success = 0
                                    err_messages = []
                                    
                                    for seq in selected_seqs_xl:
                                        row_item = df_mgr[df_mgr['SO SEQ'] == seq]
                                        if not row_item.empty:
                                            row_data_item = row_item.iloc[0]
                                            curr_stt = str(row_data_item['TRANG THAI'])
                                            
                                            if curr_stt == 'Duyệt':
                                                ccmg_dict = _extract_row_ccmg(row_data_item)
                                                
                                                success, msg = execute_hach_toan_kho(
                                                    path_gl,
                                                    row_data_item,
                                                    ccmg_dict,
                                                    df_all_tx=df_mgr,
                                                    current_seq=seq
                                                )
                                                if success:
                                                    df_mgr.loc[df_mgr['SO SEQ'] == seq, 'TRANG THAI'] = 'XL'
                                                    df_mgr.loc[df_mgr['SO SEQ'] == seq, 'NGUOI XL'] = str(current_user)
                                                    count_success += 1
                                                else:
                                                    err_messages.append(f"SEQ {seq}: {msg}")
                                            else:
                                                err_messages.append(f"SEQ {seq}: Giao dịch phải ở trạng thái 'Duyệt' mới được hạch toán kho.")
                                                
                                    for col in df_mgr.columns:
                                        if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                            df_mgr[col] = pd.to_numeric(df_mgr[col], errors='coerce').fillna(0.0)

                                    df_mgr.to_excel(path_ds, index=False)
                                    
                                    final_msg_parts = []
                                    if count_success > 0:
                                        final_msg_parts.append(f"🎉 Đã hạch toán kho thành công {count_success} giao dịch và cập nhật vào Sổ cái!")
                                    if err_messages:
                                        for err in err_messages:
                                            final_msg_parts.append(f"❌ {err}")
                                            
                                    st.session_state['persistent_msg'] = "\n\n".join(final_msg_parts)
                                    st.session_state['persistent_type'] = "error" if err_messages else "success"
                                    st.rerun()
                                except Exception as ex:
                                    st.session_state['persistent_msg'] = f"❌ Lỗi thực hiện hạch toán kho: {ex}"
                                    st.session_state['persistent_type'] = "error"
                                    st.rerun()

                    if 'sua_seq_dang_chon_gdv' in st.session_state and st.session_state['sua_seq_dang_chon_gdv'] is not None:
                        target_seq_gdv = st.session_state['sua_seq_dang_chon_gdv']
                        row_target_gdv = df_mgr[df_mgr['SO SEQ'] == target_seq_gdv]
                        
                        if not row_target_gdv.empty:
                            r_val_g = row_target_gdv.iloc[0]
                            with st.container(border=True):
                                st.markdown(f"#### ✍️ [GDV] Điều Chỉnh Cơ Cấu Mệnh Giá (Mã SEQ: `{target_seq_gdv}` - KH: `{r_val_g['TEN KH']}`)")
                                
                                edit_mg_cols_g = st.columns(5)
                                new_ccmg_gdv = {}
                                new_total_gdv = 0
                                
                                def make_callback_gdv(s_id, denomination):
                                    def cb():
                                        raw_text = st.session_state.get(f"input_gdv_{s_id}_{denomination}", "0")
                                        try:
                                            clean_str = str(raw_text).replace(",", "").strip()
                                            num_val = int(float(clean_str)) if clean_str else 0
                                            st.session_state[f"input_gdv_{s_id}_{denomination}"] = f"{num_val:,}"
                                        except:
                                            st.session_state[f"input_gdv_{s_id}_{denomination}"] = "0"
                                    return cb

                                for idx, mg in enumerate(DENOMINATIONS):
                                    with edit_mg_cols_g[idx % 5]:
                                        key_name_g = f"input_gdv_{target_seq_gdv}_{mg}"
                                        if key_name_g not in st.session_state:
                                            old_to_g = 0
                                            if mg in r_val_g and pd.notna(r_val_g[mg]):
                                                old_to_g = int(r_val_g[mg])
                                            elif str(mg) in r_val_g and pd.notna(r_val_g[str(mg)]):
                                                old_to_g = int(r_val_g[str(mg)])
                                            st.session_state[key_name_g] = f"{old_to_g:,}"

                                        st.text_input(
                                            f"{mg:,} đ", 
                                            key=key_name_g,
                                            on_change=make_callback_gdv(target_seq_gdv, mg)
                                        )
                                        
                                        try:
                                            c_str_g = str(st.session_state[key_name_g]).replace(",", "").strip()
                                            val_sua_g = int(float(c_str_g)) if c_str_g else 0
                                        except:
                                            val_sua_g = 0
                                            
                                        new_ccmg_gdv[mg] = val_sua_g
                                        new_total_gdv += val_sua_g * mg

                                st.metric(label="💰 Tổng Giá Trị Sau Khi Sửa", value=f"{new_total_gdv:,.0f} VNĐ")
                                
                                col_fg1, col_fg2 = st.columns(2)
                                with col_fg1:
                                    if st.button("💾 Lưu Sửa Đổi (Gửi Duyệt Lại)", type="primary", key="btn_gdv_save_ccmg", use_container_width=True):
                                        ok_v_gdv, msg_v_gdv = validate_chi_kha_dung(
                                            path_gl,
                                            thu_chi_val=r_val_g.get('THU CHI', ''),
                                            loai_nv_val=r_val_g.get('LOAI NV', ''),
                                            don_vi_xl_val=r_val_g.get('DON VI XL', ''),
                                            ma_kh_val=r_val_g.get('MA KH CCCD', ''),
                                            ccmg_data=new_ccmg_gdv,
                                            df_all_tx=df_mgr,
                                            exclude_seq=target_seq_gdv
                                        )
                                        if not ok_v_gdv:
                                            st.session_state['persistent_msg'] = f"{msg_v_gdv} — Chặn gửi duyệt lại, vui lòng điều chỉnh lại số tờ theo tồn khả dụng!"
                                            st.session_state['persistent_type'] = "error"
                                            st.rerun()
                                        else:
                                            try:
                                                for mg in DENOMINATIONS:
                                                    so_to_moi_g = int(new_ccmg_gdv[mg])
                                                    col_key = mg if mg in df_mgr.columns else (str(mg) if str(mg) in df_mgr.columns else None)
                                                    if col_key:
                                                        df_mgr.loc[df_mgr['SO SEQ'] == target_seq_gdv, col_key] = float(so_to_moi_g)
                                                    alt_t = f"{mg}_THANHTIEN"
                                                    if alt_t in df_mgr.columns:
                                                        df_mgr.loc[df_mgr['SO SEQ'] == target_seq_gdv, alt_t] = float(so_to_moi_g * mg)
                                                        
                                                df_mgr.loc[df_mgr['SO SEQ'] == target_seq_gdv, 'TONG THANH TIEN'] = float(new_total_gdv)
                                                df_mgr.loc[df_mgr['SO SEQ'] == target_seq_gdv, 'TRANG THAI'] = 'Chờ Phê Duyệt'
                                                
                                                for col in df_mgr.columns:
                                                    if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                                        df_mgr[col] = pd.to_numeric(df_mgr[col], errors='coerce').fillna(0.0)

                                                df_mgr.to_excel(path_ds, index=False)
                                                del st.session_state['sua_seq_dang_chon_gdv']
                                                st.session_state['persistent_msg'] = f"🎉 Đã cập nhật cơ cấu mệnh giá và chuyển giao dịch SEQ {target_seq_gdv} về trạng thái **Chờ Phê Duyệt** thành công!"
                                                st.session_state['persistent_type'] = "success"
                                                st.rerun()
                                            except Exception as ex:
                                                st.session_state['persistent_msg'] = f"❌ Lỗi: {ex}"
                                                st.session_state['persistent_type'] = "error"
                                                st.rerun()
                                with col_fg2:
                                    if st.button("↩️ Hủy Bỏ", key="btn_gdv_cancel_ccmg", use_container_width=True):
                                        del st.session_state['sua_seq_dang_chon_gdv']
                                        st.rerun()

    elif current_tab == "📊 Báo Cáo":
        st.markdown("#### 📊 Báo Cáo Giao Dịch & Tồn Quỹ Kho")
        bc_type = st.radio("Chọn loại báo cáo:", ["1. Báo cáo Giao dịch tiền mới", "2. Báo cáo Tồn quỹ kho (GL)"], horizontal=True)
        st.markdown("---")
        
        if bc_type.startswith("1"):
            st.markdown("##### 🔍 Bộ Lọc Dữ Liệu Báo Cáo Giao Dịch")
            if os.path.exists(path_ds):
                df_bc1 = pd.read_excel(path_ds)
                
                with st.container(border=True):
                    f_col1, f_col2, f_col3 = st.columns(3)
                    with f_col1:
                        min_dt = pd.to_datetime(df_bc1['NGAY HIEU LUC'], errors='coerce').min()
                        max_dt = pd.to_datetime(df_bc1['NGAY HIEU LUC'], errors='coerce').max()
                        if pd.isna(min_dt): min_dt = date.today()
                        if pd.isna(max_dt): max_dt = date.today()
                        
                        tu_ngay = st.date_input("📅 Từ Ngày", value=min_dt, key="bc_tungay")
                    with f_col2:
                        den_ngay = st.date_input("📅 Đến Ngày", value=max_dt, key="bc_denngay")
                    with f_col3:
                        all_dv = ["TẤT CẢ"] + sorted(list(df_bc1['DON VI XL'].dropna().astype(str).unique()))
                        chon_dv = st.selectbox("🏢 Đơn Vị Xử Lý", all_dv, key="bc_donvi")
                    
                    f_col4, _ = st.columns(2)
                    with f_col4:
                        all_tt = ["TẤT CẢ"] + sorted(list(df_bc1['TRANG THAI'].dropna().astype(str).unique()))
                        chon_tt = st.selectbox("📌 Trạng Thái", all_tt, key="bc_trangthai")
                
                df_filtered_bc1 = df_bc1.copy()
                df_filtered_bc1['NGAY_DT'] = pd.to_datetime(df_filtered_bc1['NGAY HIEU LUC'], errors='coerce').dt.date
                
                df_filtered_bc1 = df_filtered_bc1[
                    (df_filtered_bc1['NGAY_DT'] >= tu_ngay) & 
                    (df_filtered_bc1['NGAY_DT'] <= den_ngay)
                ]
                if chon_dv != "TẤT CẢ":
                    df_filtered_bc1 = df_filtered_bc1[df_filtered_bc1['DON VI XL'].astype(str) == chon_dv]
                if chon_tt != "TẤT CẢ":
                    df_filtered_bc1 = df_filtered_bc1[df_filtered_bc1['TRANG THAI'].astype(str) == chon_tt]
                
                df_filtered_bc1 = df_filtered_bc1.drop(columns=['NGAY_DT'], errors='ignore')
                
                st.markdown(f"##### 📋 Danh Sách Giao Dịch (Hiển thị {len(df_filtered_bc1)} bản ghi)")
                
                df_display = df_filtered_bc1.copy()
                num_cols_bc = [c for c in df_display.columns if isinstance(c, (int, float)) or str(c).isdigit() or 'THANH TIEN' in str(c) or 'TONG' in str(c)]
                for col in num_cols_bc:
                    df_display[col] = pd.to_numeric(pd.Series(df_display[col].values.ravel()), errors='coerce').fillna(0).apply(lambda x: f"{int(x):,}" if x == int(x) else f"{x:,.2f}")
                
                st.dataframe(df_display, use_container_width=True)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_filtered_bc1.to_excel(writer, index=False, sheet_name='GiaoDichTienMoi')
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 Tải Xuống Báo Cáo Giao Dịch Đã Lọc (Excel .xlsx)",
                    data=processed_data,
                    file_name=f"BaoCao_GiaoDich_{date.today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.markdown("##### 📦 Báo Cáo Tồn Quỹ Kho (GL) - Cấu Trúc 2 Khối (TRONG KHO & DON VI)")
            if os.path.exists(path_gl):
                wb_gl_view = openpyxl.load_workbook(path_gl)
                ws_gl_view = wb_gl_view.active
                
                data_gl = []
                for r in range(1, ws_gl_view.max_row + 1):
                    row_vals = [ws_gl_view.cell(row=r, column=c).value for c in range(1, ws_gl_view.max_column + 1)]
                    data_gl.append(row_vals)
                
                if len(data_gl) >= 2:
                    raw_headers = data_gl[1]
                    unique_headers = []
                    seen_counts = {}
                    
                    for idx, h in enumerate(raw_headers):
                        h_str = str(h).strip() if pd.notna(h) else f"Col_{idx+1}"
                        if idx < 2:
                            unique_headers.append(h_str)
                        else:
                            block_suffix = " (Kho)" if idx < 21 else " (ĐV)"
                            full_name = f"{h_str}{block_suffix}"
                            if full_name in seen_counts:
                                seen_counts[full_name] += 1
                                full_name = f"{full_name} {seen_counts[full_name]}"
                            else:
                                seen_counts[full_name] = 0
                            unique_headers.append(full_name)
                    
                    df_gl_display = pd.DataFrame(data_gl[2:], columns=unique_headers)
                    
                    search_dv = st.text_input("🔍 Tìm kiếm theo Mã Đơn Vị hoặc Tên Đơn Vị", key="search_gl_dv")
                    if search_dv.strip():
                        mask = df_gl_display.iloc[:, 0].astype(str).str.contains(search_dv, case=False, na=False) | \
                               df_gl_display.iloc[:, 1].astype(str).str.contains(search_dv, case=False, na=False)
                        df_gl_display = df_gl_display[mask]
                    
                    st.markdown("---")
                    st.markdown("###### 📊 Bảng Số Dư Tồn Quỹ:")
                    
                    df_gl_formatted = df_gl_display.copy()
                    for col_idx in range(2, df_gl_formatted.shape[1]):
                        col_name = df_gl_formatted.columns[col_idx]
                        col_vals = df_gl_formatted[col_name].values.ravel()
                        df_gl_formatted[col_name] = pd.to_numeric(pd.Series(col_vals), errors='coerce').fillna(0).apply(lambda x: f"{int(x):,}" if x == int(x) else f"{x:,.2f}")
                    
                    st.dataframe(df_gl_formatted, use_container_width=True)
                    
                    output_gl = BytesIO()
                    wb_gl_out = openpyxl.Workbook()
                    ws_gl_out = wb_gl_out.active
                    ws_gl_out.title = "TonQuyKho"
                    
                    for r_idx, row in enumerate(data_gl, start=1):
                        for c_idx, val in enumerate(row, start=1):
                            cell = ws_gl_out.cell(row=r_idx, column=c_idx, value=val)
                            if r_idx >= 3 and c_idx > 2 and val is not None:
                                try:
                                    cell.value = float(val)
                                    cell.number_format = '#,##0'
                                except:
                                    pass
                    
                    wb_gl_out.save(output_gl)
                    gl_excel_data = output_gl.getvalue()
                    
                    st.download_button(
                        label="📥 Tải Xuống Báo Cáo Tồn Quỹ Kho (Excel .xlsx)",
                        data=gl_excel_data,
                        file_name=f"BaoCao_TonQuyKho_GL_{date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
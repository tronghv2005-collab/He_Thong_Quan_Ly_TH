import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import os
import openpyxl
from io import BytesIO

DENOMINATIONS = [500000, 200000, 100000, 50000, 20000, 10000, 5000, 2000, 1000]

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

def get_list_accounts_from_coso(path_coso):
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

def execute_hach_toan_kho(path_gl, row_item, ccmg_data):
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
    
    header_row = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    
    def get_denomination_cols_wb(start_col, end_col):
        cols_map = {}
        for col_idx_1based in range(start_col, end_col + 1):
            val = header_row[col_idx_1based - 1]
            try:
                mg_val = int(float(val))
                if mg_val in DENOMINATIONS:
                    cols_map[mg_val] = {'to_col': col_idx_1based, 'tt_col': col_idx_1based + 1}
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

    idx_990 = get_row_index_wb('990')
    idx_dvxl = ensure_row_exists_in_gl(ws, don_vi_xl)
    idx_makh = get_row_index_wb(ma_kh) if loai_nv != "Cá Nhân" else None

    if idx_dvxl is None:
        return False, f"Lỗi: Không tìm thấy đơn vị xử lý '{don_vi_xl}' trong Sổ cái!"

    is_gdv_account = ".gdv" in don_vi_xl.lower()

    def check_du_chi_wb(target_row, kho_block_map, block_name="Kho"):
        if target_row is None:
            return False, f"Mã kho/tài khoản không tồn tại trong Sổ cái!"
        for mg in DENOMINATIONS:
            so_to_yc = float(ccmg_data.get(mg, 0) or 0)
            if so_to_yc > 0 and mg in kho_block_map:
                to_col = kho_block_map[mg]['to_col']
                val = ws.cell(row=target_row, column=to_col).value
                hien_co = float(val) if val is not None and str(val).strip() != '' else 0.0
                if hien_co < so_to_yc:
                    return False, f"Không Đủ Chi ({block_name} - Mệnh giá {mg:,} VNĐ: Chỉ còn {int(hien_co):,} tờ, yêu cầu {int(so_to_yc):,} tờ)!"
        return True, ""

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
                ws.cell(row=target_row, column=to_col, value=new_to).number_format = '#,##0'
                
                cur_tt_val = ws.cell(row=target_row, column=tt_col).value
                cur_tt = float(cur_tt_val) if cur_tt_val is not None and str(cur_tt_val).strip() != '' else 0.0
                new_tt = (cur_tt + thanh_tien) if is_tang else max(0.0, cur_tt - thanh_tien)
                ws.cell(row=target_row, column=tt_col, value=new_tt).number_format = '#,##0'
                
        if sum_col_idx <= ws.max_column:
            cur_sum_val = ws.cell(row=target_row, column=sum_col_idx).value
            cur_sum = float(cur_sum_val) if cur_sum_val is not None and str(cur_sum_val).strip() != '' else 0.0
            new_sum = max(0.0, cur_sum + total_tt_delta)
            ws.cell(row=target_row, column=sum_col_idx, value=new_sum).number_format = '#,##0'

    if thu_chi == "CHI":
        if is_gdv_account:
            ok_dv, msg_dv = check_du_chi_wb(idx_dvxl, dv_map, block_name="Khối Đơn Vị (GDV)")
            if not ok_dv: return False, f"❌ Lỗi Tồn Kho Xuất Chi tại '{don_vi_xl}': {msg_dv}"
        else:
            ok_tk, msg_tk = check_du_chi_wb(idx_dvxl, tk_map, block_name="Khối TRONG KHO")
            if not ok_tk: return False, f"❌ Lỗi Tồn Kho Xuất Chi tại '{don_vi_xl}': {msg_tk}"
            ok_dv, msg_dv = check_du_chi_wb(idx_dvxl, dv_map, block_name="Khối DON VI")
            if not ok_dv: return False, f"❌ Lỗi Tồn Kho Xuất Chi tại '{don_vi_xl}': {msg_dv}"

        if loai_nv in ["Tạm ứng", "Chi nhánh"]:
            ok_990, msg_990 = check_du_chi_wb(idx_990, tk_map, block_name="Kho Tổng 990")
            if not ok_990: return False, f"❌ Lỗi Tồn Kho tại Kho Tổng 990: {msg_990}"

    if thu_chi == "THU" and loai_nv == "Chi nhánh":
        apply_delta_block_wb(idx_990, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
    elif thu_chi == "THU" and loai_nv == "NHNN":
        apply_delta_block_wb(idx_990, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, True, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, dv_map, 40)
    elif thu_chi == "CHI" and loai_nv == "Nội bộ":
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_makh, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
    elif thu_chi == "CHI" and loai_nv == "Chi nhánh":
        apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, False, ccmg_data, dv_map, 40)
    elif thu_chi == "CHI" and loai_nv == "Cá Nhân":
        if is_gdv_account:
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        else:
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
    elif thu_chi == "CHI" and loai_nv == "Tạm ứng":
        apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        if idx_makh: apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
    elif thu_chi == "THU" and loai_nv == "Hoàn ứng":
        apply_delta_block_wb(idx_990, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, True, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, dv_map, 40)
        if idx_makh: apply_delta_block_wb(idx_makh, False, ccmg_data, dv_map, 40)
    elif loai_nv == "Điều chuyển CN":
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
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
    # 🎨 CSS TỐI ƯU TRẢI NGHIỆM NGƯỜI DÙNG: Đẩy sát lên trên, thu gọn khoảng cách các dòng widget tối đa
    st.markdown("""
        <style>
            .block-container {
                padding-top: 0.8rem !important;
                padding-bottom: 1rem !important;
                padding-left: 1.5rem !important;
                padding-right: 1.5rem !important;
            }
            div.stElementContainer {
                margin-bottom: -10px !important;
            }
            h4, h5 {
                margin-top: 0px !important;
                margin-bottom: 4px !important;
            }
            hr {
                margin-top: 8px !important;
                margin-bottom: 8px !important;
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
                    if not found: df_kh_check[col] = ''
                    changed = True
            if changed: df_kh_check.to_excel(path_kh, index=False)
        except:
            pd.DataFrame(columns=['MA CCCD', 'TEN KH', 'DIACHI KH']).to_excel(path_kh, index=False)
    else:
        pd.DataFrame(columns=['MA CCCD', 'TEN KH', 'DIACHI KH']).to_excel(path_kh, index=False)
    
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
                coso_options = [f"{str(row[col_kho]).strip()} - {str(row[col_ten]).strip() if col_ten and pd.notna(row[col_ten]) else ''}" for _, row in df_cs.iterrows() if str(row[col_kho]).strip().lower() != 'nan']
        except:
            pass

    gl_accounts = get_list_accounts_from_gl(path_gl)
    coso_accounts = get_list_accounts_from_coso(path_coso)

    if 'bind_cccd' not in st.session_state: st.session_state['bind_cccd'] = ""
    if 'bind_ten' not in st.session_state: st.session_state['bind_ten'] = ""
    if 'bind_dc' not in st.session_state: st.session_state['bind_dc'] = ""
    if 'is_nhap_moi_active' not in st.session_state: st.session_state['is_nhap_moi_active'] = False
    if 'persistent_msg' not in st.session_state: st.session_state['persistent_msg'] = None
    if 'persistent_type' not in st.session_state: st.session_state['persistent_type'] = None

    for mg in DENOMINATIONS:
        if f"to_{mg}" not in st.session_state: st.session_state[f"to_{mg}"] = 0
        if f"bo_{mg}" not in st.session_state: st.session_state[f"bo_{mg}"] = 0.0
        if f"bao_{mg}" not in st.session_state: st.session_state[f"bao_{mg}"] = 0.0

    tabs_list = ["📥 Giao Dịch Thu/Chi", "🔄 Ứng / Hoàn Tiền", "🛡️ Phê Duyệt (LĐ)", "🔍 Tìm Kiếm & Xử Lý", "📊 Báo Cáo"]
    current_tab = st.radio("Chuyển Tab Chức Năng", tabs_list, horizontal=True, label_visibility="collapsed")

    if st.session_state['persistent_msg']:
        p_type, p_msg = st.session_state['persistent_type'], st.session_state['persistent_msg']
        getattr(st, p_type)(p_msg) if p_type in ['success', 'error', 'warning'] else st.info(p_msg)
        if st.button("✖️ Đóng thông báo", key="btn_close_p_msg"):
            st.session_state['persistent_msg'], st.session_state['persistent_type'] = None, None
            st.rerun()
        st.markdown("---")

    if current_tab == "📥 Giao Dịch Thu/Chi":
        if is_giam_doc or is_ld_hub or role_upper in ["LĐTT", "ADMIN"]:
            st.warning("⚠️ Tài khoản nhóm quản lý cấp cao / Lãnh đạo không có quyền lập hoặc nhập liệu giao dịch.")
        else:
            col_title, col_btn_nhapmoi = st.columns([3, 1])
            with col_title: st.markdown("#### 📝 Lập Giao Dịch Thu / Chi Tiền Mới")
            with col_btn_nhapmoi:
                if not st.session_state['is_nhap_moi_active']:
                    if st.button("➕ Nhập Mới", type="primary", use_container_width=True):
                        st.session_state['is_nhap_moi_active'] = True
                        st.session_state['bind_cccd'], st.session_state['bind_ten'], st.session_state['bind_dc'] = "", "", ""
                        st.session_state['persistent_msg'] = None
                        if 'gd_select_kh_chung' in st.session_state: st.session_state['gd_select_kh_chung'] = "-- Chọn khách hàng từ danh sách hoặc nhập mới --"
                        for mg in DENOMINATIONS:
                            st.session_state[f"to_{mg}"] = 0
                            st.session_state[f"bo_{mg}"] = 0.0
                            st.session_state[f"bao_{mg}"] = 0.0
                        st.rerun()
                else:
                    if st.button("🔄 Làm Mới", type="secondary", use_container_width=True):
                        st.session_state['bind_cccd'], st.session_state['bind_ten'], st.session_state['bind_dc'] = "", "", ""
                        st.session_state['persistent_msg'], st.session_state['persistent_type'] = "✨ Đã làm mới màn hình!", "success"
                        if 'gd_select_kh_chung' in st.session_state: st.session_state['gd_select_kh_chung'] = "-- Chọn khách hàng từ danh sách hoặc nhập mới --"
                        for mg in DENOMINATIONS:
                            st.session_state[f"to_{mg}"] = 0
                            st.session_state[f"bo_{mg}"] = 0.0
                            st.session_state[f"bao_{mg}"] = 0.0
                        st.rerun()

            if not st.session_state['is_nhap_moi_active']:
                st.info("ℹ️ Vui lòng bấm vào nút **'➕ Nhập Mới'** ở góc trên bên phải để bắt đầu tạo giao dịch.")
            else:
                col1, col2, col3 = st.columns(3)
                with col1:
                    ngay_hieu_luc = st.date_input("Ngày Hiệu Lực", value=date.today(), key="gd_ngay")
                    loai_nv = st.selectbox("Loại Nghiệp Vụ", ["Cá Nhân", "Chi nhánh", "NHNN", "Nội bộ", "Tạm ứng", "Hoàn ứng", "Điều chuyển CN", "Điều chỉnh kiểm kê"], key="gd_loainv")
                with col2:
                    don_vi_xl = st.selectbox("Đơn Vị Xử Lý (Kho)", coso_options, key="gd_donvixl").split(" - ")[0].strip()
                    if loai_nv in ["Tạm ứng", "Cá Nhân", "Nội bộ"]:
                        thu_chi = "CHI"
                        st.selectbox("Thu / Chi", ["CHI"], disabled=True, key="gd_thuchi_locked")
                    elif loai_nv == "Hoàn ứng":
                        thu_chi = "THU"
                        st.selectbox("Thu / Chi", ["THU"], disabled=True, key="gd_thuchi_locked")
                    else:
                        thu_chi = st.selectbox("Thu / Chi", ["THU", "CHI"], key="gd_thuchi")
                with col3:
                    if loai_nv == "Nội bộ":
                        if coso_accounts:
                            sel_acc_cs = st.selectbox("Chọn Đơn Vị Nhận (Từ Thamso_coso)", ["-- Chọn đơn vị --"] + coso_accounts, key="gd_sel_acc_coso")
                            if sel_acc_cs != "-- Chọn đơn vị --":
                                ma_kh = sel_acc_cs.split(" - ")[0].strip()
                                default_ten_kh = " - ".join(sel_acc_cs.split(" - ")[1:])
                            else:
                                ma_kh = st.text_input("Mã Đơn Vị Nhận", key="gd_makh_coso_free")
                                default_ten_kh = ""
                        else:
                            ma_kh = st.text_input("Mã Đơn Vị Nhận", key="gd_makh_coso_free")
                            default_ten_kh = ""
                        diachi_kh = st.text_input("Địa Chỉ", key="gd_diachi_coso")
                        ten_kh = st.text_input("Tên Đơn Vị Nhận", value=default_ten_kh.upper(), max_chars=30, key="gd_tenkh_coso")
                    elif loai_nv in ["Chi nhánh", "Hoàn ứng", "Tạm ứng", "Điều chuyển CN"]:
                        if gl_accounts:
                            sel_acc = st.selectbox("Chọn Mã Đơn Vị / KH từ Sổ Cái", ["-- Nhập tự do hoặc chọn --"] + gl_accounts, key="gd_sel_acc")
                            if sel_acc != "-- Nhập tự do hoặc chọn --":
                                ma_kh = sel_acc.split(" - ")[0].strip()
                                default_ten_kh = " - ".join(sel_acc.split(" - ")[1:])
                            else:
                                ma_kh = st.text_input("Mã KH / Mã Đơn Vị", key="gd_makh_free")
                                default_ten_kh = ""
                        else:
                            ma_kh = st.text_input("Mã KH / Mã Đơn Vị", key="gd_makh_free")
                            default_ten_kh = ""
                        diachi_kh = st.text_input("Địa Chỉ KH", key="gd_diachi_other")
                        ten_kh = st.text_input("Tên Khách Hàng / Đơn Vị", value=default_ten_kh.upper(), max_chars=30, key="gd_tenkh")
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

                        st.selectbox("🔍 Tìm và Chọn Khách Hàng", list(kh_options_dict.keys()), key="gd_select_kh_chung", on_change=on_change_chon_kh_v2)
                        ma_kh = st.text_input("Mã CCCD / Mã KH", key="bind_cccd")
                        ten_kh = st.text_input("Tên Khách Hàng (In hoa)", max_chars=30, key="bind_ten")
                        diachi_kh = st.text_input("Địa Chỉ KH", key="bind_dc")

                st.markdown("---")
                st.markdown("##### 🔢 Nhập Số Tờ / Số Lượng Chi Tiết Theo Mệnh Giá")
                unit_mode = st.radio("Đơn vị tính:", ["Số tờ", "Theo Bó (1.000 tờ)", "Theo Bao (20.000 tờ)"], horizontal=True, key="gd_unit_mode")
                
                ccmg_data = {}
                tong_thanh_tien = 0
                
                for i in range(0, len(DENOMINATIONS), 2):
                    row_mg_cols = st.columns(2)
                    for col_idx, j in enumerate(range(i, min(i + 2, len(DENOMINATIONS)))):
                        mg = DENOMINATIONS[j]
                        with row_mg_cols[col_idx]:
                            st.markdown(f"**💵 Mệnh giá: {mg:,} VNĐ**")
                            sub_c1, sub_c2 = st.columns([3, 2])
                            with sub_c1:
                                if unit_mode.startswith("Số tờ"):
                                    so_to = int(st.number_input("Số tờ", min_value=0, step=100, format="%d", key=f"to_{mg}"))
                                elif unit_mode.startswith("Theo Bó"):
                                    so_bo = st.number_input("Số Bó", min_value=0.0, step=1.0, format="%.2f", key=f"bo_{mg}")
                                    so_to = int(so_bo * 1000)
                                else:
                                    so_bao = st.number_input("Số Bao", min_value=0.0, step=0.5, format="%.2f", key=f"bao_{mg}")
                                    so_to = int(so_bao * 20000)
                            with sub_c2:
                                so_bo_chan = int(so_to // 1000)
                                st.number_input("Số bó chẵn", value=so_bo_chan, disabled=True, key=f"bochan_{mg}")
                            ccmg_data[mg] = so_to
                            tong_thanh_tien += so_to * mg
                    st.markdown("")

                st.info(f"💰 **Tổng Giá Trị Giao Dịch:** `{tong_thanh_tien:,.0f}` VNĐ")
                
                if st.button("💾 Cập Nhật Bản Ghi (Chờ Phê Duyệt)", type="primary"):
                    if not ma_kh.strip() and loai_nv != "Điều chỉnh kiểm kê":
                        st.session_state['persistent_msg'], st.session_state['persistent_type'] = "❌ Mã định danh không được để trống!", "error"
                        st.rerun()
                    elif not ten_kh.strip():
                        st.session_state['persistent_msg'], st.session_state['persistent_type'] = "❌ Tên khách hàng không được để trống!", "error"
                        st.rerun()
                    elif tong_thanh_tien <= 0:
                        st.session_state['persistent_msg'], st.session_state['persistent_type'] = "❌ Vui lòng nhập số lượng tiền lớn hơn 0!", "error"
                        st.rerun()
                    else:
                        try:
                            if os.path.exists(path_ds):
                                df_ds = pd.read_excel(path_ds)
                            else:
                                cols = ['MA KH CCCD', 'TEN KH', 'DIACHI KH', 'DON VI XL', 'LOAI NV', 'SO SEQ', 'NGAY HIEU LUC', 'THU CHI', 'TRANG THAI']
                                for d in DENOMINATIONS: cols.extend([str(d), f"{d}_THANHTIEN"])
                                cols.extend(['TONG THANH TIEN', 'GDV TM', 'NGUOI XL', 'NGUOI PHE DUYET'])
                                df_ds = pd.DataFrame(columns=cols)
                            
                            next_seq = int(pd.to_numeric(df_ds['SO SEQ'], errors='coerce').max() + 1) if 'SO SEQ' in df_ds.columns and not df_ds.empty else 1
                            new_row = {
                                'MA KH CCCD': ma_kh.strip() if ma_kh.strip() else don_vi_xl,
                                'TEN KH': ten_kh.strip().upper(),
                                'DIACHI KH': diachi_kh.strip(),
                                'DON VI XL': don_vi_xl,
                                'LOAI NV': loai_nv,
                                'SO SEQ': float(next_seq),
                                'NGAY HIEU LUC': ngay_hieu_luc.strftime('%Y-%m-%d'),
                                'THU CHI': thu_chi,
                                'TRANG THAI': 'Chờ Phê Duyệt',
                                'GDV TM': str(current_user),
                                'NGUOI XL': '', 'NGUOI PHE DUYET': '',
                                'TONG THANH TIEN': float(tong_thanh_tien)
                            }
                            for idx_mg, mg in enumerate(DENOMINATIONS):
                                so_to_val = float(ccmg_data[mg])
                                thanh_tien_val = so_to_val * mg
                                if mg in df_ds.columns: new_row[mg] = so_to_val
                                if str(mg) in df_ds.columns: new_row[str(mg)] = so_to_val
                                tt_name = f"THANH TIEN.{idx_mg}" if idx_mg > 0 else 'THANH TIEN'
                                if tt_name in df_ds.columns: new_row[tt_name] = thanh_tien_val
                                alt_tt = f"{mg}_THANHTIEN"
                                if alt_tt in df_ds.columns: new_row[alt_tt] = thanh_tien_val

                            df_ds = pd.concat([df_ds, pd.DataFrame([{col: new_row.get(col, 0.0) if ('THANH TIEN' in str(col) or str(col).isdigit() or col in DENOMINATIONS) else str(new_row.get(col, '')) for col in df_ds.columns}])], ignore_index=True)
                            df_ds.to_excel(path_ds, index=False)
                            st.session_state['is_nhap_moi_active'] = False
                            st.session_state['persistent_msg'], st.session_state['persistent_type'] = f"🎉 Tạo giao dịch thành công (Mã SEQ: {next_seq})!", "success"
                            st.rerun()
                        except Exception as e:
                            st.session_state['persistent_msg'], st.session_state['persistent_type'] = f"❌ Lỗi: {e}", "error"
                            st.rerun()

    elif current_tab == "🔄 Ứng / Hoàn Tiền":
        st.markdown("#### 🔄 Giao Dịch Tạm Ứng / Hoàn Ứng Nhanh")
        st.info("Chức năng đồng bộ tương tự tab Giao dịch với nghiệp vụ Tạm ứng hoặc Hoàn ứng tự động.")

    elif current_tab == "🛡️ Phê Duyệt (LĐ)":
        st.markdown("#### 🛡️ Phê Duyệt Giao Dịch & Sửa Cơ Cấu Mệnh Giá")
        if not is_giam_doc and not is_ld_hub:
            st.warning("⚠️ Chức năng chỉ dành cho nhóm quyền **GIÁM ĐỐC** và **LĐ_HUB**.")
        else:
            if os.path.exists(path_ds):
                df_app = pd.read_excel(path_ds)
                if 'TRANG THAI' not in df_app.columns: df_app['TRANG THAI'] = 'Chờ Phê Duyệt'
                df_cho = df_app[df_app['TRANG THAI'].astype(str).str.strip() == 'Chờ Phê Duyệt'].copy()
                
                if df_cho.empty:
                    st.info("ℹ️ Hiện không có bản ghi nào chờ phê duyệt.")
                else:
                    df_cho_display_data = []
                    for _, r_row in df_cho.iterrows():
                        row_dict = {
                            'SO SEQ': r_row.get('SO SEQ', 0),
                            'TEN KH': r_row.get('TEN KH', ''),
                            'LOAI NV': r_row.get('LOAI NV', ''),
                            'DON VI XL': r_row.get('DON VI XL', ''),
                            'TRANG THAI': r_row.get('TRANG THAI', '')
                        }
                        for mg in DENOMINATIONS:
                            val_to = int(r_row[mg] if mg in r_row and pd.notna(r_row[mg]) else r_row.get(str(mg), 0))
                            row_dict[f"{mg:,} (Tờ)"] = f"{val_to:,}"
                            row_dict[f"{mg:,} (Bó)"] = f"{val_to // 1000:,}"
                        row_dict['TONG THANH TIEN'] = f"{int(float(r_row.get('TONG THANH TIEN', 0))):,}"
                        df_cho_display_data.append(row_dict)
                        
                    df_disp = pd.DataFrame(df_cho_display_data)
                    df_disp.insert(0, "Chọn", False)
                    
                    edited_app = st.data_editor(df_disp, hide_index=True, use_container_width=True, key="ed_ld_v3")
                    sel_app = edited_app[edited_app['Chọn'] == True]['SO SEQ'].tolist()
                    
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        if st.button("✅ Phê Duyệt Nhanh", type="primary"):
                            if sel_app:
                                df_app.loc[df_app['SO SEQ'].isin(sel_app), ['TRANG THAI', 'NGUOI PHE DUYET']] = ['Duyệt', str(current_user)]
                                df_app.to_excel(path_ds, index=False)
                                st.success(f"🎉 Đã duyệt {len(sel_app)} giao dịch!")
                                st.rerun()
                    with c2:
                        if st.button("✏️ Sửa Cơ Cấu Mệnh Giá"):
                            if len(sel_app) == 1:
                                st.session_state['sua_seq_ld'] = sel_app[0]
                                r_i = df_cho[df_cho['SO SEQ'] == sel_app[0]].iloc[0]
                                for mg in DENOMINATIONS:
                                    val = int(r_i[mg] if mg in r_i and pd.notna(r_i[mg]) else r_i.get(str(mg), 0))
                                    st.session_state[f"ld_edit_{sel_app[0]}_{mg}"] = f"{val:,}"
                                st.rerun()
                            else:
                                st.warning("⚠️ Vui lòng chọn đúng 1 dòng để sửa!")
                    with c3:
                        if st.button("🗑️ Xóa Bản Ghi"):
                            if sel_app:
                                df_app.loc[df_app['SO SEQ'].isin(sel_app), 'TRANG THAI'] = 'Xóa'
                                df_app.to_excel(path_ds, index=False)
                                st.warning("⚠️ Đã xóa bản ghi.")
                                st.rerun()
                    with c4:
                        if st.button("❌ Từ Chối"):
                            if sel_app:
                                df_app.loc[df_app['SO SEQ'].isin(sel_app), ['TRANG THAI', 'NGUOI PHE DUYET']] = ['Từ Chối Phê Duyệt', str(current_user)]
                                df_app.to_excel(path_ds, index=False)
                                st.warning("⚠️ Đã từ chối giao dịch.")
                                st.rerun()

                    # 🎯 FORM SỬA CƠ CẤU MỆNH GIÁ: 4 CỘT TRÊN 1 HÀNG, SỐ TỜ RỘNG HƠN SỐ BÓ
                    if 'sua_seq_ld' in st.session_state and st.session_state['sua_seq_ld']:
                        t_seq = st.session_state['sua_seq_ld']
                        r_t = df_app[df_app['SO SEQ'] == t_seq].iloc[0]
                        st.markdown("---")
                        st.markdown(f"##### ✍️ [LÃNH ĐẠO] Điều Chỉnh Cơ Cấu Mệnh Giá (SEQ: `{t_seq}` - KH: `{r_t['TEN KH']}`)")
                        
                        new_ccmg = {}
                        new_tot = 0
                        
                        for i in range(0, len(DENOMINATIONS), 4):
                            row_cols = st.columns(4)
                            for col_idx, ji in enumerate(range(i, min(i + 4, len(DENOMINATIONS)))):
                                mg = DENOMINATIONS[ji]
                                k_n = f"ld_edit_{t_seq}_{mg}"
                                if k_n not in st.session_state:
                                    st.session_state[k_n] = f"{int(r_t.get(mg, r_t.get(str(mg), 0))):,}"
                                
                                with row_cols[col_idx]:
                                    st.markdown(f"<p style='font-size:12px; font-weight:bold; margin-bottom:2px;'>{mg:,} VNĐ</p>", unsafe_allow_html=True)
                                    sub_c1, sub_c2 = st.columns([3, 2]) # Tỷ lệ ô nhập số tờ rộng hơn ô số bó
                                    with sub_c1:
                                        val_s = st.text_input(f"s_{mg}", key=k_n, label_visibility="collapsed")
                                    try: v_num = int(float(val_s.replace(",", "").strip())) if val_s else 0
                                    except: v_num = 0
                                    with sub_c2:
                                        st.number_input(f"bc_{mg}", value=int(v_num // 1000), disabled=True, key=f"ld_bc_{t_seq}_{mg}", label_visibility="collapsed")
                                    new_ccmg[mg] = v_num
                                    new_tot += v_num * mg
                            st.markdown("")

                        st.info(f"💰 **Tổng Giá Trị Sau Khi Sửa:** `{new_tot:,.0f}` VNĐ")
                        col_f1, col_f2 = st.columns(2)
                        with col_f1:
                            if st.button("💾 Lưu Sửa Đổi & Phê Duyệt Ngay", type="primary"):
                                for mg in DENOMINATIONS:
                                    col_k = mg if mg in df_app.columns else (str(mg) if str(mg) in df_app.columns else None)
                                    if col_k: df_app.loc[df_app['SO SEQ'] == t_seq, col_k] = float(new_ccmg[mg])
                                    if f"{mg}_THANHTIEN" in df_app.columns: df_app.loc[df_app['SO SEQ'] == t_seq, f"{mg}_THANHTIEN"] = float(new_ccmg[mg] * mg)
                                df_app.loc[df_app['SO SEQ'] == t_seq, ['TONG THANH TIEN', 'TRANG THAI', 'NGUOI PHE DUYET']] = [float(new_tot), 'Duyệt', str(current_user)]
                                df_app.to_excel(path_ds, index=False)
                                del st.session_state['sua_seq_ld']
                                st.success("🎉 Lưu và phê duyệt thành công!")
                                st.rerun()
                        with col_f2:
                            if st.button("↩️ Hủy Bỏ"):
                                del st.session_state['sua_seq_ld']
                                st.rerun()

    elif current_tab == "🔍 Tìm Kiếm & Xử Lý":
        st.markdown("#### 🔍 Xử Lý Hạch Toán Kho (XL) & Quản Lý Bản Ghi")
        if os.path.exists(path_ds):
            df_mgr = pd.read_excel(path_ds)
            if 'TRANG THAI' not in df_mgr.columns: df_mgr['TRANG THAI'] = 'Chờ Phê Duyệt'
            
            df_editable = df_mgr[df_mgr['TRANG THAI'].astype(str).str.strip().isin(['Chờ Phê Duyệt', 'Từ Chối Phê Duyệt', 'Duyệt', 'Xóa'])].copy()
            if not df_editable.empty:
                df_mgr_display_data = []
                for _, r_row in df_editable.iterrows():
                    row_dict = {
                        'SO SEQ': r_row.get('SO SEQ', 0),
                        'TEN KH': r_row.get('TEN KH', ''),
                        'LOAI NV': r_row.get('LOAI NV', ''),
                        'DON VI XL': r_row.get('DON VI XL', ''),
                        'TRANG THAI': r_row.get('TRANG THAI', '')
                    }
                    for mg in DENOMINATIONS:
                        val_to = int(r_row[mg] if mg in r_row and pd.notna(r_row[mg]) else r_row.get(str(mg), 0))
                        row_dict[f"{mg:,} (Tờ)"] = f"{val_to:,}"
                        row_dict[f"{mg:,} (Bó)"] = f"{val_to // 1000:,}"
                    row_dict['TONG THANH TIEN'] = f"{int(float(r_row.get('TONG THANH TIEN', 0))):,}"
                    df_mgr_display_data.append(row_dict)
                
                df_mgr_disp = pd.DataFrame(df_mgr_display_data)
                df_mgr_disp.insert(0, "Chọn", False)
                edited_mgr = st.data_editor(df_mgr_disp, hide_index=True, use_container_width=True, key="ed_mgr_v3")
                sel_mgr = edited_mgr[edited_mgr['Chọn'] == True]['SO SEQ'].tolist()
                
                c_a1, c_a2, c_a3 = st.columns(3)
                with c_a1:
                    if st.button("✏️ Sửa Cơ Cấu Mệnh Giá (GDV)"):
                        if len(sel_mgr) == 1:
                            st.session_state['sua_seq_gdv'] = sel_mgr[0]
                            r_i = df_editable[df_editable['SO SEQ'] == sel_mgr[0]].iloc[0]
                            for mg in DENOMINATIONS:
                                val = int(r_i[mg] if mg in r_i and pd.notna(r_i[mg]) else r_i.get(str(mg), 0))
                                st.session_state[f"gdv_edit_{sel_mgr[0]}_{mg}"] = f"{val:,}"
                            st.rerun()
                        else:
                            st.warning("⚠️ Vui lòng chọn đúng 1 dòng để sửa!")
                with c_a2:
                    if st.button("🗑️ Xóa Bản Ghi (GDV)"):
                        if sel_mgr:
                            df_mgr.loc[df_mgr['SO SEQ'].isin(sel_mgr), 'TRANG THAI'] = 'Xóa'
                            df_mgr.to_excel(path_ds, index=False)
                            st.warning("⚠️ Đã xóa bản ghi.")
                            st.rerun()
                with c_a3:
                    if st.button("⚡ Thực Thi Hạch Toán Kho (XL)", type="primary"):
                        if sel_mgr:
                            for seq in sel_mgr:
                                row_item = df_mgr[df_mgr['SO SEQ'] == seq].iloc[0]
                                if str(row_item['TRANG THAI']) == 'Duyệt':
                                    ccmg_dict = {mg: float(row_item[mg] if mg in row_item else row_item.get(str(mg), 0)) for mg in DENOMINATIONS}
                                    execute_hach_toan_kho(path_gl, row_item, ccmg_dict)
                                    df_mgr.loc[df_mgr['SO SEQ'] == seq, ['TRANG THAI', 'NGUOI XL']] = ['XL', str(current_user)]
                            df_mgr.to_excel(path_ds, index=False)
                            st.success("🎉 Hạch toán kho thành công!")
                            st.rerun()

                # 🎯 FORM SỬA CỦA GDV: 4 CỘT TRÊN 1 HÀNG, SỐ TỜ RỘNG HƠN SỐ BÓ
                if 'sua_seq_gdv' in st.session_state and st.session_state['sua_seq_gdv']:
                    t_seq_g = st.session_state['sua_seq_gdv']
                    r_t_g = df_mgr[df_mgr['SO SEQ'] == t_seq_g].iloc[0]
                    st.markdown("---")
                    st.markdown(f"##### ✍️ [GDV] Điều Chỉnh Cơ Cấu Mệnh Giá (SEQ: `{t_seq_g}` - KH: `{r_t_g['TEN KH']}`)")
                    
                    new_ccmg_g = {}
                    new_tot_g = 0
                    
                    for i in range(0, len(DENOMINATIONS), 4):
                        row_cols_g = st.columns(4)
                        for col_idx, ji in enumerate(range(i, min(i + 4, len(DENOMINATIONS)))):
                            mg = DENOMINATIONS[ji]
                            k_n_g = f"gdv_edit_{t_seq_g}_{mg}"
                            if k_n_g not in st.session_state:
                                st.session_state[k_n_g] = f"{int(r_t_g.get(mg, r_t_g.get(str(mg), 0))):,}"
                            
                            with row_cols_g[col_idx]:
                                st.markdown(f"<p style='font-size:12px; font-weight:bold; margin-bottom:2px;'>{mg:,} VNĐ</p>", unsafe_allow_html=True)
                                sub_c1, sub_c2 = st.columns([3, 2])
                                with sub_c1:
                                    val_s = st.text_input(f"sg_{mg}", key=k_n_g, label_visibility="collapsed")
                                try: v_num = int(float(val_s.replace(",", "").strip())) if val_s else 0
                                except: v_num = 0
                                with sub_c2:
                                    st.number_input(f"bcg_{mg}", value=int(v_num // 1000), disabled=True, key=f"gdv_bc_{t_seq_g}_{mg}", label_visibility="collapsed")
                                new_ccmg_g[mg] = v_num
                                new_tot_g += v_num * mg
                        st.markdown("")

                    st.info(f"💰 **Tổng Giá Trị Sau Khi Sửa:** `{new_tot_g:,.0f}` VNĐ")
                    col_fg1, col_fg2 = st.columns(2)
                    with col_fg1:
                        if st.button("💾 Lưu Sửa Đổi (Chuyển Chờ Duyệt)", type="primary"):
                            for mg in DENOMINATIONS:
                                col_k = mg if mg in df_mgr.columns else (str(mg) if str(mg) in df_mgr.columns else None)
                                if col_k: df_mgr.loc[df_mgr['SO SEQ'] == t_seq_g, col_k] = float(new_ccmg_g[mg])
                                if f"{mg}_THANHTIEN" in df_mgr.columns: df_mgr.loc[df_mgr['SO SEQ'] == t_seq_g, f"{mg}_THANHTIEN"] = float(new_ccmg_g[mg] * mg)
                            df_mgr.loc[df_mgr['SO SEQ'] == t_seq_g, ['TONG THANH TIEN', 'TRANG THAI']] = [float(new_tot_g), 'Chờ Phê Duyệt']
                            df_mgr.to_excel(path_ds, index=False)
                            del st.session_state['sua_seq_gdv']
                            st.success("🎉 Lưu và chuyển chờ duyệt thành công!")
                            st.rerun()
                    with col_fg2:
                        if st.button("↩️ Hủy Bỏ Sửa"):
                            del st.session_state['sua_seq_gdv']
                            st.rerun()

    elif current_tab == "📊 Báo Cáo":
        st.markdown("#### 📊 Báo Cáo Giao Dịch & Tồn Quỹ Kho")
        bc_type = st.radio("Chọn loại báo cáo:", ["1. Báo cáo Giao dịch", "2. Báo cáo Tồn quỹ kho (GL)"])
        if bc_type.startswith("1") and os.path.exists(path_ds):
            st.dataframe(pd.read_excel(path_ds), use_container_width=True)
        elif os.path.exists(path_gl):
            wb_gl = openpyxl.load_workbook(path_gl)
            ws_gl = wb_gl.active
            data_gl = [[ws_gl.cell(row=r, column=c).value for c in range(1, ws_gl.max_column + 1)] for r in range(1, ws_gl.max_row + 1)]
            if len(data_gl) >= 2:
                st.dataframe(pd.DataFrame(data_gl[2:], columns=[str(h) for h in data_gl[1]]), use_container_width=True)
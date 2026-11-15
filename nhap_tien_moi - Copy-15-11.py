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
            ['990.1', 'KHO 194 TRẦN QUANG KHẢI'] + [0.0]*38
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

def execute_hach_toan_kho(path_gl, row_item, ccmg_data):
    """
    Hạch toán chuẩn xác vào file gl_tienmo.xlsx theo các trường hợp nghiệp vụ.
    Kiểm tra tồn kho khối ĐƠN VỊ của đơn vị xuất đối với nghiệp vụ Điều chuyển CN hoặc Nội bộ.
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

    idx_990 = get_row_index_wb('990')
    idx_dvxl = get_row_index_wb(don_vi_xl)
    idx_makh = get_row_index_wb(ma_kh)

    if idx_dvxl is None:
        return False, f"Lỗi: Đơn vị xử lý '{don_vi_xl}' không tồn tại trong Sổ cái gl_tienmo.xlsx!"

    def check_du_chi_wb(target_row, kho_block_map):
        if target_row is None:
            return False, f"Mã kho/tài khoản không tồn tại trong Sổ cái!"
        for mg in DENOMINATIONS:
            so_to_yc = float(ccmg_data.get(mg, 0) or 0)
            if so_to_yc > 0 and mg in kho_block_map:
                to_col = kho_block_map[mg]['to_col']
                val = ws.cell(row=target_row, column=to_col).value
                hien_co = float(val) if val is not None and str(val).strip() != '' else 0.0
                if hien_co < so_to_yc:
                    return False, f"Không Đủ Chi (Mệnh giá {mg:,} VNĐ không đủ tồn kho)!"
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

    # ==================== CÁC TRƯỜNG HỢP NGHIỆP VỤ ====================
    
    if thu_chi == "THU" and loai_nv == "Chi nhánh":
        if idx_makh is None:
            return False, f"Lỗi: Mã khách hàng/chi nhánh '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
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
        ok1, msg1 = check_du_chi_wb(idx_dvxl, tk_map)
        ok2, msg2 = check_du_chi_wb(idx_dvxl, dv_map)
        if not ok1 or not ok2:
            return False, "Không Đủ Chi (Đơn vị xuất không đủ tồn kho nội bộ)!"
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_makh, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
            
    elif thu_chi == "CHI" and loai_nv == "Chi nhánh":
        if idx_makh is None:
            return False, f"Lỗi: Mã chi nhánh '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        ok1, msg1 = check_du_chi_wb(idx_990, tk_map)
        ok2, msg2 = check_du_chi_wb(idx_makh, dv_map)
        if not ok1 or not ok2:
            return False, "Không Đủ Chi (Kho tổng hoặc Chi nhánh không đủ số dư)!"
        apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_makh, False, ccmg_data, dv_map, 40)
        
    elif thu_chi == "CHI" and loai_nv == "Cá Nhân":
        is_gdv_account = ".gdv" in don_vi_xl.lower()
        if is_gdv_account:
            ok, msg = check_du_chi_wb(idx_dvxl, dv_map)
            if not ok: return False, msg
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        else:
            ok1, msg1 = check_du_chi_wb(idx_990, tk_map)
            ok2, msg2 = check_du_chi_wb(idx_dvxl, tk_map)
            ok3, msg3 = check_du_chi_wb(idx_dvxl, dv_map)
            if not ok1 or not ok2 or not ok3:
                return False, "Không Đủ Chi (Kho quỹ không đủ tồn kho)!"
            apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_990, False, ccmg_data, dv_map, 40)
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
            
    elif thu_chi == "CHI" and loai_nv == "Tạm ứng":
        if idx_makh is None:
            return False, f"Lỗi: Mã GDV nhận tạm ứng '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        ok1, msg1 = check_du_chi_wb(idx_990, tk_map)
        ok2, msg2 = check_du_chi_wb(idx_dvxl, tk_map)
        ok3, msg3 = check_du_chi_wb(idx_dvxl, dv_map)
        if not ok1 or not ok2 or not ok3:
            return False, "Không Đủ Chi (Kho cơ sở không đủ tồn kho để tạm ứng)!"
        apply_delta_block_wb(idx_990, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
            
    elif thu_chi == "THU" and loai_nv == "Hoàn ứng":
        if idx_makh is None:
            return False, f"Lỗi: Mã GDV hoàn ứng '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        ok, msg = check_du_chi_wb(idx_makh, dv_map)
        if not ok: return False, msg
        apply_delta_block_wb(idx_990, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_990, True, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
        apply_delta_block_wb(idx_dvxl, True, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_makh, False, ccmg_data, dv_map, 40)
        
    elif loai_nv == "Điều chuyển CN":
        if idx_makh is None:
            return False, f"Lỗi: Mã chi nhánh nhận điều chuyển '{ma_kh}' phải có sẵn trong Sổ cái gl_tienmo.xlsx!"
        
        ok_dv, msg_dv = check_du_chi_wb(idx_dvxl, dv_map)
        if not ok_dv:
            return False, f"Không Đủ Chi (Khối Đơn Vị của đơn vị xuất '{don_vi_xl}' không đủ số dư để điều chuyển): {msg_dv}"
            
        apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        apply_delta_block_wb(idx_makh, True, ccmg_data, dv_map, 40)
        
    elif loai_nv == "Điều chỉnh kiểm kê":
        if thu_chi == "THU":
            apply_delta_block_wb(idx_dvxl, True, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, True, ccmg_data, dv_map, 40)
        else:
            ok, msg = check_du_chi_wb(idx_dvxl, tk_map)
            if not ok: return False, msg
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, tk_map, 21)
            apply_delta_block_wb(idx_dvxl, False, ccmg_data, dv_map, 40)
        
    else:
        apply_delta_block_wb(idx_dvxl, thu_chi == "THU", ccmg_data, dv_map, 40)

    wb.save(path_gl)
    return True, "Hạch toán kho thành công!"

def show_nhap_tien_moi_page(base_dir, shared_data_dir, current_user="System", user_role="NV"):
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
    
    is_giam_doc = str(user_role).strip().upper() == "GIÁM ĐỐC"
    
    coso_options = ["CS1 - Cơ sở 1", "CS2 - Cơ sở 2", "990 - Kho Tổng"]
    if os.path.exists(path_coso):
        try:
            df_cs = pd.read_excel(path_coso)
            if 'Kho' in df_cs.columns and 'Tên_cơ_sở' in df_cs.columns:
                coso_options = [f"{row['Kho']} - {row['Tên_cơ_sở']}" for _, row in df_cs.iterrows()]
        except:
            pass

    gl_accounts = get_list_accounts_from_gl(path_gl)
    coso_accounts = get_list_accounts_from_coso(path_coso)

    # --- KHỞI TẠO TRẠNG THÁI SESSION_STATE ---
    if 'bind_cccd' not in st.session_state: st.session_state['bind_cccd'] = ""
    if 'bind_ten' not in st.session_state: st.session_state['bind_ten'] = ""
    if 'bind_dc' not in st.session_state: st.session_state['bind_dc'] = ""
    if 'is_nhap_moi_active' not in st.session_state: st.session_state['is_nhap_moi_active'] = False

    for mg in DENOMINATIONS:
        if f"to_{mg}" not in st.session_state: st.session_state[f"to_{mg}"] = 0
        if f"bo_{mg}" not in st.session_state: st.session_state[f"bo_{mg}"] = 0.0
        if f"bao_{mg}" not in st.session_state: st.session_state[f"bao_{mg}"] = 0.0

    tabs_list = [
        "📥 Giao Dịch Thu/Chi", "🔄 Ứng / Hoàn Tiền", "🛡️ Phê Duyệt (LĐ)", "🔍 Tìm Kiếm & Xử Lý", "📊 Báo Cáo"
    ]
    
    current_tab = st.radio("Chuyển Tab Chức Năng", tabs_list, horizontal=True, label_visibility="collapsed")

    if current_tab == "📥 Giao Dịch Thu/Chi":
        if str(user_role).strip().upper() in ["GIÁM ĐỐC", "LĐTT", "ADMIN"]:
            st.warning("⚠️ Tài khoản nhóm quản lý cấp cao không có quyền lập hoặc nhập liệu giao dịch.")
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
                        if 'gd_select_kh_chung' in st.session_state:
                            st.session_state['gd_select_kh_chung'] = "-- Chọn khách hàng từ danh sách hoặc nhập mới --"
                        for mg in DENOMINATIONS:
                            st.session_state[f"to_{mg}"] = 0
                            st.session_state[f"bo_{mg}"] = 0.0
                            st.session_state[f"bao_{mg}"] = 0.0
                        st.rerun()
                else:
                    if st.button("🔄 Đang Nhập Mới", type="secondary", use_container_width=True):
                        st.session_state['bind_cccd'] = ""
                        st.session_state['bind_ten'] = ""
                        st.session_state['bind_dc'] = ""
                        if 'gd_select_kh_chung' in st.session_state:
                            st.session_state['gd_select_kh_chung'] = "-- Chọn khách hàng từ danh sách hoặc nhập mới --"
                        for mg in DENOMINATIONS:
                            st.session_state[f"to_{mg}"] = 0
                            st.session_state[f"bo_{mg}"] = 0.0
                            st.session_state[f"bao_{mg}"] = 0.0
                        st.success("✨ Đã làm mới màn hình nhập liệu!")
                        st.rerun()

            if not st.session_state['is_nhap_moi_active']:
                st.info("ℹ️ Vui lòng bấm vào nút **'➕ Nhập Mới'** ở góc trên bên phải để bắt đầu tạo giao dịch.")
            else:
                col1, col2, col3 = st.columns(3)
                with col1:
                    ngay_hieu_luc = st.date_input("Ngày Hiệu Lực", value=date.today(), key="gd_ngay")
                    loai_nv_options = ["Cá Nhân", "Chi nhánh", "NHNN", "Nội bộ", "Tạm ứng", "Hoàn ứng", "Điều chuyển CN", "Điều chỉnh kiểm kê"]
                    loai_nv = st.selectbox("Loại Nghiệp Vụ", loai_nv_options, index=0, key="gd_loainv")
                with col2:
                    chon_coso = st.selectbox("Đơn Vị Xử Lý (Kho)", coso_options, key="gd_donvixl")
                    don_vi_xl = chon_coso.split(" - ")[0].strip()
                    
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
                        diachi_kh = st.text_input("Địa Chỉ (Không bắt buộc)", key="gd_diachi_coso")
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
                        diachi_kh = st.text_input("Địa Chỉ KH (Không bắt buộc)", key="gd_diachi_other")
                        ten_kh = st.text_input("Tên Khách Hàng / Đơn Vị", value=default_ten_kh.upper(), max_chars=30, key="gd_tenkh")
                    else:
                        df_kh_all = pd.DataFrame(columns=['MA CCCD', 'TEN KH', 'DIACHI KH'])
                        if os.path.exists(path_kh):
                            try:
                                df_kh_all = pd.read_excel(path_kh)
                                df_kh_all.columns = [str(c).strip().upper() for c in df_kh_all.columns]
                                for std_c in ['MA CCCD', 'TEN KH', 'DIACHI KH']:
                                    if std_c not in df_kh_all.columns:
                                        for alt in df_kh_all.columns:
                                            if std_c.replace(" ", "") in alt.replace(" ", ""):
                                                df_kh_all.rename(columns={alt: std_c}, inplace=True)
                                                break
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
                            "🔍 Tìm và Chọn Khách Hàng (Tự động điền dữ liệu)",
                            list(kh_options_dict.keys()),
                            key="gd_select_kh_chung",
                            on_change=on_change_chon_kh_v2
                        )

                        ma_kh = st.text_input("Mã CCCD / Mã KH", key="bind_cccd")
                        ten_kh = st.text_input("Tên Khách Hàng (In hoa)", max_chars=30, key="bind_ten")
                        diachi_kh = st.text_input("Địa Chỉ KH", key="bind_dc")

                st.markdown("---")
                st.markdown("##### 🔢 Nhập Số Tờ / Số Lượng Chi Tiết Theo Mệnh Giá")
                
                unit_mode = st.radio("Chọn đơn vị tính nhập liệu:", ["Số tờ (Step 100 / Tự do)", "Theo Bó (1 Bó = 1,000 tờ)", "Theo Bao (1 Bao = 20,000 tờ)"], horizontal=True, key="gd_unit_mode")
                
                mg_cols = st.columns(5)
                ccmg_data = {}
                tong_thanh_tien = 0
                
                for idx, mg in enumerate(DENOMINATIONS):
                    with mg_cols[idx % 5]:
                        if unit_mode.startswith("Số tờ"):
                            so_luong_nhap = st.number_input(f"Mệnh giá {mg:,} VNĐ", min_value=0, step=100, format="%d", key=f"to_{mg}")
                            so_to = int(so_luong_nhap)
                        elif unit_mode.startswith("Theo Bó"):
                            so_bo = st.number_input(f"Mệnh giá {mg:,} (Số Bó)", min_value=0.0, step=1.0, format="%.2f", key=f"bo_{mg}")
                            so_to = int(so_bo * 1000)
                        else:
                            so_bao = st.number_input(f"Mệnh giá {mg:,} (Số Bao)", min_value=0.0, step=0.5, format="%.2f", key=f"bao_{mg}")
                            so_to = int(so_bao * 20000)
                            
                        ccmg_data[mg] = so_to
                        tong_thanh_tien += so_to * mg

                st.info(f"💰 **Tổng Giá Trị Giao Dịch:** `{tong_thanh_tien:,.0f}` VNĐ")
                
                if st.button("💾 Cập Nhật Bản Ghi (Chờ Phê Duyệt)", type="primary"):
                    if not ma_kh.strip() and loai_nv != "Điều chỉnh kiểm kê":
                        st.error("❌ Mã CCCD / Mã KH / Mã Đơn Vị nhận không được để trống!")
                    elif not ten_kh.strip():
                        st.error("❌ Tên khách hàng / Đơn vị nhận không được để trống!")
                    elif tong_thanh_tien <= 0:
                        st.error("❌ Vui lòng nhập số lượng tiền giao dịch lớn hơn 0!")
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
                            st.success(f"🎉 Đã lập bản ghi giao dịch thành công (Mã SEQ: `{next_seq}`)! Vui lòng bấm nút '➕ Nhập Mới' ở trên nếu muốn tạo giao dịch tiếp theo.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Lỗi lưu giao dịch: {e}")

    elif current_tab == "🔄 Ứng / Hoàn Tiền":
        st.markdown("#### 🔄 Giao Dịch Tạm Ứng / Hoàn Ứng Nhanh")
        st.info("Chức năng đồng bộ tương tự tab Giao dịch với nghiệp vụ Tạm ứng hoặc Hoàn ứng tự động.")

    elif current_tab == "🛡️ Phê Duyệt (LĐ)":
        st.markdown("#### 🛡️ Phê Duyệt Giao Dịch & Sửa Cơ Cấu Mệnh Giá (Dành cho Lãnh Đạo / GIÁM ĐỐC)")
        
        if not is_giam_doc:
            st.warning("⚠️ Tài khoản của bạn không có thẩm quyền phê duyệt. Chức năng này chỉ dành riêng cho nhóm quyền **GIÁM ĐỐC**.")
        else:
            if os.path.exists(path_ds):
                df_app = pd.read_excel(path_ds)
                if 'TRANG THAI' not in df_app.columns:
                    df_app['TRANG THAI'] = 'Chờ Phê Duyệt'
                    
                df_cho = df_app[df_app['TRANG THAI'].astype(str).str.strip() == 'Chờ Phê Duyệt'].copy()
                
                if df_cho.empty:
                    st.info("ℹ️ Hiện không có bản ghi nào chờ phê duyệt.")
                else:
                    st.markdown("##### 📋 Chọn giao dịch cần duyệt hoặc xem chi tiết cấu trúc mệnh giá bên dưới:")
                    
                    cols_to_show_app = ['SO SEQ', 'TEN KH', 'LOAI NV', 'TRANG THAI']
                    for mg in DENOMINATIONS:
                        if mg in df_cho.columns:
                            cols_to_show_app.append(mg)
                        elif str(mg) in df_cho.columns:
                            cols_to_show_app.append(str(mg))
                    if 'TONG THANH TIEN' in df_cho.columns:
                        cols_to_show_app.append('TONG THANH TIEN')
                        
                    valid_view_cols_app = [c for c in cols_to_show_app if c in df_cho.columns]
                    df_cho_display = df_cho[valid_view_cols_app].copy()
                    
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
                        key="editor_pheduyet_list_v73"
                    )
                    
                    selected_rows_app = edited_list_app[edited_list_app['Chọn'] == True]
                    selected_seqs_app = selected_rows_app['SO SEQ'].tolist()
                    
                    st.markdown("---")
                    col_b1, col_b2, col_b3, col_b4 = st.columns(4)
                    with col_b1:
                        if st.button("✅ Phê Duyệt Nhanh", type="primary", key="btn_ld_duyet_nhanh"):
                            if not selected_seqs_app:
                                st.warning("⚠️ Vui lòng tích chọn ít nhất một giao dịch!")
                            else:
                                try:
                                    df_app['TRANG THAI'] = df_app['TRANG THAI'].astype(str)
                                    df_app['NGUOI PHE DUYET'] = df_app['NGUOI PHE DUYET'].astype(str)
                                    df_app.loc[df_app['SO SEQ'].isin(selected_seqs_app), 'TRANG THAI'] = 'Duyệt'
                                    df_app.loc[df_app['SO SEQ'].isin(selected_seqs_app), 'NGUOI PHE DUYET'] = str(current_user)
                                    
                                    for col in df_app.columns:
                                        if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                            df_app[col] = pd.to_numeric(df_app[col], errors='coerce').fillna(0.0)

                                    df_app.to_excel(path_ds, index=False)
                                    st.success(f"🎉 Đã phê duyệt thành công {len(selected_seqs_app)} giao dịch!")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Lỗi phê duyệt: {ex}")
                    with col_b2:
                        if st.button("✏️ Sửa Cơ Cấu Mệnh Giá", key="btn_ld_mo_sua_ccmg"):
                            if len(selected_seqs_app) != 1:
                                st.warning("⚠️ Vui lòng tích chọn **chính xác 1 dòng duy nhất** để mở form sửa!")
                            else:
                                target_s = selected_seqs_app[0]
                                st.session_state['sua_seq_dang_chon_ld'] = target_s
                                # Khởi tạo session state lưu tạm giá trị text có dấu phẩy cho form sửa LĐ
                                r_init = df_cho[df_cho['SO SEQ'] == target_s].iloc[0]
                                for mg in DENOMINATIONS:
                                    val_mg = 0
                                    if mg in r_init and pd.notna(r_init[mg]):
                                        val_mg = int(r_init[mg])
                                    elif str(mg) in r_init and pd.notna(r_init[str(mg)]):
                                        val_mg = int(r_init[str(mg)])
                                    st.session_state[f"input_ld_{target_s}_{mg}"] = f"{val_mg:,}"
                                st.rerun()
                    with col_b3:
                        if st.button("🗑️ Xóa Bản Ghi (Giám Đốc)", key="btn_ld_xoa"):
                            if not selected_seqs_app:
                                st.warning("⚠️ Vui lòng tích chọn bản ghi cần xóa!")
                            else:
                                try:
                                    df_app.loc[df_app['SO SEQ'].isin(selected_seqs_app), 'TRANG THAI'] = 'Xóa'
                                    for col in df_app.columns:
                                        if col in DENOMINATIONS or str(col).isdigit() or 'THANH TIEN' in str(col) or 'TONG' in str(col) or col == 'SO SEQ':
                                            df_app[col] = pd.to_numeric(df_app[col], errors='coerce').fillna(0.0)
                                    df_app.to_excel(path_ds, index=False)
                                    st.warning(f"⚠️ Giám đốc đã xóa {len(selected_seqs_app)} bản ghi.")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Lỗi xóa: {ex}")
                    with col_b4:
                        if st.button("❌ Từ Chối", key="btn_ld_tuchoi"):
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
                                    st.warning(f"⚠️ Đã từ chối {len(selected_seqs_app)} giao dịch.")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Lỗi từ chối: {ex}")

                    if 'sua_seq_dang_chon_ld' in st.session_state and st.session_state['sua_seq_dang_chon_ld'] is not None:
                        target_seq = st.session_state['sua_seq_dang_chon_ld']
                        row_target_df = df_app[df_app['SO SEQ'] == target_seq]
                        
                        if not row_target_df.empty:
                            r_val = row_target_df.iloc[0]
                            st.markdown("---")
                            st.markdown(f"#### ✍️ [GIÁM ĐỐC] Điều Chỉnh Cơ Cấu Mệnh Giá Cho Giao Dịch (Mã SEQ: `{target_seq}` - KH: `{r_val['TEN KH']}`)")
                            
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
                                        f"Mệnh giá {mg:,} VNĐ", 
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

                            st.info(f"💰 **Tổng Giá Trị Sau Khi Sửa:** `{new_total_val:,.0f}` VNĐ")
                            
                            col_f1, col_f2 = st.columns(2)
                            with col_f1:
                                if st.button("💾 Lưu Sửa Đổi & Phê Duyệt Ngay", type="primary", key="btn_ld_save_approve"):
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
                                        st.success(f"🎉 Đã cập nhật cơ cấu mệnh giá và phê duyệt thành công giao dịch SEQ {target_seq}!")
                                        st.rerun()
                                    except Exception as ex:
                                        st.error(f"❌ Lỗi: {ex}")
                            with col_f2:
                                if st.button("↩️ Hủy Bỏ", key="btn_ld_cancel_ccmg"):
                                    del st.session_state['sua_seq_dang_chon_ld']
                                    st.rerun()

    elif current_tab == "🔍 Tìm Kiếm & Xử Lý":
        st.markdown("#### 🔍 Xử Lý Hạch Toán Kho (XL) & Quản Lý Bản Ghi (Nhân Viên)")
        
        if is_giam_doc:
            st.warning("⚠️ Tài khoản nhóm quyền GIÁM ĐỐC chỉ có thẩm quyền phê duyệt.")
        else:
            if os.path.exists(path_ds):
                df_mgr = pd.read_excel(path_ds)
                if 'TRANG THAI' not in df_mgr.columns:
                    df_mgr['TRANG THAI'] = 'Chờ Phê Duyệt'
                
                df_editable_list = df_mgr[df_mgr['TRANG THAI'].astype(str).str.strip().isin(['Chờ Phê Duyệt', 'Từ Chối Phê Duyệt', 'Duyệt', 'Xóa'])].copy()
                
                if df_editable_list.empty:
                    st.info("ℹ️ Hiện không có bản ghi giao dịch nào cần xử lý.")
                else:
                    st.markdown("##### ✏️ Quản Lý Xử Lý Hạch Toán Kho (Chỉ hiện các cột TEN KH, LOAI NV, TRANG THAI, các mệnh giá và TONG THANH TIEN):")
                    
                    cols_to_show_base = ['SO SEQ', 'TEN KH', 'LOAI NV', 'TRANG THAI']
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
                        key="editor_nhanvien_suaxoa_v73"
                    )
                    
                    selected_rows_mgr = edited_df_mgr[edited_df_mgr['Chọn'] == True]
                    selected_seqs_xl = selected_rows_mgr['SO SEQ'].tolist()
                    
                    st.markdown("---")
                    col_act1, col_act2, col_act3 = st.columns(3)
                    with col_act1:
                        if st.button("✏️ Sửa Cơ Cấu Mệnh Giá (GDV)", key="btn_gdv_mo_sua_ccmg"):
                            if len(selected_seqs_xl) != 1:
                                st.warning("⚠️ Vui lòng tích chọn **chính xác 1 dòng duy nhất** để mở form sửa!")
                            else:
                                target_seq_gdv = selected_seqs_xl[0]
                                r_check = df_mgr.loc[df_mgr['SO SEQ'] == target_seq_gdv, 'TRANG THAI'].values[0]
                                if str(r_check).strip() in ['Chờ Phê Duyệt', 'Từ Chối Phê Duyệt']:
                                    st.session_state['sua_seq_dang_chon_gdv'] = target_seq_gdv
                                    # Khởi tạo session state lưu tạm giá trị text có dấu phẩy cho form sửa GDV
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
                        if st.button("🗑️ Xóa Bản Ghi (GDV)", key="btn_xoa_nv"):
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
                                        st.success(f"⚠️ Đã chuyển các bản ghi sang trạng thái **Xóa** thành công.")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Lỗi xóa: {ex}")
                    with col_act3:
                        if st.button("⚡ Thực Thi Hạch Toán Kho (XL) Dòng Đã Duyệt", type="primary", key="btn_xl_compact"):
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
                                                ccmg_dict = {}
                                                for mg in DENOMINATIONS:
                                                    val_to = 0
                                                    if mg in row_data_item:
                                                        val_to = row_data_item[mg]
                                                    elif str(mg) in row_data_item:
                                                        val_to = row_data_item[str(mg)]
                                                    ccmg_dict[mg] = float(val_to) if pd.notna(val_to) else 0.0
                                                
                                                success, msg = execute_hach_toan_kho(path_gl, row_data_item, ccmg_dict)
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
                                    if count_success > 0:
                                        st.success(f"🎉 Đã hạch toán kho thành công {count_success} giao dịch, cập nhật số dư vào file `gl_tienmo.xlsx` và chuyển trạng thái thành `XL`!")
                                    if err_messages:
                                        for err in err_messages:
                                            st.error(err)
                                    if count_success > 0 or err_messages:
                                        st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Lỗi thực hiện hạch toán kho: {ex}")

                    if 'sua_seq_dang_chon_gdv' in st.session_state and st.session_state['sua_seq_dang_chon_gdv'] is not None:
                        target_seq_gdv = st.session_state['sua_seq_dang_chon_gdv']
                        row_target_gdv = df_mgr[df_mgr['SO SEQ'] == target_seq_gdv]
                        
                        if not row_target_gdv.empty:
                            r_val_g = row_target_gdv.iloc[0]
                            st.markdown("---")
                            st.markdown(f"#### ✍️ [GDV] Điều Chỉnh Cơ Cấu Mệnh Giá Cho Giao Dịch (Mã SEQ: `{target_seq_gdv}` - KH: `{r_val_g['TEN KH']}`)")
                            
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
                                        f"Mệnh giá {mg:,} VNĐ", 
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

                            st.info(f"💰 **Tổng Giá Trị Sau Khi Sửa:** `{new_total_gdv:,.0f}` VNĐ")
                            
                            col_fg1, col_fg2 = st.columns(2)
                            with col_fg1:
                                if st.button("💾 Lưu Sửa Đổi (Chuyển về Chờ Phê Duyệt)", type="primary", key="btn_gdv_save_ccmg"):
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
                                        st.success(f"🎉 Đã cập nhật cơ cấu mệnh giá và chuyển giao dịch SEQ {target_seq_gdv} về trạng thái **Chờ Phê Duyệt** thành công!")
                                        st.rerun()
                                    except Exception as ex:
                                        st.error(f"❌ Lỗi: {ex}")
                            with col_fg2:
                                if st.button("↩️ Hủy Bỏ", key="btn_gdv_cancel_ccmg"):
                                    del st.session_state['sua_seq_dang_chon_gdv']
                                    st.rerun()

    elif current_tab == "📊 Báo Cáo":
        st.markdown("#### 📊 Báo Cáo Giao Dịch & Tồn Quỹ Kho")
        bc_type = st.radio("Chọn loại báo cáo:", ["1. Báo cáo Giao dịch tiền mới", "2. Báo cáo Tồn quỹ kho (GL)"])
        
        if bc_type.startswith("1"):
            st.markdown("##### 🔍 Bộ Lọc Dữ Liệu Báo Cáo Giao Dịch")
            if os.path.exists(path_ds):
                df_bc1 = pd.read_excel(path_ds)
                
                f_col1, f_col2, f_col3 = st.columns(3)
                with f_col1:
                    min_dt = pd.to_datetime(df_bc1['NGAY HIEU LUC'], errors='coerce').min()
                    max_dt = pd.to_datetime(df_bc1['NGAY HIEU LUC'], errors='coerce').max()
                    if pd.isna(min_dt): min_dt = date.today()
                    if pd.isna(max_dt): max_dt = date.today()
                    
                    tu_ngay = st.date_input("Từ Ngày", value=min_dt, key="bc_tungay")
                with f_col2:
                    den_ngay = st.date_input("Đến Ngày", value=max_dt, key="bc_denngay")
                with f_col3:
                    all_dv = ["TẤT CẢ"] + sorted(list(df_bc1['DON VI XL'].dropna().astype(str).unique()))
                    chon_dv = st.selectbox("Đơn Vị Xử Lý (DON VI XL)", all_dv, key="bc_donvi")
                
                f_col4, _ = st.columns(2)
                with f_col4:
                    all_tt = ["TẤT CẢ"] + sorted(list(df_bc1['TRANG THAI'].dropna().astype(str).unique()))
                    chon_tt = st.selectbox("Trạng Thái (TRANG THAI)", all_tt, key="bc_trangthai")
                
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
                
                st.markdown(f"##### 📋 Danh Sách Giao Dịch (Hiển thị {len(df_filtered_bc1)} bản ghi sau lọc)")
                
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
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
                    st.markdown("###### 📊 Bảng Số Dư Tồn Quỹ (Định dạng dấu phẩy hàng ngàn thông minh):")
                    
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
                        label="📥 Tải Xuống Báo Cáo Tồn Quỹ Kho (Excel .xlsx Định Dạng Chuẩn)",
                        data=gl_excel_data,
                        file_name=f"BaoCao_TonQuyKho_GL_{date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
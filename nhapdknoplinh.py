import streamlit as st
import pandas as pd
import os
from datetime import datetime, date, timedelta, timezone

def show_dang_ky_nop_linh_page(base_dir, shared_data_dir, current_user="System"):
    
    st.subheader("📝 QUẢN LÝ ĐĂNG KÝ NỘP LĨNH TIỀN MẶT & SỐ DƯ KHẢ DỤNG")
    st.markdown(f"👤 **Cán bộ thao tác hiện tại:** `{current_user}`")
    
    param_dir = os.path.join(base_dir, "parameters")
    path_thamso_km = os.path.join(param_dir, "Thamso_KM.xlsx")
    path_thamso_tiente = os.path.join(param_dir, "Thamso_tiente.xlsx")
    path_data_ds = os.path.join(shared_data_dir, "Ds_dangkynoplinh.xlsx")
    
    os.makedirs(shared_data_dir, exist_ok=True)
    os.makedirs(param_dir, exist_ok=True)

    TZ_VN = timezone(timedelta(hours=7))

    # 1. Danh sách Cơ sở cố định
    list_coso = ["Cơ sở 1", "Cơ sở 2", "CS HĐ", "CS HP", "Hub NVC"]

    # 2. Đọc danh sách Khách hàng từ Thamso_KM.xlsx
    list_kh = []
    if os.path.exists(path_thamso_km):
        try:
            df_km = pd.read_excel(path_thamso_km)
            col_kh = next((c for c in df_km.columns if 'tenkhachhang' in str(c).lower() or 'khách hàng' in str(c).lower()), df_km.columns[0])
            list_kh = df_km[col_kh].dropna().astype(str).tolist()
        except Exception:
            pass
    if not list_kh: list_kh = ["CN Ba Đình", "PGD Hoàn Kiếm", "CN Đống Đa"]

    # 3. Đọc danh sách Tiền tệ từ Thamso_tiente.xlsx
    list_tiente = []
    if os.path.exists(path_thamso_tiente):
        try:
            df_tt = pd.read_excel(path_thamso_tiente)
            col_tt = next((c for c in df_tt.columns if 'loai_tien' in str(c).lower() or 'tiền' in str(c).lower() or 'tiente' in str(c).lower()), df_tt.columns[0])
            list_tiente = df_tt[col_tt].dropna().astype(str).tolist()
        except Exception:
            pass
    if not list_tiente: list_tiente = ["VND", "USD", "EUR", "JPY"]

    # Khởi tạo file dữ liệu nếu chưa có
    if not os.path.exists(path_data_ds):
        try:
            base_cols = [
                'Cơ sở', 'Ngày lập', 'Ngày hiệu lực', 'Giờ Nộp Rút', 
                'Tên khách hàng', 'Loại GD', 'Số tiền', 'Loại tiền', 'Trạng thái',
                'NguoiNhap', 'NguoiXL', 'Loại bản ghi', 'Tổng Khả Dụng'
            ]
            df_init = pd.DataFrame(columns=base_cols)
            df_init.to_excel(path_data_ds, index=False)
        except Exception as e:
            st.error(f"❌ Không thể tạo file dữ liệu: {e}")

    action = st.radio(
        "Chọn chức năng:", 
        ["Thêm mới giao dịch", "Nhập Số Dư Khả Dụng", "Xóa bản ghi (Chờ XL)", "Hoàn tất GD (Chờ XL)", "Báo cáo Đăng ký Nộp Lĩnh"], 
        horizontal=True
    )
    st.markdown("---")

    # ---------------- 1. THÊM MỚI GIAO DỊCH ----------------
    if action == "Thêm mới giao dịch":
        st.markdown("#### ➕ Nhập liệu đăng ký Nộp / Lĩnh mới")
        
        if 'form_data' not in st.session_state: st.session_state.form_data = None
        if 'confirm_ghi' not in st.session_state: st.session_state.confirm_ghi = False

        col1, col2 = st.columns(2)
        with col1:
            t_coso = st.selectbox("1. Tên Cơ sở", list_coso, key="inp_coso")
            t_ngay_dk = datetime.now(TZ_VN).date()
            st.text(f"2. Ngày Đăng ký: {t_ngay_dk.strftime('%d/%m/%Y')}")
            
            t_ngay_hl = st.date_input("3. Ngày hiệu lực", value=t_ngay_dk, key="inp_ngay_hl")
            
            t_gio_nop_rut_str = st.text_input("4. Giờ Nộp / Rút (Gõ liền số, vd: 930 hoặc 1430)", value="", key="inp_gio", placeholder="Ví dụ: 930")
            
            raw_g = ''.join([c for c in t_gio_nop_rut_str if c.isdigit()])
            if len(raw_g) == 3:
                formatted_gio = f"0{raw_g[0]}:{raw_g[1:]}:00"
                h, m, s = 0, int(raw_g[0]), int(raw_g[1:])
            elif len(raw_g) == 4:
                formatted_gio = f"{raw_g[:2]}:{raw_g[2:]}:00"
                h, m, s = int(raw_g[:2]), int(raw_g[2:4]), 0
            elif len(raw_g) >= 6:
                formatted_gio = f"{raw_g[:2]}:{raw_g[2:4]}:{raw_g[4:6]}"
                h, m, s = int(raw_g[:2]), int(raw_g[2:4]), int(raw_g[4:6])
            else:
                formatted_gio = "08:30:00"
                h, m, s = 8, 30, 0
            
            st.caption(f"💡 Giờ nhận diện tự động: **{formatted_gio}**")
            
            st.markdown("---")
            st.markdown("**5. Chọn Tên CN/PGD:**")
            search_box = st.text_input("🔍 Gõ từ khóa lọc nhanh PGD:", "", key="pgd_search_input")
            
            filtered_kh = [kh for kh in list_kh if search_box.lower() in kh.lower()] if search_box.strip() else list_kh
            if not filtered_kh: filtered_kh = ["Không tìm thấy kết quả phù hợp"]
            
            t_kh = st.selectbox(f"Kết quả lọc ({len(filtered_kh)} PGD):", filtered_kh, label_visibility="collapsed", key="inp_kh")
            st.markdown("---")
            
        with col2:
            t_loai_gd = st.selectbox("6. Loại giao dịch", ["Nộp", "Lĩnh"], key="inp_loai_gd")
            t_loai_tien = st.selectbox("7. Loại tiền", list_tiente, key="inp_tiente")
            
            raw_sotien_input = st.text_input("8. Số tiền", value="0", key="inp_sotien")
            
            try:
                clean_chars = ''.join([c for c in raw_sotien_input if c.isdigit()])
                numeric_val = float(clean_chars) if clean_chars else 0.0
                formatted_display = f"{numeric_val:,.0f}"
            except:
                numeric_val = 0.0
                formatted_display = "0"
            
            st.markdown(f"💡 **Số tiền hiển thị chính xác:** <span style='color: #D4849D; font-size: 1.2em; font-weight: bold;'>{formatted_display}</span>", unsafe_allow_html=True)
            
        submitted = st.button("💾 Ghi dữ liệu", type="primary", use_container_width=True)
        
        if submitted:
            try:
                ngay_gio_hl_dt = datetime.combine(t_ngay_hl, datetime.min.time().replace(hour=h, minute=m, second=s))
                
                st.session_state.form_data = {
                    'Cơ sở': t_coso,
                    'Ngày lập': t_ngay_dk.strftime('%Y-%m-%d'),
                    'Ngày hiệu lực': ngay_gio_hl_dt,
                    'Giờ Nộp Rút': formatted_gio,
                    'Tên khách hàng': t_kh,
                    'Loại GD': t_loai_gd,
                    'Số tiền': numeric_val,
                    'Loại tiền': t_loai_tien,
                    'Trạng thái': 'Chờ XL',
                    'NguoiNhap': str(current_user),
                    'NguoiXL': str(current_user),
                    'Loại bản ghi': 'GiaoDich'
                }
                st.session_state.confirm_ghi = True
            except Exception as e:
                st.error(f"❌ Lỗi dữ liệu hoặc định dạng giờ: {e}")

        if st.session_state.get('confirm_ghi', False) and st.session_state.get('form_data') is not None:
            st.warning("⚠️ **Bạn có chắc chắn muốn ghi bản ghi đăng ký nộp lĩnh này vào hệ thống không?**")
            
            c_yes, c_no = st.columns(2)
            with c_yes:
                if st.button("✅ Yes (Đồng ý ghi)", use_container_width=True, key="btn_yes_ghi"):
                    try:
                        new_row = st.session_state.form_data
                        df_current = pd.read_excel(path_data_ds) if os.path.exists(path_data_ds) else pd.DataFrame(columns=new_row.keys())
                        df_current = pd.concat([df_current, pd.DataFrame([new_row])], ignore_index=True)
                        df_current.to_excel(path_data_ds, index=False)
                        
                        st.success(f"🎉 Ghi dữ liệu thành công bởi `{current_user}`! Trạng thái: [Chờ XL]")
                        st.session_state.confirm_ghi = False
                        st.session_state.form_data = None
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Lỗi khi ghi file: {e}")
            with c_no:
                if st.button("❌ No (Hủy bỏ)", use_container_width=True, key="btn_no_ghi"):
                    st.info("ℹ️ Đã hủy thao tác ghi dữ liệu.")
                    st.session_state.confirm_ghi = False
                    st.session_state.form_data = None
                    st.rerun()

    # ---------------- 2. NHẬP SỐ DƯ KHẢ DỤNG (QUÉT TÌM VÀ GHI ĐÈ) ----------------
    elif action == "Nhập Số Dư Khả Dụng":
        st.markdown("#### 💰 Màn Hình Nhập Số Dư Khả Dụng")
        
        with st.container(border=True):
            col_sk1, col_sk2, col_sk3 = st.columns(3)
            with col_sk1:
                sd_ngay_hl = st.date_input("📅 Ngày hiệu lực", value=datetime.now(TZ_VN).date(), key="sd_ngay_hl")
            with col_sk2:
                sd_coso = st.selectbox("🏢 Cơ sở", list_coso, key="sd_coso")
            with col_sk3:
                sd_tiente = st.selectbox("💵 Loại tiền", list_tiente, key="sd_tiente")

            st.markdown("---")
            raw_sodu_input = st.text_input("💵 Nhập tổng số tiền Số Dư Khả Dụng", value="0", key="sd_sotien_input")
            
            try:
                clean_sodu_chars = ''.join([c for c in raw_sodu_input if c.isdigit()])
                numeric_sodu = float(clean_sodu_chars) if clean_sodu_chars else 0.0
                formatted_sodu_disp = f"{numeric_sodu:,.0f}"
            except:
                numeric_sodu = 0.0
                formatted_sodu_disp = "0"
            
            st.markdown(f"💡 **Số dư khả dụng xác nhận:** <span style='color: #2E8B57; font-size: 1.3em; font-weight: bold;'>{formatted_sodu_disp}</span>", unsafe_allow_html=True)

        if st.button("💾 Lưu Số Dư Khả Dụng", type="primary", use_container_width=True):
            try:
                df_all = pd.read_excel(path_data_ds) if os.path.exists(path_data_ds) else pd.DataFrame()
                date_str = sd_ngay_hl.strftime('%Y-%m-%d')
                
                # Quét tìm xem đã tồn tại số dư khả dụng của Cơ sở + Ngày hiệu lực + Loại tiền chưa
                found_idx = -1
                if not df_all.empty and 'Loại bản ghi' in df_all.columns:
                    df_all['Tmp_Date'] = pd.to_datetime(df_all['Ngày hiệu lực'], errors='coerce').dt.strftime('%Y-%m-%d')
                    for idx, row in df_all.iterrows():
                        if (str(row.get('Loại bản ghi', '')) == 'SoDuKhaDung' and
                            str(row.get('Cơ sở', '')).strip() == str(sd_coso).strip() and
                            str(row.get('Tmp_Date', '')) == date_str and
                            str(row.get('Loại tiền', '')).strip() == str(sd_tiente).strip()):
                            found_idx = idx
                            break
                    df_all = df_all.drop(columns=['Tmp_Date'], errors='ignore')

                if found_idx != -1:
                    # Đã có -> GHI ĐÈ thay thế trực tiếp giá trị số dư
                    df_all.loc[found_idx, 'Số tiền'] = float(numeric_sodu)
                    df_all.loc[found_idx, 'Tổng Khả Dụng'] = float(numeric_sodu)
                    df_all.loc[found_idx, 'Ngày lập'] = datetime.now(TZ_VN).strftime('%Y-%m-%d')
                    df_all.loc[found_idx, 'NguoiNhap'] = str(current_user)
                    msg = f"Đã ghi đè cập nhật số dư khả dụng cơ sở `{sd_coso}` ngày `{sd_ngay_hl.strftime('%d/%m/%Y')}` thành công!"
                else:
                    # Chưa có -> Thêm dòng mới
                    new_sodu_row = {
                        'Cơ sở': sd_coso,
                        'Ngày lập': datetime.now(TZ_VN).strftime('%Y-%m-%d'),
                        'Ngày hiệu lực': datetime.combine(sd_ngay_hl, datetime.min.time()),
                        'Giờ Nộp Rút': '00:00:00',
                        'Tên khách hàng': 'ĐIỀU CHỈNH SỐ DƯ KHẢ DỤNG',
                        'Loại GD': 'Số Dư Khả Dụng',
                        'Số tiền': float(numeric_sodu),
                        'Loại tiền': sd_tiente,
                        'Trạng thái': 'HT',
                        'NguoiNhap': str(current_user),
                        'NguoiXL': str(current_user),
                        'Loại bản ghi': 'SoDuKhaDung',
                        'Tổng Khả Dụng': float(numeric_sodu)
                    }
                    df_all = pd.concat([df_all, pd.DataFrame([new_sodu_row])], ignore_index=True)
                    msg = f"Đã thêm mới số dư khả dụng cơ sở `{sd_coso}` ngày `{sd_ngay_hl.strftime('%d/%m/%Y')}` thành công!"

                df_all.to_excel(path_data_ds, index=False)
                st.success(f"🎉 {msg}")
            except Exception as e:
                st.error(f"❌ Lỗi lưu số dư khả dụng: {e}")

    # ---------------- 3. XÓA BẢN GHI ----------------
    elif action == "Xóa bản ghi (Chờ XL)":
        st.markdown("#### 🗑️ Quản lý Xóa bản ghi (Chờ XL)")
        if not os.path.exists(path_data_ds): return
        
        df_all = pd.read_excel(path_data_ds, dtype={'NguoiNhap': str, 'NguoiXL': str})
        if 'NguoiXL' not in df_all.columns: df_all['NguoiXL'] = ''
        if 'Loại bản ghi' not in df_all.columns: df_all['Loại bản ghi'] = 'GiaoDich'
        
        df_cho = df_all[(df_all['Trạng thái'] == 'Chờ XL') & (df_all['Loại bản ghi'] == 'GiaoDich')].copy()
        
        if df_cho.empty:
            st.info("ℹ️ Không có giao dịch nào đang ở trạng thái 'Chờ XL'.")
        else:
            df_cho.insert(0, 'Chọn', False)
            edited_df = st.data_editor(
                df_cho,
                column_config={'Chọn': st.column_config.CheckboxColumn("Tích chọn", required=True)},
                disabled=[c for c in df_cho.columns if c != 'Chọn'],
                use_container_width=True,
                key="editor_xoa"
            )
            
            if st.button("⚠️ Xác nhận Xóa các dòng đã tích chọn", type="primary"):
                selected_rows = edited_df[edited_df['Chọn'] == True]
                if selected_rows.empty:
                    st.warning("⚠️ Bạn chưa tích chọn dòng nào để xóa!")
                else:
                    success_count = 0
                    error_messages = []
                    now_vn = datetime.now(TZ_VN)
                    
                    for real_idx, row in selected_rows.iterrows():
                        try:
                            raw_val = row['Ngày hiệu lực']
                            dt_hieu_luc = raw_val if isinstance(raw_val, datetime) else datetime.strptime(str(raw_val)[:10], '%Y-%m-%d')
                            ngay_hieu_luc = dt_hieu_luc.date()
                            ngay_hien_hanh = now_vn.date()
                            
                            du_dieu_kien_xoa = False
                            if ngay_hieu_luc > ngay_hien_hanh:
                                du_dieu_kien_xoa = True
                            elif ngay_hieu_luc == ngay_hien_hanh:
                                base_dt = datetime.combine(ngay_hieu_luc, datetime.min.time())
                                if now_vn.replace(tzinfo=None) <= base_dt + timedelta(hours=3):
                                    du_dieu_kien_xoa = True
                            
                            if du_dieu_kien_xoa:
                                df_all.loc[real_idx, 'Trạng thái'] = 'Xóa'
                                df_all.loc[real_idx, 'NguoiXL'] = str(current_user)
                                success_count += 1
                            else:
                                error_messages.append(f"Dòng {real_idx} đã quá hạn 3 tiếng so với giờ hiệu lực.")
                        except Exception as e:
                            error_messages.append(f"Dòng {real_idx} lỗi: {e}")
                    
                    if success_count > 0:
                        df_all.to_excel(path_data_ds, index=False)
                        st.success(f"✅ Đã xóa thành công {success_count} bản ghi bởi `{current_user}`!")
                    for err in error_messages:
                        st.error(f"❌ {err}")
                    if success_count > 0:
                        st.rerun()

    # ---------------- 4. HOÀN TẤT GD (ĐỐI NỘP + / TRỪ LĨNH - THEO NGUYÊN TẮC) ----------------
    elif action == "Hoàn tất GD (Chờ XL)":
        st.markdown("#### ✅ Hoàn tất Giao dịch & Cập Nhật Số Dư Khả Dụng (+ Nộp / - Lĩnh)")
        if not os.path.exists(path_data_ds): return
        
        df_all = pd.read_excel(path_data_ds, dtype={'NguoiNhap': str, 'NguoiXL': str})
        if 'NguoiXL' not in df_all.columns: df_all['NguoiXL'] = ''
        if 'Loại bản ghi' not in df_all.columns: df_all['Loại bản ghi'] = 'GiaoDich'
        
        df_cho = df_all[(df_all['Trạng thái'] == 'Chờ XL') & (df_all['Loại bản ghi'] == 'GiaoDich')].copy()
        
        if df_cho.empty:
            st.info("ℹ️ Không có giao dịch 'Chờ XL' nào.")
        else:
            df_cho.insert(0, 'Chọn', False)
            edited_df_ht = st.data_editor(
                df_cho,
                column_config={'Chọn': st.column_config.CheckboxColumn("Tích chọn", required=True)},
                disabled=[c for c in df_cho.columns if c != 'Chọn'],
                use_container_width=True,
                key="editor_hoantat"
            )
            
            if st.button("🚀 Xác nhận Hoàn tất (+ Nộp / - Lĩnh số dư khả dụng)", type="primary"):
                selected_rows_ht = edited_df_ht[edited_df_ht['Chọn'] == True]
                if selected_rows_ht.empty:
                    st.warning("⚠️ Bạn chưa tích chọn dòng nào để hoàn tất!")
                else:
                    success_count = 0
                    error_messages_ht = []
                    now_vn = datetime.now(TZ_VN)
                    
                    for real_idx, row in selected_rows_ht.iterrows():
                        try:
                            coso_gd = str(row['Cơ sở']).strip()
                            loai_gd = str(row['Loại GD']).strip()
                            sotien_gd = float(row['Số tiền'] or 0)
                            loaitien_gd = str(row['Loại tiền']).strip()
                            
                            raw_date = row['Ngày hiệu lực']
                            date_str = pd.to_datetime(raw_date, errors='coerce').strftime('%Y-%m-%d')
                            if not date_str or date_str == 'NaT':
                                date_str = now_vn.strftime('%Y-%m-%d')

                            # 1. Quét tìm dòng số dư khả dụng ứng với Cơ sở + Ngày hiệu lực + Loại tiền
                            found_sd_idx = -1
                            for idx_row, r_data in df_all.iterrows():
                                if str(r_data.get('Loại bản ghi', '')) == 'SoDuKhaDung':
                                    r_coso = str(r_data.get('Cơ sở', '')).strip()
                                    r_tiente = str(r_data.get('Loại tiền', '')).strip()
                                    r_ngay = pd.to_datetime(r_data.get('Ngày hiệu lực', ''), errors='coerce').strftime('%Y-%m-%d')
                                    if r_coso == coso_gd and r_tiente == loaitien_gd and r_ngay == date_str:
                                        found_sd_idx = idx_row
                                        break

                            if found_sd_idx != -1:
                                # Đã có -> Áp dụng công thức (+ Nộp / - Lĩnh) và GHI ĐÈ
                                cur_kd = float(df_all.loc[found_sd_idx, 'Tổng Khả Dụng'] or 0)
                                if loai_gd == 'Nộp':
                                    new_kd = cur_kd + sotien_gd
                                else:
                                    new_kd = max(0.0, cur_kd - sotien_gd)
                                
                                df_all.loc[found_sd_idx, 'Tổng Khả Dụng'] = new_kd
                                df_all.loc[found_sd_idx, 'Số tiền'] = new_kd
                            else:
                                # Chưa có -> Khởi tạo dòng số dư mới với giá trị tương ứng
                                init_kd = sotien_gd if loai_gd == 'Nộp' else 0.0
                                new_sodu_row = {
                                    'Cơ sở': coso_gd,
                                    'Ngày lập': datetime.now(TZ_VN).strftime('%Y-%m-%d'),
                                    'Ngày hiệu lực': pd.to_datetime(date_str),
                                    'Giờ Nộp Rút': '00:00:00',
                                    'Tên khách hàng': 'ĐIỀU CHỈNH TỰ ĐỘNG KHI HOÀN TẤT',
                                    'Loại GD': 'Số Dư Khả Dụng',
                                    'Số tiền': float(init_kd),
                                    'Loại tiền': loaitien_gd,
                                    'Trạng thái': 'HT',
                                    'NguoiNhap': str(current_user),
                                    'NguoiXL': str(current_user),
                                    'Loại bản ghi': 'SoDuKhaDung',
                                    'Tổng Khả Dụng': float(init_kd)
                                }
                                df_all = pd.concat([df_all, pd.DataFrame([new_sodu_row])], ignore_index=True)

                            # 2. Cập nhật trạng thái giao dịch thành Hoàn tất (HT)
                            df_all.loc[real_idx, 'Trạng thái'] = 'HT'
                            df_all.loc[real_idx, 'NguoiXL'] = str(current_user)
                            success_count += 1
                        except Exception as e:
                            error_messages_ht.append(f"Dòng {real_idx} lỗi: {e}")
                        
                    if success_count > 0:
                        df_all.to_excel(path_data_ds, index=False)
                        st.success(f"🎉 Hoàn tất thành công {success_count} giao dịch theo nguyên tắc (+ Nộp / - Lĩnh) và đã cập nhật số dư!")
                    for err in error_messages_ht:
                        st.error(f"❌ {err}")
                    if success_count > 0:
                        st.rerun()

    # ---------------- 5. BÁO CÁO ----------------
    elif action == "Báo cáo Đăng ký Nộp Lĩnh":
        st.markdown("#### 📊 Báo Cáo Tổng Hợp Đăng Ký Nộp Lĩnh & Số Dư Khả Dụng Thực Tế (TL)")
        if not os.path.exists(path_data_ds):
            st.warning("⚠️ Chưa có file dữ liệu.")
            return
            
        df_rep = pd.read_excel(path_data_ds)
        if df_rep.empty:
            st.warning("⚠️ Dữ liệu trống.")
            return

        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1: rep_start = st.date_input("Từ ngày hiệu lực", value=datetime.now(TZ_VN).date(), key="rep_start")
        with col_f2: rep_end = st.date_input("Đến ngày hiệu lực", value=datetime.now(TZ_VN).date(), key="rep_end")
        with col_f3: rep_coso = st.selectbox("Chọn Cơ sở", ["All"] + list_coso, key="rep_coso")

        col_f4, col_f5 = st.columns(2)
        available_tt = df_rep['Trạng thái'].dropna().astype(str).unique().tolist() if 'Trạng thái' in df_rep.columns else ["Chờ XL", "HT"]
        with col_f4: rep_tt = st.multiselect("Chọn Trạng thái", available_tt, default=available_tt, key="rep_sel_tt")
        with col_f5: rep_tiente = st.selectbox("Chọn Loại tiền tệ báo cáo", list_tiente, key="rep_sel_tiente")

        if st.button("⚡ Chạy Báo Cáo", type="primary", key="btn_run_bc"):
            if 'Loại bản ghi' not in df_rep.columns:
                df_rep['Loại bản ghi'] = 'GiaoDich'

            df_rep['Clean_Date'] = pd.to_datetime(df_rep['Ngày hiệu lực'], errors='coerce').dt.date
            df_rep = df_rep.dropna(subset=['Clean_Date'])

            start_d = pd.to_datetime(rep_start).date()
            end_d = pd.to_datetime(rep_end).date()

            mask = (df_rep['Clean_Date'] >= start_d) & (df_rep['Clean_Date'] <= end_d)
            if rep_coso != "All": 
                mask = mask & (df_rep['Cơ sở'].astype(str).str.strip() == str(rep_coso).strip())
            if rep_tt:
                mask = mask & (df_rep['Trạng thái'].astype(str).isin(rep_tt))
            mask = mask & (df_rep['Loại tiền'].astype(str).str.strip() == str(rep_tiente).strip())

            df_filtered = df_rep[mask].copy()

            if df_filtered.empty:
                st.info("ℹ️ Không tìm thấy dữ liệu phù hợp với các tiêu chí lọc.")
            else:
                # Quét tìm chính xác số dư khả dụng gốc trong file ứng với điều kiện
                mask_sd = (df_filtered['Loại bản ghi'] == 'SoDuKhaDung')
                base_sodu = df_filtered.loc[mask_sd, 'Số tiền'].sum() if mask_sd.any() else 0.0

                # Tổng Thu và Chi ở trạng thái "Chờ XL"
                mask_cho_xl = (df_filtered['Trạng thái'] == 'Chờ XL') & (df_filtered['Loại bản ghi'] == 'GiaoDich')
                thu_cho_xl = df_filtered.loc[mask_cho_xl & (df_filtered['Loại GD'] == 'Nộp'), 'Số tiền'].sum()
                chi_cho_xl = df_filtered.loc[mask_cho_xl & (df_filtered['Loại GD'] == 'Lĩnh'), 'Số tiền'].sum()

                # Công thức: Số dư Khả dụng TL = Số dư Khả dụng Gốc + Thu (Chờ XL) - Chi (Chờ XL)
                so_du_kha_dung_tl = base_sodu + thu_cho_xl - chi_cho_xl

                st.markdown("---")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("💰 Số Dư Khả Dụng Gốc", f"{base_sodu:,.0f}")
                c2.metric("📈 Thu (Chờ XL)", f"{thu_cho_xl:,.0f}")
                c3.metric("📉 Chi (Chờ XL)", f"{chi_cho_xl:,.0f}")
                c4.metric(f"🌟 Số Dư Khả Dụng TL ({rep_tiente})", f"{so_du_kha_dung_tl:,.0f}")
                st.markdown("---")

                # Hiển thị bảng chi tiết giao dịch
                df_gd_only = df_filtered[df_filtered['Loại bản ghi'] == 'GiaoDich'].copy()
                if not df_gd_only.empty:
                    df_gd_only['Số tiền Nộp'] = df_gd_only.apply(lambda r: r['Số tiền'] if str(r['Loại GD']).strip() == 'Nộp' else 0, axis=1)
                    df_gd_only['Số tiền Lĩnh'] = df_gd_only.apply(lambda r: r['Số tiền'] if str(r['Loại GD']).strip() == 'Lĩnh' else 0, axis=1)

                    if 'NguoiNhap' not in df_gd_only.columns: df_gd_only['NguoiNhap'] = ''
                    if 'NguoiXL' not in df_gd_only.columns: df_gd_only['NguoiXL'] = ''

                    table_out = pd.DataFrame({
                        'Cơ sở': df_gd_only['Cơ sở'],
                        'Người nhập': df_gd_only['NguoiNhap'].fillna('').astype(str),
                        'Người xử lý': df_gd_only['NguoiXL'].fillna('').astype(str),
                        'Ngày Đăng ký': df_gd_only['Ngày lập'],
                        'Ngày hiệu lực': df_gd_only['Ngày hiệu lực'],
                        'Giờ Nộp Rút': df_gd_only['Giờ Nộp Rút'],
                        'CN/PGD': df_gd_only['Tên khách hàng'],
                        'Trạng thái': df_gd_only['Trạng thái'],
                        f'Số tiền Nộp ({rep_tiente})': df_gd_only['Số tiền Nộp'],
                        f'Số tiền Lĩnh ({rep_tiente})': df_gd_only['Số tiền Lĩnh']
                    })
                    st.dataframe(table_out.style.format({f'Số tiền Nộp ({rep_tiente})': '{:,.0f}', f'Số tiền Lĩnh ({rep_tiente})': '{:,.0f}'}), use_container_width=True)
                else:
                    st.info("ℹ️ Không có giao dịch phát sinh nào trong khoảng thời gian và bộ lọc đã chọn.")
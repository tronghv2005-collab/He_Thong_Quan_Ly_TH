import streamlit as st
import pandas as pd
import os
from datetime import datetime, date, timedelta

def show_dang_ky_nop_linh_page(base_dir, shared_data_dir, current_user="System"):
    
    st.subheader("📝 QUẢN LÝ ĐĂNG KÝ NỘP LĨNH TIỀN MẶT")
    st.markdown(f"👤 **Cán bộ thao tác hiện tại:** `{current_user}`")
    
    param_dir = os.path.join(base_dir, "parameters")
    path_thamso_coso = os.path.join(param_dir, "Thamso_coso.xlsx")
    path_thamso_km = os.path.join(param_dir, "Thamso_KM.xlsx")
    path_thamso_tiente = os.path.join(param_dir, "Thamso_tiente.xlsx")
    path_data_ds = os.path.join(shared_data_dir, "Ds_dangkynoplinh.xlsx")
    
    os.makedirs(shared_data_dir, exist_ok=True)
    os.makedirs(param_dir, exist_ok=True)

    # 1. Đọc danh sách Cơ sở
    list_coso = []
    if os.path.exists(path_thamso_coso):
        try:
            df_cs = pd.read_excel(path_thamso_coso)
            col_cs = next((c for c in df_cs.columns if 'co_so' in str(c).lower() or 'cơ sở' in str(c).lower() or 'coso' in str(c).lower()), df_cs.columns[0])
            list_coso = df_cs[col_cs].dropna().astype(str).tolist()
        except Exception:
            pass
    if not list_coso: list_coso = ["Cơ sở 1", "Cơ sở 2", "Hub NVC"]

    # 2. Đọc danh sách Khách hàng (250+ PGD)
    list_kh = []
    if os.path.exists(path_thamso_km):
        try:
            df_km = pd.read_excel(path_thamso_km)
            col_kh = next((c for c in df_km.columns if 'tenkhachhang' in str(c).lower() or 'khách hàng' in str(c).lower()), df_km.columns[0])
            list_kh = df_km[col_kh].dropna().astype(str).tolist()
        except Exception:
            pass
    if not list_kh: list_kh = ["CN Ba Đình", "PGD Hoàn Kiếm", "CN Đống Đa"]

    # 3. Đọc danh sách Tiền tệ
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
            df_init = pd.DataFrame(columns=[
                'Cơ sở', 'Ngày lập', 'Ngày hiệu lực', 'Giờ Nộp Rút', 
                'Tên khách hàng', 'Loại GD', 'Số tiền', 'Loại tiền', 'Trạng thái',
                'NguoiNhap', 'NguoiXL'
            ])
            df_init.to_excel(path_data_ds, index=False)
        except Exception as e:
            st.error(f"❌ Không thể tạo file dữ liệu: {e}")

    action = st.radio("Chọn chức năng:", ["Thêm mới giao dịch", "Xóa bản ghi (Chờ XL)", "Hoàn tất GD (Chờ XL)", "Báo cáo Đăng ký Nộp Lĩnh"], horizontal=True)
    st.markdown("---")

    # ---------------- 1. THÊM MỚI ----------------
    if action == "Thêm mới giao dịch":
        st.markdown("#### ➕ Nhập liệu đăng ký Nộp / Lĩnh mới")
        
        if 'form_data' not in st.session_state:
            st.session_state.form_data = None
        if 'confirm_ghi' not in st.session_state:
            st.session_state.confirm_ghi = False

        col1, col2 = st.columns(2)
        with col1:
            t_coso = st.selectbox("1. Tên Cơ sở", list_coso, key="inp_coso")
            t_ngay_dk = date.today()
            st.text(f"2. Ngày Đăng ký: {t_ngay_dk.strftime('%d/%m/%Y')}")
            
            t_ngay_hl = st.date_input("3. Ngày hiệu lực", value=date.today(), key="inp_ngay_hl")
            
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
                    'NguoiXL': str(current_user)  # Lúc chờ xử lý, Người xử lý ban đầu trùng với Người nhập
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
                        if os.path.exists(path_data_ds):
                            df_current = pd.read_excel(path_data_ds, dtype={'NguoiNhap': str, 'NguoiXL': str})
                            if 'NguoiNhap' not in df_current.columns: df_current['NguoiNhap'] = ''
                            if 'NguoiXL' not in df_current.columns: df_current['NguoiXL'] = ''
                        else:
                            df_current = pd.DataFrame(columns=new_row.keys())
                            
                        df_current = pd.concat([df_current, pd.DataFrame([new_row])], ignore_index=True)
                        
                        df_current['NguoiNhap'] = df_current['NguoiNhap'].fillna('').astype(str)
                        df_current['NguoiXL'] = df_current['NguoiXL'].fillna('').astype(str)
                        
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

    # ---------------- 2. XÓA BẢN GHI ----------------
    elif action == "Xóa bản ghi (Chờ XL)":
        st.markdown("#### 🗑️ Quản lý Xóa bản ghi (Tích chọn các dòng cần xóa bên dưới)")
        if not os.path.exists(path_data_ds): return
        
        df_all = pd.read_excel(path_data_ds, dtype={'NguoiNhap': str, 'NguoiXL': str})
        if 'NguoiXL' not in df_all.columns: df_all['NguoiXL'] = ''
        if 'NguoiNhap' not in df_all.columns: df_all['NguoiNhap'] = ''
        
        df_all['NguoiXL'] = df_all['NguoiXL'].fillna('').astype(str)
        df_all['NguoiNhap'] = df_all['NguoiNhap'].fillna('').astype(str)
        
        df_cho = df_all[df_all['Trạng thái'] == 'Chờ XL'].copy()
        
        if df_cho.empty:
            st.info("ℹ️ Không có bản ghi nào đang ở trạng thái 'Chờ XL'.")
        else:
            df_cho.insert(0, 'Chọn', False)
            
            edited_df = st.data_editor(
                df_cho,
                column_config={'Chọn': st.column_config.CheckboxColumn("Tích chọn", required=True)},
                disabled=[c for c in df_cho.columns if c != 'Chọn'],
                hide_index=False,
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
                    
                    for real_idx, row in selected_rows.iterrows():
                        try:
                            raw_val = row['Ngày hiệu lực']
                            if isinstance(raw_val, datetime):
                                dt_hieu_luc = raw_val
                            else:
                                s_val = str(raw_val).strip()
                                if len(s_val) >= 10 and s_val[4] == '-' and s_val[7] == '-':
                                    dt_hieu_luc = datetime.strptime(s_val[:10], '%Y-%m-%d')
                                else:
                                    dt_hieu_luc = datetime.now()
                            
                            ngay_hieu_luc = dt_hieu_luc.date()
                            ngay_hien_hanh = datetime.now().date()
                            
                            du_dieu_kien_xoa = False
                            if ngay_hieu_luc > ngay_hien_hanh:
                                du_dieu_kien_xoa = True
                            elif ngay_hieu_luc == ngay_hien_hanh:
                                base_dt = dt_hieu_luc if isinstance(dt_hieu_luc, datetime) else datetime.combine(ngay_hieu_luc, datetime.min.time())
                                gio_gioi_han = base_dt + timedelta(hours=3)
                                if datetime.now() <= gio_gioi_han:
                                    du_dieu_kien_xoa = True
                            
                            if du_dieu_kien_xoa:
                                df_all.loc[real_idx, 'Trạng thái'] = 'Xóa'
                                df_all.loc[real_idx, 'NguoiXL'] = str(current_user)  # Ghi nhận người thực hiện Xóa
                                success_count += 1
                            else:
                                error_messages.append(f"Dòng {real_idx} quá hạn 3 tiếng.")
                        except Exception as e:
                            error_messages.append(f"Dòng {real_idx} lỗi: {e}")
                    
                    if success_count > 0:
                        df_all['NguoiXL'] = df_all['NguoiXL'].fillna('').astype(str)
                        df_all.to_excel(path_data_ds, index=False)
                        st.success(f"✅ Đã xóa thành công {success_count} bản ghi bởi `{current_user}`!")
                    
                    for err in error_messages:
                        st.error(f"❌ {err}")
                        
                    if success_count > 0:
                        st.rerun()

    # ---------------- 3. HOÀN TẤT GD ----------------
    elif action == "Hoàn tất GD (Chờ XL)":
        st.markdown("#### ✅ Hoàn tất Giao dịch (Tích chọn các dòng cần hoàn tất bên dưới)")
        if not os.path.exists(path_data_ds): return
        
        df_all = pd.read_excel(path_data_ds, dtype={'NguoiNhap': str, 'NguoiXL': str})
        if 'NguoiXL' not in df_all.columns: df_all['NguoiXL'] = ''
        if 'NguoiNhap' not in df_all.columns: df_all['NguoiNhap'] = ''
        
        df_all['NguoiXL'] = df_all['NguoiXL'].fillna('').astype(str)
        df_all['NguoiNhap'] = df_all['NguoiNhap'].fillna('').astype(str)
        
        df_cho = df_all[df_all['Trạng thái'] == 'Chờ XL'].copy()
        
        if df_cho.empty:
            st.info("ℹ️ Không có bản ghi 'Chờ XL'.")
        else:
            df_cho.insert(0, 'Chọn', False)
            
            edited_df_ht = st.data_editor(
                df_cho,
                column_config={'Chọn': st.column_config.CheckboxColumn("Tích chọn", required=True)},
                disabled=[c for c in df_cho.columns if c != 'Chọn'],
                hide_index=False,
                use_container_width=True,
                key="editor_hoantat"
            )
            
            if st.button("🚀 Xác nhận Hoàn tất các dòng đã tích chọn", type="primary"):
                selected_rows_ht = edited_df_ht[edited_df_ht['Chọn'] == True]
                
                if selected_rows_ht.empty:
                    st.warning("⚠️ Bạn chưa tích chọn dòng nào để hoàn tất!")
                else:
                    success_count_ht = 0
                    error_messages_ht = []
                    
                    for real_idx, row in selected_rows_ht.iterrows():
                        try:
                            raw_val = row['Ngày hiệu lực']
                            if isinstance(raw_val, datetime):
                                dt_hieu_luc = raw_val
                            else:
                                s_val = str(raw_val).strip()
                                if len(s_val) >= 10 and s_val[4] == '-' and s_val[7] == '-':
                                    dt_hieu_luc = datetime.strptime(s_val[:10], '%Y-%m-%d')
                                else:
                                    dt_hieu_luc = datetime.now()
                                    
                            ngay_hieu_luc = dt_hieu_luc.date()
                            ngay_hien_hanh = datetime.now().date()
                            
                            if ngay_hieu_luc != ngay_hien_hanh:
                                error_messages_ht.append(f"Dòng {real_idx}: Ngày hiệu lực không phải ngày hôm nay.")
                            else:
                                base_dt = dt_hieu_luc if isinstance(dt_hieu_luc, datetime) else datetime.combine(ngay_hieu_luc, datetime.min.time())
                                gio_gioi_han = base_dt + timedelta(hours=3)
                                if datetime.now() <= gio_gioi_han:
                                    df_all.loc[real_idx, 'Trạng thái'] = 'HT'
                                    df_all.loc[real_idx, 'NguoiXL'] = str(current_user)  # Ghi nhận người thực hiện Hoàn tất
                                    success_count_ht += 1
                                else:
                                    error_messages_ht.append(f"Dòng {real_idx}: Đã quá thời hạn 3 tiếng.")
                        except Exception as e:
                            error_messages_ht.append(f"Dòng {real_idx} lỗi: {e}")
                    
                    if success_count_ht > 0:
                        df_all['NguoiXL'] = df_all['NguoiXL'].fillna('').astype(str)
                        df_all.to_excel(path_data_ds, index=False)
                        st.success(f"✅ Đã hoàn tất thành công {success_count_ht} giao dịch bởi `{current_user}`!")
                    
                    for err in error_messages_ht:
                        st.error(f"❌ {err}")
                        
                    if success_count_ht > 0:
                        st.rerun()

    # ---------------- 4. BÁO CÁO ----------------
    elif action == "Báo cáo Đăng ký Nộp Lĩnh":
        st.markdown("#### 📊 Báo Cáo Tổng Hợp Đăng Ký Nộp Lĩnh")
        if not os.path.exists(path_data_ds):
            st.warning("⚠️ Chưa có file dữ liệu đăng ký.")
            return
            
        try:
            df_rep = pd.read_excel(path_data_ds, dtype={'NguoiNhap': str, 'NguoiXL': str})
            if 'NguoiNhap' not in df_rep.columns: df_rep['NguoiNhap'] = ''
            if 'NguoiXL' not in df_rep.columns: df_rep['NguoiXL'] = ''
            df_rep['NguoiNhap'] = df_rep['NguoiNhap'].fillna('').astype(str)
            df_rep['NguoiXL'] = df_rep['NguoiXL'].fillna('').astype(str)
        except Exception as e:
            st.error(f"❌ Không thể đọc file dữ liệu: {e}")
            return

        if df_rep.empty:
            st.warning("⚠️ File dữ liệu đang trống.")
            return

        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1: rep_start = st.date_input("Từ ngày hiệu lực", value=date.today(), key="rep_start_date")
        with col_f2: rep_end = st.date_input("Đến ngày hiệu lực", value=date.today(), key="rep_end_date")
        with col_f3: rep_coso = st.selectbox("Chọn Cơ sở", ["All"] + list_coso, key="rep_sel_coso")
        
        available_tt = df_rep['Trạng thái'].dropna().astype(str).unique().tolist() if 'Trạng thái' in df_rep.columns else ["Chờ XL", "HT"]
        with col_f4: rep_tt = st.multiselect("Chọn Trạng thái", available_tt, default=available_tt, key="rep_sel_tt")

        rep_tiente = st.selectbox("Chọn Loại tiền tệ báo cáo", list_tiente, key="rep_sel_tiente")

        if st.button("⚡ Chạy Báo Cáo", type="primary", key="btn_chay_bc"):
            def safe_parse_date(val):
                if pd.isna(val):
                    return None
                if isinstance(val, datetime):
                    return val.date()
                try:
                    s = str(val).strip()
                    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
                        return datetime.strptime(s[:10], '%Y-%m-%d').date()
                except:
                    pass
                return None

            df_rep['Clean_Date'] = df_rep['Ngày hiệu lực'].apply(safe_parse_date)
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
                st.info(f"ℹ️ Không tìm thấy bản ghi nào trong khoảng từ {start_d.strftime('%d/%m/%Y')} đến {end_d.strftime('%d/%m/%Y')}.")
            else:
                df_filtered['Số tiền Nộp'] = df_filtered.apply(lambda r: r['Số tiền'] if str(r['Loại GD']).strip() == 'Nộp' else 0, axis=1)
                df_filtered['Số tiền Lĩnh'] = df_filtered.apply(lambda r: r['Số tiền'] if str(r['Loại GD']).strip() == 'Lĩnh' else 0, axis=1)

                # Kiểm tra an toàn cột NguoiNhap & NguoiXL luôn hiển thị đầy đủ
                if 'NguoiNhap' not in df_filtered.columns: df_filtered['NguoiNhap'] = ''
                if 'NguoiXL' not in df_filtered.columns: df_filtered['NguoiXL'] = ''

                table_out = pd.DataFrame({
                    'Cơ sở': df_filtered['Cơ sở'],
                    'Người nhập (NguoiNhap)': df_filtered['NguoiNhap'].fillna('').astype(str),
                    'Người xử lý (NguoiXL)': df_filtered['NguoiXL'].fillna('').astype(str),
                    'Ngày Đăng ký': df_filtered['Ngày lập'],
                    'Ngày hiệu lực': df_filtered['Ngày hiệu lực'],
                    'Giờ Nộp Rút': df_filtered['Giờ Nộp Rút'],
                    'CN/PGD': df_filtered['Tên khách hàng'],
                    'Trạng thái': df_filtered['Trạng thái'],
                    f'Số tiền Nộp ({rep_tiente})': df_filtered['Số tiền Nộp'],
                    f'Số tiền Lĩnh ({rep_tiente})': df_filtered['Số tiền Lĩnh']
                })

                format_dict = {
                    f'Số tiền Nộp ({rep_tiente})': '{:,.0f}',
                    f'Số tiền Lĩnh ({rep_tiente})': '{:,.0f}'
                }
                st.dataframe(table_out.style.format(format_dict), use_container_width=True)

                sum_nop = table_out[f'Số tiền Nộp ({rep_tiente})'].sum()
                sum_linh = table_out[f'Số tiền Lĩnh ({rep_tiente})'].sum()
                chenh_lech = sum_nop - sum_linh

                st.markdown("---")
                m1, m2, m3 = st.columns(3)
                m1.metric(f"📈 Tổng Nộp ({rep_tiente})", f"{sum_nop:,.0f}")
                m2.metric(f"📉 Tổng Lĩnh ({rep_tiente})", f"{sum_linh:,.0f}")
                m3.metric(f"⚖️ Chênh lệch (Nộp - Lĩnh)", f"{chenh_lech:,.0f}")
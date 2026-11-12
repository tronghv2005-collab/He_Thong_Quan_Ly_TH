import streamlit as st
import pandas as pd
import os
from datetime import datetime, date, timedelta, timezone

def show_quan_ly_loi_page(base_dir, shared_data_dir, current_user="System"):
    st.subheader("⚠️ QUẢN LÝ LỖI / VI PHẠM NGHIỆP VỤ KHO QUỸ")
    st.markdown(f"👤 **Cán bộ thao tác hiện tại:** `{current_user}`")
    
    param_dir = os.path.join(base_dir, "parameters")
    path_thamso_canbo = os.path.join(param_dir, "Thamso_canbo.xlsx")
    path_thamso_loi = os.path.join(param_dir, "Thamso_loi.xlsx")
    path_data_loi = os.path.join(shared_data_dir, "Ds_loi.xlsx")
    
    os.makedirs(shared_data_dir, exist_ok=True)
    os.makedirs(param_dir, exist_ok=True)

    TZ_VN = timezone(timedelta(hours=7))
    today_vn = datetime.now(TZ_VN).date()

    # --- TRÍCH XUẤT CHÍNH XÁC MÃ NV HIỆN TẠI ---
    current_manv = str(current_user).split('-')[0].strip()

    # --- 1. ĐỌC THAM SỐ CÁN BỘ, CƠ SỞ VÀ NHÓM QUYỀN ---
    dict_manv_coso = {} 
    dict_manv_quyen = {} 
    list_canbo_display = [str(current_user)]

    if os.path.exists(path_thamso_canbo):
        try:
            df_cb = pd.read_excel(path_thamso_canbo, dtype=str)
            col_manv = next((c for c in df_cb.columns if 'manv' in str(c).lower() or 'ma_nv' in str(c).lower()), df_cb.columns[0])
            col_ten = next((c for c in df_cb.columns if 'ten' in str(c).lower() or 'tên' in str(c).lower()), None)
            col_coso = next((c for c in df_cb.columns if 'co_so' in str(c).lower() or 'cơ sở' in str(c).lower() or 'coso' in str(c).lower()), None)
            col_quyen = next((c for c in df_cb.columns if 'nhomquyen' in str(c).lower() or 'nhóm quyền' in str(c).lower() or 'quyen' in str(c).lower()), None)
            
            list_canbo_display = []
            for _, row in df_cb.iterrows():
                m_nv = str(row[col_manv]).strip()
                t_nv = str(row[col_ten]).strip() if (col_ten and pd.notnull(row[col_ten])) else m_nv
                
                # Lấy trực tiếp giá trị từ cột TenNV đã được ghép sẵn
                list_canbo_display.append(t_nv)
                
                if col_coso:
                    dict_manv_coso[m_nv] = str(row[col_coso]).strip()
                if col_quyen:
                    dict_manv_quyen[m_nv] = str(row[col_quyen]).strip().upper()
        except Exception as e:
            st.warning(f"⚠️ Không đọc được tham số từ file cán bộ: {e}")

    if not list_canbo_display: list_canbo_display = [str(current_user)]

    current_user_role = dict_manv_quyen.get(current_manv, "NV") 
    st.markdown(f"🔑 **Nhóm quyền của bạn (Mã: `{current_manv}`):** `{current_user_role}`")

    # --- 2. ĐỌC THAM SỐ LỖI & HIỂN THỊ KÈM HÀNH VI VI PHẠM ---
    if not os.path.exists(path_thamso_loi):
        try:
            df_init_loi = pd.DataFrame({
                'MaLoi': ['L01', 'L02'],
                'MucLoi': ['Nghiêm trọng', 'Trung bình'],
                'NhomLoi': ['Kho quỹ', 'Giao dịch'],
                'NhomNV': ['Thu gom', 'Tiếp quỹ'],
                'HanhViViPham': ['Sai lệch tồn quỹ', 'Chậm giờ giao dịch']
            })
            df_init_loi.to_excel(path_thamso_loi, index=False)
        except:
            pass

    dict_thamso_loi = {}
    list_maloi_display = ['L01 - Sai lệch tồn quỹ']
    
    if os.path.exists(path_thamso_loi):
        try:
            df_l = pd.read_excel(path_thamso_loi, dtype=str)
            if 'MaLoi' in df_l.columns:
                list_maloi_display = []
                for _, r in df_l.iterrows():
                    m_loi = str(r['MaLoi']).strip()
                    h_vi = str(r.get('HanhViViPham', '')).strip()
                    
                    label_loi = f"{m_loi} - {h_vi}" if h_vi else m_loi
                    list_maloi_display.append(label_loi)
                    
                    dict_thamso_loi[m_loi] = {
                        'MucLoi': str(r.get('MucLoi', '')).strip(),
                        'NhomLoi': str(r.get('NhomLoi', '')).strip(),
                        'NhomNV': str(r.get('NhomNV', '')).strip(),
                        'HanhViViPham': h_vi
                    }
        except Exception as e:
            st.warning(f"⚠️ Lỗi đọc file Thamso_loi.xlsx: {e}")

    if not os.path.exists(path_data_loi):
        try:
            df_init_ds = pd.DataFrame(columns=[
                'Cơ sở', 'Ngày nhập', 'Ngày phát sinh lỗi', 'Trạng thái', 'Tên cán bộ',
                'Mã lỗi', 'Mục lỗi', 'Nhóm lỗi', 'Nhóm nghiệp vụ', 'Hành vi vi phạm',
                'Mô tả lỗi', 'Đề xuất xử lý', 'NguoiNhap', 'NguoiDuyet'
            ])
            df_init_ds.to_excel(path_data_loi, index=False)
        except Exception as e:
            st.error(f"❌ Không thể tạo file dữ liệu lỗi: {e}")

    # --- 3. PHÂN QUYỀN MENU CHỨC NĂNG ---
    available_actions = []
    if current_user_role in ["NV", "KSV"]:
        available_actions.extend(["➕ Thêm mới lỗi", "🔍 Tìm kiếm & Sửa lỗi (Từ chối)"])
    if current_user_role in ["KSV", "LĐTT", "LDTT"]:
        available_actions.append("🛡️ Phê duyệt lỗi (KSV/LĐTT)")
    available_actions.append("📊 Báo cáo tổng hợp lỗi")

    if not available_actions:
        st.error("❌ Tài khoản của bạn chưa được cấu hình Nhóm quyền hợp lệ trong file Thamso_canbo.xlsx.")
        return

    action = st.radio("Chọn chức năng quản lý lỗi:", available_actions, horizontal=True)
    st.markdown("---")

    # ---------------- 1. THÊM MỚI LỖI ----------------
    if action == "➕ Thêm mới lỗi":
        st.markdown("#### 📝 Màn hình nhập liệu phát sinh lỗi mới")
        with st.form("form_nhap_loi"):
            col1, col2 = st.columns(2)
            with col1:
                st.text(f"1. Ngày nhập (Hiện hành): {today_vn.strftime('%d/%m/%Y')}")
                t_ngay_phat_sinh = st.date_input("2. Ngày phát sinh lỗi (Phải trước ngày hiện hành)", value=today_vn - timedelta(days=0))
                
                t_ten_cb_chon = st.selectbox("3. Cán bộ liên quan", list_canbo_display)
                t_ma_nv_chon = t_ten_cb_chon.split('-')[0].strip() if '-' in t_ten_cb_chon else t_ten_cb_chon.split()[0].strip()
                
                t_coso = dict_manv_coso.get(t_ma_nv_chon, "Chưa xác định cơ sở")
                st.markdown(f"📍 **Cơ sở trực thuộc:** <span style='color: #1F4E78; font-weight: bold;'>{t_coso}</span>", unsafe_allow_html=True)
                
                # Chọn mã lỗi kèm hành vi vi phạm
                t_maloi_chon = st.selectbox("4. Mã lỗi & Hành vi vi phạm", list_maloi_display)
                t_maloi = t_maloi_chon.split('-')[0].strip()
                
                info_loi = dict_thamso_loi.get(t_maloi, {})
                t_mucloi = info_loi.get('MucLoi', '')
                t_nhomloi = info_loi.get('NhomLoi', '')
                
            with col2:
                t_nhomnv = info_loi.get('NhomNV', '')
                t_hanhvi = info_loi.get('HanhViViPham', '')
                
                st.text_input("5. Mục lỗi (Tự động theo Mã lỗi)", value=t_mucloi, disabled=True)
                st.text_input("6. Nhóm lỗi (Tự động theo Mã lỗi)", value=t_nhomloi, disabled=True)
                st.text_input("7. Nhóm nghiệp vụ (Tự động theo Mã lỗi)", value=t_nhomnv, disabled=True)
                st.text_input("8. Hành vi vi phạm (Tự động theo Mã lỗi)", value=t_hanhvi, disabled=True)

            st.markdown("---")
            t_mota = st.text_area("9. Mô tả lỗi", placeholder="Nhập chi tiết diễn biến sự việc...")
            t_dexuat = st.text_area("10. Đề xuất xử lý", placeholder="Nhập đề xuất hướng giải quyết...")

            submit_btn = st.form_submit_button("💾 Lưu bản ghi lỗi", use_container_width=True)

            if submit_btn:
                if t_ngay_phat_sinh > today_vn:
                    st.error("❌ Lỗi: Ngày phát sinh lỗi bắt buộc phải nhỏ hơn hoặc bằng Ngày hiện hành!")
                elif not t_mota.strip():
                    st.warning("⚠️ Vui lòng nhập mô tả chi tiết lỗi.")
                else:
                    try:
                        new_record = {
                            'Cơ sở': t_coso,
                            'Ngày nhập': today_vn.strftime('%Y-%m-%d'),
                            'Ngày phát sinh lỗi': t_ngay_phat_sinh.strftime('%Y-%m-%d'),
                            'Trạng thái': 'Chờ phê duyệt',
                            'Tên cán bộ': t_ten_cb_chon,
                            'Mã lỗi': t_maloi,
                            'Mục lỗi': t_mucloi,
                            'Nhóm lỗi': t_nhomloi,
                            'Nhóm nghiệp vụ': t_nhomnv,
                            'Hành vi vi phạm': t_hanhvi,
                            'Mô tả lỗi': t_mota,
                            'Đề xuất xử lý': t_dexuat,
                            'NguoiNhap': str(current_manv),
                            'NguoiDuyet': ''
                        }
                        df_old = pd.read_excel(path_data_loi, dtype={'NguoiNhap': str, 'NguoiDuyet': str}) if os.path.exists(path_data_loi) else pd.DataFrame(columns=new_record.keys())
                        df_new = pd.concat([df_old, pd.DataFrame([new_record])], ignore_index=True)
                        df_new['NguoiDuyet'] = df_new['NguoiDuyet'].fillna('').astype(str)
                        df_new.to_excel(path_data_loi, index=False)
                        st.success("🎉 Lưu bản ghi lỗi thành công! Trạng thái hiện tại: **Chờ phê duyệt**.")
                    except Exception as e:
                        st.error(f"❌ Lỗi khi ghi dữ liệu: {e}")

    # ---------------- 2. TÌM KIẾM & SỬA LỖI ----------------
    elif action == "🔍 Tìm kiếm & Sửa lỗi (Từ chối)":
        st.markdown("#### 🔍 Tìm kiếm danh sách lỗi & Chỉnh sửa bản ghi bị 'Từ chối'")
        if not os.path.exists(path_data_loi):
            st.info("Chưa có dữ liệu lỗi.")
            return

        df_all = pd.read_excel(path_data_loi, dtype={'NguoiNhap': str, 'NguoiDuyet': str})
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1: f_ngay = st.date_input("Lọc theo Ngày nhập", value=None)
        with c_f2: f_tt = st.selectbox("Lọc theo Trạng thái", ["Tất cả", "Chờ phê duyệt", "Từ chối", "Phê duyệt", "Hủy bỏ lỗi"])
        with c_f3: f_tukhoa = st.text_input("Tìm kiếm theo từ khóa", "")

        mask = pd.Series([True] * len(df_all))
        if f_ngay: mask &= (df_all['Ngày nhập'].astype(str) == f_ngay.strftime('%Y-%m-%d'))
        if f_tt != "Tất cả": mask &= (df_all['Trạng thái'].astype(str) == f_tt)
        if f_tukhoa.strip(): mask &= df_all.astype(str).apply(lambda row: row.str.contains(f_tukhoa, case=False).any(), axis=1)

        st.dataframe(df_all[mask].copy(), use_container_width=True)
        
        tu_choi_records = df_all[df_all['Trạng thái'] == 'Từ chối']
        if not tu_choi_records.empty:
            st.markdown("##### ✏️ Chỉnh sửa bản ghi ở trạng thái **'Từ chối'**")
            selected_idx = st.selectbox("Chọn dòng lỗi để sửa:", tu_choi_records.index, format_func=lambda i: f"Dòng {i} - CB: {tu_choi_records.loc[i, 'Tên cán bộ']}")
            row_edit = df_all.loc[selected_idx]
            with st.form("form_sua_loi"):
                s_mota = st.text_area("Cập nhật Mô tả lỗi", value=str(row_edit['Mô tả lỗi']))
                s_dexuat = st.text_area("Cập nhật Đề xuất xử lý", value=str(row_edit['Đề xuất xử lý']))
                if st.form_submit_button("🔄 Lưu lại & Gửi duyệt lại", type="primary"):
                    df_all.loc[selected_idx, 'Mô tả lỗi'] = s_mota
                    df_all.loc[selected_idx, 'Đề xuất xử lý'] = s_dexuat
                    df_all.loc[selected_idx, 'Trạng thái'] = 'Chờ phê duyệt'
                    df_all['NguoiDuyet'] = df_all['NguoiDuyet'].fillna('').astype(str)
                    df_all.to_excel(path_data_loi, index=False)
                    st.success("✅ Đã cập nhật thành công!")
                    st.rerun()

    # ---------------- 3. PHÊ DUYỆT LỖI ----------------
    elif action == "🛡️ Phê duyệt lỗi (KSV/LĐTT)":
        st.markdown("#### 🛡️ Màn hình Phê duyệt Lỗi (Kiểm soát toàn diện)")
        if not os.path.exists(path_data_loi): return
        df_all = pd.read_excel(path_data_loi, dtype={'NguoiNhap': str, 'NguoiDuyet': str})
        if 'NguoiDuyet' not in df_all.columns: df_all['NguoiDuyet'] = ''
        
        df_cho_duyet = df_all[df_all['Trạng thái'] == 'Chờ phê duyệt'].copy()
        if df_cho_duyet.empty:
            st.info("🎉 Không có bản ghi nào đang chờ phê duyệt.")
        else:
            for idx, row in df_cho_duyet.iterrows():
                nguoi_nhap_ban_ghi = str(row.get('NguoiNhap', '')).strip()
                is_self_created = (nguoi_nhap_ban_ghi == current_manv)

                with st.expander(f"📌 [Dòng {idx}] Cơ sở: {row.get('Cơ sở', '')} | Cán bộ: {row.get('Tên cán bộ', '')} | Mã lỗi: {row.get('Mã lỗi', '')}" + (" ⚠️ (Do chính bạn nhập)" if is_self_created else "")):
                    col_info1, col_info2 = st.columns(2)
                    with col_info1:
                        st.markdown(f"- **Ngày nhập:** {row.get('Ngày nhập', '')}")
                        st.markdown(f"- **Ngày phát sinh lỗi:** {row.get('Ngày phát sinh lỗi', '')}")
                        st.markdown(f"- **Cán bộ liên quan:** {row.get('Tên cán bộ', '')}")
                        st.markdown(f"- **Cơ sở:** {row.get('Cơ sở', '')}")
                        st.markdown(f"- **Mã lỗi:** {row.get('Mã lỗi', '')}")
                        st.markdown(f"- **Mục lỗi:** {row.get('Mục lỗi', '')}")
                    with col_info2:
                        st.markdown(f"- **Nhóm lỗi:** {row.get('Nhóm lỗi', '')}")
                        st.markdown(f"- **Nhóm nghiệp vụ:** {row.get('Nhóm nghiệp vụ', '')}")
                        st.markdown(f"- **Hành vi vi phạm:** {row.get('Hành vi vi phạm', '')}")
                        st.markdown(f"- **Mã NV người nhập (NguoiNhap):** `{nguoi_nhap_ban_ghi}`")
                        st.markdown(f"- **Trạng thái:** `{row.get('Trạng thái', '')}`")

                    st.markdown("---")
                    st.markdown(f"📝 **Mô tả lỗi:** {row.get('Mô tả lỗi', '')}")
                    st.markdown(f"💡 **Đề xuất xử lý:** {row.get('Đề xuất xử lý', '')}")
                    st.markdown("---")

                    if is_self_created and current_user_role == "KSV":
                        st.warning("⚠️ **Bạn là Kiểm soát viên (KSV) đồng thời là người nhập bản ghi này. Theo quy định kiểm soát nội bộ, bạn không được phép tự phê duyệt bản ghi do chính mình tạo ra!**")
                    else:
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            if st.button("✅ Phê duyệt", key=f"pd_{idx}", use_container_width=True):
                                df_all['NguoiDuyet'] = df_all['NguoiDuyet'].fillna('').astype(str)
                                df_all.loc[idx, 'Trạng thái'] = 'Phê duyệt'
                                df_all.loc[idx, 'NguoiDuyet'] = str(current_manv)
                                df_all.to_excel(path_data_loi, index=False)
                                st.success("✅ Đã phê duyệt thành công!")
                                st.rerun()
                        with c2:
                            if st.button("❌ Từ chối", key=f"tc_{idx}", use_container_width=True):
                                df_all['NguoiDuyet'] = df_all['NguoiDuyet'].fillna('').astype(str)
                                df_all.loc[idx, 'Trạng thái'] = 'Từ chối'
                                df_all.loc[idx, 'NguoiDuyet'] = str(current_manv)
                                df_all.to_excel(path_data_loi, index=False)
                                st.success("❌ Đã chuyển sang trạng thái Từ chối!")
                                st.rerun()
                        with c3:
                            if st.button("🚫 Hủy bỏ", key=f"hb_{idx}", use_container_width=True):
                                df_all['NguoiDuyet'] = df_all['NguoiDuyet'].fillna('').astype(str)
                                df_all.loc[idx, 'Trạng thái'] = 'Hủy bỏ lỗi'
                                df_all.loc[idx, 'NguoiDuyet'] = str(current_manv)
                                df_all.to_excel(path_data_loi, index=False)
                                st.success("🚫 Đã hủy bỏ bản ghi lỗi!")
                                st.rerun()

    # ---------------- 4. BÁO CÁO ----------------
    elif action == "📊 Báo cáo tổng hợp lỗi":
        st.markdown("#### 📊 Báo Cáo Tổng Hợp Vi Phạm Kho Quỹ")
        if not os.path.exists(path_data_loi): return
        df_rep = pd.read_excel(path_data_loi, dtype={'NguoiNhap': str, 'NguoiDuyet': str})
        
        if not df_rep.empty and 'Trạng thái' in df_rep.columns:
            st.markdown("##### 🔍 Bộ lọc báo cáo:")
            selected_tt = st.multiselect("Lọc theo Trạng thái:", options=df_rep['Trạng thái'].dropna().astype(str).unique().tolist(), default=df_rep['Trạng thái'].dropna().astype(str).unique().tolist())
            search_kw = st.text_input("Tìm kiếm nhanh:", "")
            
            df_filtered = df_rep.copy()
            if selected_tt: df_filtered = df_filtered[df_filtered['Trạng thái'].astype(str).isin(selected_tt)]
            if search_kw.strip(): df_filtered = df_filtered[df_filtered.astype(str).apply(lambda row: row.str.contains(search_kw, case=False).any(), axis=1)]
            
            st.dataframe(df_filtered, use_container_width=True)
            if not df_filtered.empty:
                st.bar_chart(df_filtered['Trạng thái'].value_counts())
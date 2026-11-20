import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import io
import textwrap
from datetime import datetime, date, timedelta, timezone

# Danh sách toàn bộ các cột văn bản cần chuẩn hóa 100 ký tự/dòng
TEXT_COLS_100 = ['Mục lỗi', 'Nhóm lỗi', 'Nhóm nghiệp vụ', 'Hành vi vi phạm', 'Mô tả lỗi', 'Đề xuất xử lý']

def tu_dong_xuong_dong(text, width=100):
    """Tự động ngắt dòng tối đa `width` ký tự nhưng giữ nguyên vẹn từ."""
    if text is None or pd.isna(text):
        return ""
    lines = str(text).splitlines()
    wrapped_lines = []
    for line in lines:
        if line.strip():
            wrapped_lines.extend(textwrap.wrap(line, width=width, break_long_words=False))
        else:
            wrapped_lines.append("")
    return "\n".join(wrapped_lines)

def kich_hoat_tu_dong_xuong_dong_o_nhap(max_chars=100):
    """Chèn JavaScript để tự động xuống dòng ngay khi gõ vượt quá 100 ký tự trên 1 dòng trong text_area."""
    js_code = f"""
    <script>
    const doc = window.parent.document;
    function attachAutoWrap() {{
        const textareas = doc.querySelectorAll('textarea');
        textareas.forEach(ta => {{
            if (ta.dataset.autoWrapAttached) return;
            ta.dataset.autoWrapAttached = "true";
            
            ta.addEventListener('input', function() {{
                const start = this.selectionStart;
                const val = this.value;
                const lines = val.split('\\n');
                let changed = false;
                let newLines = [];
                
                for (let i = 0; i < lines.length; i++) {{
                    let line = lines[i];
                    while (line.length > {max_chars}) {{
                        changed = true;
                        let breakIdx = line.lastIndexOf(' ', {max_chars});
                        if (breakIdx <= 0) breakIdx = {max_chars};
                        newLines.push(line.substring(0, breakIdx));
                        line = line.substring(breakIdx).replace(/^\\s+/, '');
                    }}
                    newLines.push(line);
                }}
                
                if (changed) {{
                    const diff = newLines.join('\\n').length - val.length;
                    this.value = newLines.join('\\n');
                    this.selectionStart = this.selectionEnd = Math.max(0, start + diff);
                    this.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
            }});
        }});
    }}
    setInterval(attachAutoWrap, 800);
    </script>
    """
    components.html(js_code, height=0, width=0)

def show_quan_ly_loi_page(base_dir, shared_data_dir, current_user="System"):
    st.subheader("⚠️ QUẢN LÝ LỖI / VI PHẠM NGHIỆP VỤ KHO QUỸ")
    st.markdown(f"👤 **Cán bộ thao tác hiện tại:** `{current_user}`")
    
    # Kích hoạt tự động ngắt dòng 100 ký tự trực tiếp trên các ô nhập liệu
    kich_hoat_tu_dong_xuong_dong_o_nhap(max_chars=100)
    
    param_dir = os.path.join(base_dir, "parameters")
    path_thamso_canbo = os.path.join(param_dir, "Thamso_canbo.xlsx")
    path_thamso_loi = os.path.join(param_dir, "Thamso_loi.xlsx")
    path_data_loi = os.path.join(shared_data_dir, "Ds_loi.xlsx")
    
    os.makedirs(shared_data_dir, exist_ok=True)
    os.makedirs(param_dir, exist_ok=True)

    TZ_VN = timezone(timedelta(hours=7))
    today_vn = datetime.now(TZ_VN).date()
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
                list_canbo_display.append(t_nv)
                if col_coso: dict_manv_coso[m_nv] = str(row[col_coso]).strip()
                if col_quyen: dict_manv_quyen[m_nv] = str(row[col_quyen]).strip().upper()
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
                    h_vi = str(r.get('HanhViViPham', '')).strip() if pd.notnull(r.get('HanhViViPham')) else ''
                    label_loi = f"{m_loi} - {h_vi}" if h_vi else m_loi
                    list_maloi_display.append(label_loi)
                    dict_thamso_loi[m_loi] = {
                        'MucLoi': str(r.get('MucLoi', '')).strip() if pd.notnull(r.get('MucLoi')) else '',
                        'NhomLoi': str(r.get('NhomLoi', '')).strip() if pd.notnull(r.get('NhomLoi')) else '',
                        'NhomNV': str(r.get('NhomNV', '')).strip() if pd.notnull(r.get('NhomNV')) else '',
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
                t_ngay_phat_sinh = st.date_input("2. Ngày phát sinh lỗi (Phải trước ngày hiện hành)", value=today_vn)
                
                t_ten_cb_chon = st.selectbox("3. Cán bộ liên quan", list_canbo_display)
                t_ma_nv_chon = t_ten_cb_chon.split('-')[0].strip() if '-' in t_ten_cb_chon else t_ten_cb_chon.split()[0].strip()
                
                t_coso = dict_manv_coso.get(t_ma_nv_chon, "Chưa xác định cơ sở")
                st.markdown(f"📍 **Cơ sở trực thuộc:** <span style='color: #1F4E78; font-weight: bold;'>{t_coso}</span>", unsafe_allow_html=True)
                
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
            t_mota = st.text_area("9. Mô tả lỗi (Tự động xuống dòng khi vượt quá 100 ký tự/dòng)", placeholder="Nhập chi tiết diễn biến sự việc...", height=120)
            t_dexuat = st.text_area("10. Đề xuất xử lý (Tự động xuống dòng khi vượt quá 100 ký tự/dòng)", placeholder="Nhập đề xuất hướng giải quyết...", height=120)

            submit_btn = st.form_submit_button("💾 Lưu bản ghi lỗi", use_container_width=True)

            if submit_btn:
                if t_ngay_phat_sinh > today_vn:
                    st.error("❌ Lỗi: Ngày phát sinh lỗi bắt buộc phải nhỏ hơn hoặc bằng Ngày hiện hành!")
                elif not t_mota.strip():
                    st.warning("⚠️ Vui lòng nhập mô tả chi tiết lỗi.")
                else:
                    try:
                        # Chuẩn hóa ngắt dòng 100 ký tự cho cả các trường lấy từ Thamso_loi lẫn trường nhập tay
                        new_record = {
                            'Cơ sở': t_coso,
                            'Ngày nhập': today_vn.strftime('%Y-%m-%d'),
                            'Ngày phát sinh lỗi': t_ngay_phat_sinh.strftime('%Y-%m-%d'),
                            'Trạng thái': 'Chờ phê duyệt',
                            'Tên cán bộ': t_ten_cb_chon,
                            'Mã lỗi': t_maloi,
                            'Mục lỗi': tu_dong_xuong_dong(t_mucloi, width=100),
                            'Nhóm lỗi': tu_dong_xuong_dong(t_nhomloi, width=100),
                            'Nhóm nghiệp vụ': tu_dong_xuong_dong(t_nhomnv, width=100),
                            'Hành vi vi phạm': tu_dong_xuong_dong(t_hanhvi, width=100),
                            'Mô tả lỗi': tu_dong_xuong_dong(t_mota, width=100),
                            'Đề xuất xử lý': tu_dong_xuong_dong(t_dexuat, width=100),
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

        st.dataframe(df_all[mask].fillna('').copy(), use_container_width=True)
        
        tu_choi_records = df_all[df_all['Trạng thái'] == 'Từ chối']
        if not tu_choi_records.empty:
            st.markdown("##### ✏️ Chỉnh sửa bản ghi ở trạng thái **'Từ chối'**")
            selected_idx = st.selectbox("Chọn dòng lỗi để sửa:", tu_choi_records.index, format_func=lambda i: f"Dòng {i} - CB: {tu_choi_records.loc[i, 'Tên cán bộ']}")
            row_edit = df_all.loc[selected_idx]
            with st.form("form_sua_loi"):
                val_mota_edit = "" if pd.isna(row_edit['Mô tả lỗi']) else str(row_edit['Mô tả lỗi'])
                val_dexuat_edit = "" if pd.isna(row_edit['Đề xuất xử lý']) else str(row_edit['Đề xuất xử lý'])
                s_mota = st.text_area("Cập nhật Mô tả lỗi (Tự ngắt 100 ký tự/dòng)", value=val_mota_edit, height=120)
                s_dexuat = st.text_area("Cập nhật Đề xuất xử lý (Tự ngắt 100 ký tự/dòng)", value=val_dexuat_edit, height=120)
                if st.form_submit_button("🔄 Lưu lại & Gửi duyệt lại", type="primary"):
                    df_all.loc[selected_idx, 'Mô tả lỗi'] = tu_dong_xuong_dong(s_mota, width=100)
                    df_all.loc[selected_idx, 'Đề xuất xử lý'] = tu_dong_xuong_dong(s_dexuat, width=100)
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
                    st.markdown(f"📝 **Mô tả lỗi:**\n\n{'' if pd.isna(row.get('Mô tả lỗi')) else row.get('Mô tả lỗi')}")
                    st.markdown(f"💡 **Đề xuất xử lý:**\n\n{'' if pd.isna(row.get('Đề xuất xử lý')) else row.get('Đề xuất xử lý')}")
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
        if not os.path.exists(path_data_loi):
            st.info("Chưa có dữ liệu lỗi.")
            return

        df_rep = pd.read_excel(path_data_loi, dtype={'NguoiNhap': str, 'NguoiDuyet': str})
        
        if not df_rep.empty and 'Trạng thái' in df_rep.columns:
            st.markdown("##### 🔍 Bộ lọc báo cáo:")
            
            col_d1, col_d2, col_d3 = st.columns([1.2, 1, 1])
            with col_d1:
                cot_ngay_loc = st.selectbox("Lọc thời gian theo:", options=["Ngày phát sinh lỗi", "Ngày nhập"], index=0)
            
            df_rep['_ngay_temp'] = pd.to_datetime(df_rep[cot_ngay_loc], errors='coerce').dt.date
            valid_dates = df_rep['_ngay_temp'].dropna()
            
            default_from = valid_dates.min() if not valid_dates.empty else today_vn.replace(day=1)
            default_to = valid_dates.max() if not valid_dates.empty else today_vn

            with col_d2:
                tu_ngay = st.date_input("Từ ngày:", value=default_from, format="DD/MM/YYYY")
            with col_d3:
                den_ngay = st.date_input("Đến ngày:", value=default_to, format="DD/MM/YYYY")

            if tu_ngay > den_ngay:
                st.error("❌ Lỗi chọn ngày: 'Từ ngày' không được lớn hơn 'Đến ngày'. Vui lòng chọn lại!")
                return

            col_f1, col_f2 = st.columns([1.5, 1])
            with col_f1:
                all_tt = df_rep['Trạng thái'].dropna().astype(str).unique().tolist()
                selected_tt = st.multiselect("Lọc theo Trạng thái:", options=all_tt, default=all_tt)
            with col_f2:
                search_kw = st.text_input("Tìm kiếm nhanh (Cán bộ, Cơ sở, Mã lỗi...):", "")
            
            df_filtered = df_rep.copy()
            df_filtered = df_filtered[
                (df_filtered['_ngay_temp'].notna()) & 
                (df_filtered['_ngay_temp'] >= tu_ngay) & 
                (df_filtered['_ngay_temp'] <= den_ngay)
            ]
            
            if selected_tt:
                df_filtered = df_filtered[df_filtered['Trạng thái'].astype(str).isin(selected_tt)]
            else:
                df_filtered = df_filtered.iloc[0:0]
                
            if search_kw.strip():
                df_filtered = df_filtered[df_filtered.drop(columns=['_ngay_temp']).fillna('').astype(str).apply(
                    lambda row: row.str.contains(search_kw, case=False).any(), axis=1
                )]
            
            # Xóa cột tạm và làm sạch toàn bộ giá trị trống (NaN -> "")
            df_display = df_filtered.drop(columns=['_ngay_temp']).fillna('').copy()
            
            # Chuẩn hóa đồng nhất 100 ký tự/dòng cho cả các trường từ Thamso_loi và trường nhập tay
            for col_wrap in TEXT_COLS_100:
                if col_wrap in df_display.columns:
                    df_display[col_wrap] = df_display[col_wrap].apply(lambda x: tu_dong_xuong_dong(x, width=100))

            # Tính toán số dòng lớn nhất trên tất cả các cột văn bản để căn chiều cao hàng chuẩn xác
            max_lines_per_row = []
            for _, r in df_display.iterrows():
                row_line_counts = [
                    len(str(r.get(col_c, '')).split('\n')) 
                    for col_c in TEXT_COLS_100 if col_c in df_display.columns
                ]
                max_lines_per_row.append(max([1] + row_line_counts))

            col_info, col_export = st.columns([2, 1])
            with col_info:
                st.markdown(f"📋 **Tổng số bản ghi tìm thấy:** `{len(df_display):,}` (Từ `{tu_ngay.strftime('%d/%m/%Y')}` đến `{den_ngay.strftime('%d/%m/%Y')}`)")
            
            with col_export:
                if not df_display.empty:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_display.to_excel(writer, index=False, sheet_name='BaoCao_Loi')
                        worksheet = writer.sheets['BaoCao_Loi']
                        
                        from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
                        from openpyxl.utils import get_column_letter
                        
                        thin_border = Border(
                            left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin')
                        )
                        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                        header_font = Font(color="FFFFFF", bold=True)

                        # Định dạng tiêu đề và độ rộng cột đồng nhất
                        for col_idx, col_name in enumerate(df_display.columns, 1):
                            cell = worksheet.cell(row=1, column=col_idx)
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            cell.border = thin_border
                            
                            col_letter = get_column_letter(col_idx)
                            if col_name in TEXT_COLS_100:
                                # Tính độ rộng thực tế của cột (tối đa 105 nếu có dòng chạm mốc 100 ký tự, gọn hơn nếu nội dung ngắn)
                                max_line_len = max(
                                    (max((len(line) for line in str(v).split('\n')), default=0) for v in df_display[col_name]),
                                    default=0
                                )
                                optimal_width = max(max_line_len + 4, len(str(col_name)) + 4)
                                worksheet.column_dimensions[col_letter].width = min(optimal_width, 105)
                            else:
                                col_series = df_display[col_name].fillna('')
                                max_val_len = max((len(str(v)) for v in col_series), default=0)
                                max_len = max(max_val_len, len(str(col_name))) + 4
                                worksheet.column_dimensions[col_letter].width = min(max_len, 35)
                        
                        worksheet.row_dimensions[1].height = 28

                        # Tự động căn chiều cao từng hàng theo đúng số dòng đã ngắt 100 ký tự
                        for row_idx, num_lines in enumerate(max_lines_per_row, start=2):
                            worksheet.row_dimensions[row_idx].height = max(22, num_lines * 16 + 6)
                            for col_idx in range(1, len(df_display.columns) + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.alignment = Alignment(wrap_text=True, vertical='top')
                                cell.border = thin_border

                    buffer.seek(0)
                    ten_file_excel = f"BaoCao_Loi_{tu_ngay.strftime('%d%m%Y')}_{den_ngay.strftime('%d%m%Y')}.xlsx"
                    st.download_button(
                        label="📥 Xuất báo cáo ra Excel",
                        data=buffer,
                        file_name=ten_file_excel,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )

            # HIỂN THỊ BÁO CÁO TRÊN WEB ĐỒNG NHẤT 100 KÝ TỰ/DÒNG
            if not df_display.empty:
                st.markdown("""
                <style>
                .smart-report-table { width: 100%; border-collapse: collapse; font-size: 13px; }
                .smart-report-table th { background-color: #1F4E78; color: white; padding: 8px; text-align: left; border: 1px solid #ddd; position: sticky; top: 0; white-space: nowrap; }
                .smart-report-table td { padding: 8px; border: 1px solid #ddd; vertical-align: top; white-space: pre-wrap; word-wrap: break-word; }
                .smart-report-table tr:nth-child(even) { background-color: #f9f9f9; }
                .col-100char { min-width: 260px; max-width: 720px; font-family: monospace, sans-serif; line-height: 1.4; }
                </style>
                """, unsafe_allow_html=True)

                html_rows = ["<div style='max-height: 550px; overflow: auto; border: 1px solid #ccc;'><table class='smart-report-table'><thead><tr>"]
                for col in df_display.columns:
                    html_rows.append(f"<th>{col}</th>")
                html_rows.append("</tr></thead><tbody>")

                for _, row in df_display.iterrows():
                    html_rows.append("<tr>")
                    for col in df_display.columns:
                        val = "" if pd.isna(row[col]) else str(row[col])
                        if col in TEXT_COLS_100:
                            html_rows.append(f"<td class='col-100char'>{val}</td>")
                        else:
                            html_rows.append(f"<td style='white-space: nowrap;'>{val}</td>")
                    html_rows.append("</tr>")
                html_rows.append("</tbody></table></div>")
                
                st.markdown("".join(html_rows), unsafe_allow_html=True)

                st.markdown("##### 📈 Biểu đồ phân bổ theo Trạng thái")
                st.bar_chart(df_display['Trạng thái'].value_counts())
            else:
                st.info("Không có bản ghi nào khớp với điều kiện lọc.")
        else:
            st.info("Chưa có dữ liệu vi phạm để hiển thị báo cáo.")
import streamlit as st
import pandas as pd
import io
import os

# --- TỐI ƯU CẤP ĐỘ 1: SỬ DỤNG CACHE ĐỂ KHÔNG ĐỌC LẠI FILE EXCEL GỐC NẾU KHÔNG ĐỔI ---
@st.cache_data(show_spinner=False)
def doc_excel_toi_uu(file_path):
    """Đọc file Excel siêu tốc, tự động tìm dòng tiêu đề và tối ưu bộ nhớ"""
    if not os.path.exists(file_path):
        return None
    try:
        # Đọc nhanh 20 dòng đầu để dò tìm dòng tiêu đề chính xác
        df_preview = pd.read_excel(file_path, nrows=20, header=None)
        header_idx = 0
        for i, row in df_preview.iterrows():
            row_str = str(row.values).lower()
            if 'ngày' in row_str or 'số điện' in row_str or 'tuyến' in row_str:
                header_idx = i
                break
        
        # Đọc chính thức toàn bộ file từ dòng tiêu đề đã tìm thấy
        df_clean = pd.read_excel(file_path, skiprows=header_idx)
        return df_clean
    except Exception as e:
        return None

def lay_gia_tri_sla_an_toan(df_bang_sla, ten_chi_tieu, ten_co_so, gia_tri_mac_dinh):
    """Hàm tra cứu định mức SLA an toàn tuyệt đối"""
    if df_bang_sla is None or df_bang_sla.empty:
        return gia_tri_mac_dinh
    try:
        dk1 = df_bang_sla['TÊN CHỈ TIÊU'].astype(str).str.strip().str.upper() == str(ten_chi_tieu).strip().upper()
        dk2 = df_bang_sla['CƠ SỞ'].astype(str).str.strip().str.upper() == str(ten_co_so).strip().upper()
        matched_row = df_bang_sla[dk1 & dk2]
        if not matched_row.empty:
            val = matched_row.iloc[0]['GIÁ TRỊ']
            if pd.notna(val):
                return val
    except:
        pass
    return gia_tri_mac_dinh

def show_sla_report_page():
    st.subheader("⏱️ MODULE ĐỐI CHIẾU & ĐÁNH GIÁ CHẤT LƯỢNG SLA KHO QUỸ (SIÊU TỐC)")
    st.markdown("---")
    
    base_dir = os.path.dirname(os.path.abspath(__file__))
    shared_data_dir = os.path.join(base_dir, "data_source")
    param_dir = os.path.join(base_dir, "parameters")
    os.makedirs(param_dir, exist_ok=True)
    os.makedirs(shared_data_dir, exist_ok=True)
    
    param_bang_sla_path = os.path.join(param_dir, "BANGSLA.xlsx")
    
    if not os.path.exists(param_bang_sla_path):
        danh_sach_co_so = ["Cơ sở 1", "Cơ sở 2", "HUB NVC", "CS HĐ", "CS HP"]
        sample_data = [
            {"TÊN CHỈ TIÊU": "SLADAUNGAY", "CƠ SỞ": cs, "GIÁ TRỊ": "08:30"} for cs in danh_sach_co_so
        ] + [
            {"TÊN CHỈ TIÊU": "SLACUOINGAY", "CƠ SỞ": cs, "GIÁ TRỊ": "18:15"} for cs in danh_sach_co_so
        ] + [
            {"TÊN CHỈ TIÊU": "SLATUYENDAUNGAY", "CƠ SỞ": cs, "GIÁ TRỊ": 60} for cs in danh_sach_co_so
        ] + [
            {"TÊN CHỈ TIÊU": "SLATUYENCUOINGAY", "CƠ SỞ": cs, "GIÁ TRỊ": 90} for cs in danh_sach_co_so
        ]
        pd.DataFrame(sample_data).to_excel(param_bang_sla_path, index=False)

    df_bang_sla = None
    try:
        if os.path.exists(param_bang_sla_path):
            df_bang_sla = pd.read_excel(param_bang_sla_path)
            df_bang_sla.columns = [str(c).strip().upper() for c in df_bang_sla.columns]
    except:
        pass

    file1_path = None
    file2_path = os.path.join(shared_data_dir, "Ds_TiepQuyThuGom.xlsx")
    
    if os.path.exists(shared_data_dir):
        for f in os.listdir(shared_data_dir):
            if f.lower().startswith("qlnt") and f.endswith(".xlsx"):
                file1_path = os.path.join(shared_data_dir, f)
                break
                
    missing_files = []
    if not file1_path or not os.path.exists(file1_path):
        missing_files.append("File Qlnt (bắt đầu bằng 'qlnt...xlsx')")
    if not os.path.exists(file2_path):
        missing_files.append("File 'Ds_TiepQuyThuGom.xlsx'")
        
    if missing_files:
        st.error(f"❌ Không tìm thấy file dữ liệu trong thư mục `data_source`: {', '.join(missing_files)}")
        return
        
    try:
        df1_preview = doc_excel_toi_uu(file1_path)
        if df1_preview is None:
            st.error("❌ Không thể đọc file Qlnt.")
            return
        df1_preview.columns = [str(c).strip() for c in df1_preview.columns]
        chuoi_ngay_f1 = df1_preview['Ngày'].astype(str).str.strip().str[:10].str.replace('-', '/', regex=False)
        df1_preview['Ngay_Date'] = pd.to_datetime(chuoi_ngay_f1, format='%Y/%m/%d', errors='coerce')
        months = df1_preview['Ngay_Date'].dt.strftime('%m/%Y').dropna().unique().tolist()
        months = sorted(months, reverse=True)
    except Exception as e:
        st.error(f"❌ Lỗi đọc cấu trúc ngày tháng từ file Qlnt: {e}")
        return

    st.markdown("---")
    col_chon_ky, col_nut_chay = st.columns([2, 1])
    with col_chon_ky:
        thang_chon = st.selectbox("🎯 Chọn Kỳ/Tháng Báo Cáo SLA:", months if months else ["Không rõ tháng"], key="sla_thang_chon")
        
    with col_nut_chay:
        st.markdown("<br>", unsafe_allow_html=True)
        chay_bao_cao = st.button("🚀 Chạy Phân Tích SLA (Tốc độ ánh sáng)", type="primary", use_container_width=True)

    if "sla_df_report" not in st.session_state:
        st.session_state.sla_df_report = None
    if "sla_thang_hien_tai" not in st.session_state:
        st.session_state.sla_thang_hien_tai = ""
    if "sla_has_data" not in st.session_state:
        st.session_state.sla_has_data = False

    if chay_bao_cao:
        try:
            with st.spinner("Đang tối ưu hóa và xử lý dữ liệu ở tốc độ cao..."):
                df1 = doc_excel_toi_uu(file1_path)
                df2 = doc_excel_toi_uu(file2_path)
                
                if df1 is None or df2 is None:
                    st.error("❌ Không thể đọc dữ liệu file Excel nguồn.")
                    return

                df1.columns = [str(c).strip() for c in df1.columns]
                df2.columns = [str(c).strip() for c in df2.columns]
                
                # Ép kiểu ngày tháng cực nhanh bằng vector hóa
                chuoi_ngay_f1 = df1['Ngày'].astype(str).str.strip().str[:10].str.replace('-', '/', regex=False)
                df1['Ngay_Date'] = pd.to_datetime(chuoi_ngay_f1, format='%Y/%m/%d', errors='coerce')
                
                chuoi_ngay_f2 = df2['Ngày hiệu lực'].astype(str).str.strip().str[:10].str.replace('-', '/', regex=False)
                df2['Ngay_Date'] = pd.to_datetime(chuoi_ngay_f2, format='%d/%m/%Y', errors='coerce')
                
                # Chỉ lọc đúng tháng người dùng chọn NGAY TỪ ĐẦU để giảm dung lượng tính toán
                thang_nam_f1 = df1['Ngay_Date'].dt.strftime('%m/%Y')
                thang_nam_f2 = df2['Ngay_Date'].dt.strftime('%m/%Y')
                
                df1_filtered = df1[thang_nam_f1 == thang_chon].copy()
                df2_filtered = df2[thang_nam_f2 == thang_chon].copy()
                
                if df1_filtered.empty or df2_filtered.empty:
                    st.warning(f"⚠️ Không tìm thấy dữ liệu giao dịch trong kỳ {thang_chon}.")
                    st.session_state.sla_df_report = None
                    st.session_state.sla_has_data = False
                    st.session_state.sla_thang_hien_tai = thang_chon
                    return
                
                # Tối ưu xử lý thời gian DT bằng định dạng chuẩn cố định thay vì 'mixed' chậm chạp
                df1_filtered['Thời gian DT'] = pd.to_datetime(df1_filtered['Thời gian'], errors='coerce')
                df1_filtered['Ngay_Key'] = df1_filtered['Ngay_Date'].dt.strftime('%Y-%m-%d')
                df2_filtered['Ngay_Key'] = df2_filtered['Ngay_Date'].dt.strftime('%Y-%m-%d')
                
                df1_filtered['SoDien_Key'] = df1_filtered['Số điện'].astype(str).str.strip().str.replace('.0', '', regex=False)
                df2_filtered['SoDien_Key'] = df2_filtered['Số điện'].astype(str).str.strip().str.replace('.0', '', regex=False)
                    
                df1_filtered['So_Phut_Trong_Ngay'] = df1_filtered['Thời gian DT'].dt.hour * 60 + df1_filtered['Thời gian DT'].dt.minute
                
                col_co_so_name = 'Cơ sở' if 'Cơ sở' in df2_filtered.columns else df2_filtered.columns[1]
                df2_lookup = df2_filtered[['SoDien_Key', 'Ngay_Key', col_co_so_name, 'Tên khách hàng', 'Tuyến']].copy()
                df2_lookup.rename(columns={col_co_so_name: 'CoSo_Clean'}, inplace=True)
                df2_lookup = df2_lookup.drop_duplicates(['SoDien_Key', 'Ngay_Key'])
                
                df1_filtered = df1_filtered.merge(df2_lookup, on=['SoDien_Key', 'Ngay_Key'], how='left')
                df1_filtered['CoSo_Clean'] = df1_filtered['CoSo_Clean'].fillna("Cơ sở 1")
                
                df1_filtered['Loai_NV_Lower'] = df1_filtered['Loại nghiệp vụ'].astype(str).str.lower().str.strip()
                df1_filtered['Thao_Tac_Lower'] = df1_filtered['Thao tác'].astype(str).str.lower().str.strip()
                
                ket_qua_cac_loai = []
                danh_sach_co_so_unique = df1_filtered['CoSo_Clean'].unique()

                # --- TỐI ƯU CẤP ĐỘ 2: Gom nhóm định mức SLA theo cơ sở để vector hóa một lần ---
                for cs in danh_sach_co_so_unique:
                    val_gio_t1 = str(lay_gia_tri_sla_an_toan(df_bang_sla, "SLADAUNGAY", cs, "08:30"))
                    val_gio_t2 = str(lay_gia_tri_sla_an_toan(df_bang_sla, "SLACUOINGAY", cs, "18:15"))
                    
                    try:
                        h_m1 = val_gio_t1.split(':')
                        phut_moc1 = int(h_m1[0]) * 60 + int(h_m1[1])
                    except:
                        phut_moc1 = 8 * 60 + 30
                        
                    try:
                        h_m2 = val_gio_t2.split(':')
                        phut_moc2 = int(h_m2[0]) * 60 + int(h_m2[1])
                    except:
                        phut_moc2 = 18 * 60 + 15

                    # Lọc điều kiện 1 (Đầu ngày) cho cơ sở này
                    mask_t1 = (
                        (df1_filtered['CoSo_Clean'] == cs) &
                        (df1_filtered['Loai_NV_Lower'] == 'tiếp quỹ đầu ngày') &
                        (df1_filtered['Thao_Tac_Lower'] == 'giao') &
                        (df1_filtered['So_Phut_Trong_Ngay'] > phut_moc1) &
                        (df1_filtered['So_Phut_Trong_Ngay'] < (11 * 60))
                    )
                    df_t1_sub = df1_filtered[mask_t1].copy()
                    if not df_t1_sub.empty:
                        df_t1_sub['Tuyến (file2)'] = df_t1_sub['Tuyến'].fillna("")
                        df_t1_sub['Cơ sở hiển thị'] = cs
                        df_t1_sub['Phân loại SLA'] = f"1 - Tiếp quỹ trễ đầu ngày (>{val_gio_t1})"
                        ket_qua_cac_loai.append(df_t1_sub)

                    # Lọc điều kiện 2 (Cuối ngày) cho cơ sở này
                    mask_t2 = (
                        (df1_filtered['CoSo_Clean'] == cs) &
                        (df1_filtered['Loai_NV_Lower'] == 'thu gom cuối ngày') &
                        (df1_filtered['Thao_Tac_Lower'] == 'nhận') &
                        (df1_filtered['So_Phut_Trong_Ngay'] > phut_moc2)
                    )
                    df_t2_sub = df1_filtered[mask_t2].copy()
                    if not df_t2_sub.empty:
                        df_t2_sub['Tuyến (file2)'] = df_t2_sub['Tuyến'].fillna("")
                        df_t2_sub['Cơ sở hiển thị'] = cs
                        df_t2_sub['Phân loại SLA'] = f"2 - Thu gom trễ cuối ngày (>{val_gio_t2})"
                        ket_qua_cac_loai.append(df_t2_sub)

                df1_with_tuyen = df1_filtered.dropna(subset=['Tuyến', 'Thời gian DT'])
                
                # --- TỐI ƯU CẤP ĐỘ 3: Tính toán nhóm tuyến siêu nhanh bằng Pandas Groupby ---
                if not df1_with_tuyen.empty:
                    # Tiếp quỹ vận hành dài
                    mask_d3 = (df1_with_tuyen['Loai_NV_Lower'] == 'tiếp quỹ đầu ngày') & (df1_with_tuyen['Thao_Tac_Lower'] == 'giao')
                    df_t3_raw = df1_with_tuyen[mask_d3]
                    if not df_t3_raw.empty:
                        t3_grp = df_t3_raw.groupby(['Ngay_Key', 'Tuyến', 'CoSo_Clean'])['Thời gian DT'].agg(['min', 'max']).reset_index()
                        t3_grp['Diff_Phut'] = (t3_grp['max'] - t3_grp['min']).dt.total_seconds() / 60
                        
                        for cs in danh_sach_co_so_unique:
                            max_phut_chuan = float(lay_gia_tri_sla_an_toan(df_bang_sla, "SLATUYENDAUNGAY", cs, 60))
                            t3_sub_match = t3_grp[(t3_grp['CoSo_Clean'] == cs) & (t3_grp['Diff_Phut'] > max_phut_chuan)]
                            
                            if not t3_sub_match.empty:
                                merged_t3 = df_t3_raw.merge(t3_sub_match[['Ngay_Key', 'Tuyến', 'CoSo_Clean']], on=['Ngay_Key', 'Tuyến', 'CoSo_Clean'])
                                if not merged_t3.empty:
                                    merged_t3['Tuyến (file2)'] = merged_t3['Tuyến']
                                    merged_t3['Cơ sở hiển thị'] = cs
                                    merged_t3['Phân loại SLA'] = f"3 - Tuyến Tiếp quỹ vận hành dài (>{int(max_phut_chuan)} phút)"
                                    ket_qua_cac_loai.append(merged_t3)

                    # Thu gom vận hành dài
                    mask_d4 = (df1_with_tuyen['Loai_NV_Lower'] == 'thu gom cuối ngày') & (df1_with_tuyen['Thao_Tac_Lower'] == 'nhận')
                    df_t4_raw = df1_with_tuyen[mask_d4]
                    if not df_t4_raw.empty:
                        t4_grp = df_t4_raw.groupby(['Ngay_Key', 'Tuyến', 'CoSo_Clean'])['Thời gian DT'].agg(['min', 'max']).reset_index()
                        t4_grp['Diff_Phut'] = (t4_grp['max'] - t4_grp['min']).dt.total_seconds() / 60
                        
                        for cs in danh_sach_co_so_unique:
                            max_phut_chuan = float(lay_gia_tri_sla_an_toan(df_bang_sla, "SLATUYENCUOINGAY", cs, 90))
                            t4_sub_match = t4_grp[(t4_grp['CoSo_Clean'] == cs) & (t4_grp['Diff_Phut'] > max_phut_chuan)]
                            
                            if not t4_sub_match.empty:
                                merged_t4 = df_t4_raw.merge(t4_sub_match[['Ngay_Key', 'Tuyến', 'CoSo_Clean']], on=['Ngay_Key', 'Tuyến', 'CoSo_Clean'])
                                if not merged_t4.empty:
                                    merged_t4['Tuyến (file2)'] = merged_t4['Tuyến']
                                    merged_t4['Cơ sở hiển thị'] = cs
                                    merged_t4['Phân loại SLA'] = f"4 - Tuyến Thu gom vận hành dài (>{int(max_phut_chuan)} phút)"
                                    ket_qua_cac_loai.append(merged_t4)

                if len(ket_qua_cac_loai) > 0:
                    df_final = pd.concat(ket_qua_cac_loai, ignore_index=True)
                    
                    thu_tu_co_so = {"Cơ sở 1": 1, "Cơ sở 2": 2, "HUB NVC": 3, "CS HĐ": 4, "CS HP": 5}
                    thu_tu_nghiep_vu = {
                        "1 - Tiếp quỹ trễ đầu ngày": 1,
                        "3 - Tuyến Tiếp quỹ vận hành dài": 2,
                        "2 - Thu gom trễ cuối ngày": 3,
                        "4 - Tuyến Thu gom vận hành dài": 4
                    }
                    
                    df_final['Sort_CoSo'] = df_final['Cơ sở hiển thị'].map(thu_tu_co_so).fillna(99)
                    
                    def extract_base_sla(val):
                        val_str = str(val)
                        for k in thu_tu_nghiep_vu.keys():
                            if val_str.startswith(k):
                                return thu_tu_nghiep_vu[k]
                        return 99
                        
                    df_final['Sort_NghiepVu'] = df_final['Phân loại SLA'].apply(extract_base_sla)
                    
                    df_final = df_final.sort_values(
                        by=['Sort_CoSo', 'Sort_NghiepVu', 'Ngay_Key', 'Tuyến (file2)', 'Thời gian DT'], 
                        ascending=[True, True, True, True, True]
                    )
                    df_final = df_final.drop(columns=['Sort_CoSo', 'Sort_NghiepVu'])

                    # Bảng định nghĩa các cột báo cáo (đã bỏ cột nghiệp vụ thừa)
                    cac_cot_yeu_cau = {
                        'Ngày': 'Ngày (file 1)',
                        'Tuyến (file2)': 'Tuyến (file2)',
                        'Cơ sở hiển thị': 'Cơ sở (file2)',
                        'Tên khách hàng': 'Tên khách Hàng (file 2)',
                        'Cán bộ': 'Cán bộ (file 1)',
                        'Thao tác': 'Thao tác (File 1)',
                        'Thời gian': 'Thời gian (file 1)',
                        'Phân loại SLA': 'Lý do vi phạm SLA'
                    }
                    for col in cac_cot_yeu_cau.keys():
                        if col not in df_final.columns: 
                            df_final[col] = ""
                        
                    df_report = df_final[list(cac_cot_yeu_cau.keys())].rename(columns=cac_cot_yeu_cau)
                    df_report = df_report.drop_duplicates(subset=['Ngày (file 1)', 'Cơ sở (file2)', 'Tuyến (file2)', 'Tên khách Hàng (file 2)', 'Thời gian (file 1)', 'Lý do vi phạm SLA']) 
                    
                    df_report = df_report.reset_index(drop=True)
                    df_report.insert(0, 'STT', range(1, len(df_report) + 1))
                    
                    st.session_state.sla_df_report = df_report
                    st.session_state.sla_thang_hien_tai = thang_chon
                    st.session_state.sla_has_data = True
                else:
                    st.session_state.sla_df_report = None
                    st.session_state.sla_thang_hien_tai = thang_chon
                    st.session_state.sla_has_data = False
                    
        except Exception as e:
            st.error(f"❌ Phát sinh lỗi nghiêm trọng trong quá trình xử lý: {str(e)}")
            st.session_state.sla_thang_hien_tai = ""

    # --- KHU VỰC HIỂN THỊ KẾT QUẢ ---
    if st.session_state.sla_thang_hien_tai:
        st.markdown("---")
        st.subheader(f"📊 Kết quả Phân tích Vi phạm SLA - Kỳ: {st.session_state.sla_thang_hien_tai}")
        
        if st.session_state.sla_has_data and st.session_state.sla_df_report is not None:
            df_report = st.session_state.sla_df_report
            st.warning(f"⚠️ Phát hiện tổng cộng {len(df_report)} bản ghi vi phạm trong kỳ {st.session_state.sla_thang_hien_tai}!")
            st.dataframe(df_report, use_container_width=True)
            
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                df_report.to_excel(writer, index=False, sheet_name='Bao cao vi pham SLA')
            
            file_name_xuat = f"Bao_cao_SLA_Ky_{st.session_state.sla_thang_hien_tai.replace('/', '_')}.xlsx"
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Tải Xuống File Excel Báo Cáo SLA Kỳ Này",
                data=output_buffer.getvalue(),
                file_name=file_name_xuat,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success(f"🎉 Tuyệt vời! Kỳ {st.session_state.sla_thang_hien_tai} không phát hiện trường hợp nào vi phạm SLA theo định mức.")
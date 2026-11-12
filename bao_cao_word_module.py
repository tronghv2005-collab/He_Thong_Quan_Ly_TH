import os
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def chuan_hoa_ten_nghiep_vu(nv_raw):
    """Ánh xạ tên nghiệp vụ kỹ thuật sang tên hiển thị tiếng Việt chuẩn mực, sinh động"""
    mapping = {
        'Tiếp quỹ đầu ngày': 'Tiếp quỹ đầu ngày',
        'Tiếp quỹ trong ngày': 'Tiếp quỹ trong ngày',
        'Thu gom trong ngày': 'Thu gom trong ngày',
        'Thu gom cuối ngày': 'Thu gom cuối ngày',
        'Thu hộ KH tại địa chỉ': 'Thu hộ khách hàng tại địa chỉ',
        'Chi hộ KH tại địa chỉ': 'Chi hộ khách hàng tại địa chỉ',
        'Thu khách hàng khác': 'Thu khách hàng khác',
        'Chi khách hàng khác': 'Chi khách hàng khác',
        'Thu tiền (TCTD Nộp tiền)': 'Thu tiền TCTD nộp',
        'Thu quỹ TCTD': 'Thu quỹ TCTD',
        'TCTD nộp': 'TCTD nộp tiền',
        'Chi tiền (TCTD Lĩnh tiền)': 'Chi tiền TCTD lĩnh',
        'Tiếp quỹ': 'Tiếp quỹ TCTD',
        'TCTD lĩnh': 'TCTD lĩnh tiền',
        'Chi nhánh nộp tiền': 'Chi nhánh nộp tiền',
        'Chi nhánh lĩnh': 'Chi nhánh lĩnh tiền',
        'Thu tiền cơ sở khác': 'Thu tiền cơ sở khác',
        'Chi tiền cơ sở khác': 'Chi tiền cơ sở khác'
    }
    nv_clean = str(nv_raw).strip()
    return mapping.get(nv_clean, nv_clean)

def tao_nhan_xet_sinh_dong(df):
    nhan_xet_paragraphs = []
    if df.empty:
        return ["Không có dữ liệu số liệu để phân tích chi tiết trong kỳ."]

    df_CoSo = df[~df['CoSo'].isin(['Toàn trung tâm', ''])]
    
    for coso_name, group in df_CoSo.groupby('CoSo'):
        bien_dong_list = []
        for _, row in group.iterrows():
            nv_raw = row.get('LoaiNghiepVu', '')
            nv = chuan_hoa_ten_nghiep_vu(nv_raw)
            dt = row.get('LoaiDoiTuong', '')
            dt_str = f" ({dt})" if dt and dt != 'Toàn bộ' else ""
            
            # 1. Biến động Lượt
            l_prev = row.get('Tinh_Luot_Kỳ Trước', 0)
            l_curr = row.get('Tinh_Luot_Kỳ Báo Cáo', 0)
            if l_prev > 0:
                p_luot = ((l_curr - l_prev) / l_prev) * 100
                if abs(p_luot) >= 5:
                    stt_luot = "tăng" if p_luot > 0 else "giảm"
                    bien_dong_list.append(f"- Lượt giao dịch **{nv}**{dt_str} {stt_luot} **{abs(p_luot):.1f}%** (từ {l_prev:,.0f} lên {l_curr:,.0f} lượt).")

            # 2. Biến động Doanh số Tiền VND
            v_prev = row.get('Bo_VND_Kỳ Trước', 0)
            v_curr = row.get('Bo_VND_Kỳ Báo Cáo', 0)
            if v_prev == 0:
                v_prev = row.get('VND_ThanhTien_Kỳ Trước', 0) / 1e6
                v_curr = row.get('VND_ThanhTien_Kỳ Báo Cáo', 0) / 1e6

            if v_prev > 0:
                p_vnd = ((v_curr - v_prev) / v_prev) * 100
                if abs(p_vnd) >= 5:
                    stt_vnd = "tăng" if p_vnd > 0 else "giảm"
                    bien_dong_list.append(f"- Doanh số tiền **{nv}**{dt_str} {stt_vnd} **{abs(p_vnd):.1f}%** (thay đổi từ {v_prev:,.1f} lên {v_curr:,.1f} triệu/tỷ VNĐ).")
        
        if bien_dong_list:
            unique_highlights = list(dict.fromkeys(bien_dong_list))
            text_coso = f"• **{coso_name}**: Hoạt động kho quỹ ghi nhận các điểm nhấn biến động:\n" + "\n".join(unique_highlights[:5])
        else:
            text_coso = f"• **{coso_name}**: Các chỉ tiêu về lượt giao dịch và doanh số tiền mặt duy trì ổn định, biên độ dao động thấp."
            
        nhan_xet_paragraphs.append(text_coso)

    # Phân tích Toàn trung tâm
    df_tt = df[df['CoSo'] == 'Toàn trung tâm']
    if not df_tt.empty:
        tt_list = []
        for _, row in df_tt.iterrows():
            nv_raw = row.get('LoaiNghiepVu', '')
            nv = chuan_hoa_ten_nghiep_vu(nv_raw)
            l_prev = row.get('Tinh_Luot_Kỳ Trước', 0)
            l_curr = row.get('Tinh_Luot_Kỳ Báo Cáo', 0)
            if l_prev > 0:
                p_l = ((l_curr - l_prev) / l_prev) * 100
                if abs(p_l) >= 5:
                    stt = "tăng" if p_l > 0 else "giảm"
                    tt_list.append(f"- Tổng lượt **{nv}**: {stt} **{abs(p_l):.1f}%**.")
            
            v_prev = row.get('Bo_VND_Kỳ Trước', 0)
            v_curr = row.get('Bo_VND_Kỳ Báo Cáo', 0)
            if v_prev > 0:
                p_v = ((v_curr - v_prev) / v_prev) * 100
                if abs(p_v) >= 5:
                    stt_v = "tăng" if p_v > 0 else "giảm"
                    tt_list.append(f"- Tổng doanh số tiền **{nv}**: {stt_v} **{abs(p_v):.1f}%**.")
                    
        if tt_list:
            unique_tt = list(dict.fromkeys(tt_list))
            text_tt = f"• **Toàn trung tâm (Tổng kết chung)**:\n" + "\n".join(unique_tt[:5])
            nhan_xet_paragraphs.append(text_tt)

    return nhan_xet_paragraphs

def tao_bao_cao_word_danh_gia(df_hien_thi, start_prev, end_prev, start_curr, end_curr, output_docx_path):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("BÁO CÁO ĐÁNH GIÁ BIẾN ĐỘNG HOẠT ĐỘNG KHO QUỸ & DOANH SỐ TIỀN MẶT")
    run_title.bold = True
    run_title.font.name = 'Times New Roman'
    run_title.font.size = Pt(15)
    run_title.font.color.rgb = RGBColor(31, 78, 120)

    p_time = doc.add_paragraph()
    p_time.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_time = p_time.add_run(
        f"Kỳ trước: Từ ngày {start_prev.strftime('%d/%m/%Y')} đến {end_prev.strftime('%d/%m/%Y')}  |  "
        f"Kỳ báo cáo: Từ ngày {start_curr.strftime('%d/%m/%Y')} đến {end_curr.strftime('%d/%m/%Y')}"
    )
    run_time.italic = True
    run_time.font.name = 'Times New Roman'
    run_time.font.size = Pt(11)
    
    doc.add_paragraph()

    p_heading1 = doc.add_paragraph()
    r_h1 = p_heading1.add_run("I. PHÂN TÍCH BIẾN ĐỘNG LƯỢT GIAO DỊCH & DOANH SỐ TIỀN THEO CƠ SỞ")
    r_h1.bold = True
    r_h1.font.name = 'Times New Roman'
    r_h1.font.size = Pt(12)
    r_h1.font.color.rgb = RGBColor(31, 78, 120)

    danh_sach_nhan_xet = tao_nhan_xet_sinh_dong(df_hien_thi)
    for nx_text in danh_sach_nhan_xet:
        p_nx = doc.add_paragraph()
        p_nx.paragraph_format.space_after = Pt(6)
        parts = nx_text.split("**: ")
        if len(parts) == 2:
            r_bold = p_nx.add_run(parts[0].replace("• ", "").replace("**", "") + ": ")
            r_bold.bold = True
            r_bold.font.name = 'Times New Roman'
            r_bold.font.size = Pt(11)
            
            r_content = p_nx.add_run(parts[1])
            r_content.font.name = 'Times New Roman'
            r_content.font.size = Pt(11)
        else:
            r_norm = p_nx.add_run(nx_text)
            r_norm.font.name = 'Times New Roman'
            r_norm.font.size = Pt(11)

    doc.add_paragraph()

    p_heading3 = doc.add_paragraph()
    r_h3 = p_heading3.add_run("II. ĐỀ XUẤT VÀ KIẾN NGHỊ")
    r_h3.bold = True
    r_h3.font.name = 'Times New Roman'
    r_h3.font.size = Pt(12)
    r_h3.font.color.rgb = RGBColor(31, 78, 120)

    p_conclusion = doc.add_paragraph(
        "- Các cơ sở tiếp tục duy trì nghiêm ngặt các quy định về an toàn kho quỹ, phòng chống rủi ro tiền giả, tiền lẫn loại.\n"
        "- Đối với các đơn vị ghi nhận tốc độ tăng trưởng doanh số tiền mặt và lượt giao dịch lớn, Lãnh đạo Trung tâm cần chủ động phương án điều tiết quỹ nghiệp vụ và bố trí nhân sự tiếp quỹ linh hoạt."
    )
    if p_conclusion.runs:
        p_conclusion.runs[0].font.name = 'Times New Roman'
        p_conclusion.runs[0].font.size = Pt(11)

    os.makedirs(os.path.dirname(output_docx_path), exist_ok=True)
    doc.save(output_docx_path)
    return output_docx_path
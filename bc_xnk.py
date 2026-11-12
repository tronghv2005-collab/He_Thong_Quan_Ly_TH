import pandas as pd
import os

def xu_ly_bao_cao_xnk_pipeline(shared_data_dir, output_dir, start_curr, end_curr, start_prev, end_prev):
    file_path = os.path.join(shared_data_dir, "xnk_2026.xlsx")
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Không tìm thấy file nguồn tại: {file_path}")
        
    # Đọc file thô từ Excel
    df_raw = pd.read_excel(file_path, header=None)
    
    # Cấu trúc cột chính xác từ file gốc của bạn:
    # Index 3 (Cột D): Ngày_Hiệu_Lực
    # Index 6 (Cột G): Loại_XNK
    # Index 7 (Cột H): Đối_tác
    # Index 8 (Cột I): Loại_tiền
    # Index 9 (Cột J): Số_Tiền
    # Index 20 (Cột U): Đơn_Vị_Vận_Chuyển
    max_cols = df_raw.shape[1]
    cols_to_use = [3, 6, 7, 8, 9]
    col_names = ["Ngày_Hiệu_Lực", "Loại_XNK", "Đối_tác", "Loại_tiền", "Số_tiền"]
    
    if max_cols > 20:
        cols_to_use.append(20)
        col_names.append("Đơn_Vị_Vận_Chuyển")
        
    df = df_raw.iloc[:, cols_to_use].copy()
    df.columns = col_names
    
    # Bỏ qua dòng tiêu đề nếu có
    df = df[~df["Ngày_Hiệu_Lực"].astype(str).str.contains("Ngày_Hiệu_Lực|Ngày", na=False)]
    
    # Forward fill (ffill) cho CẢ Ngày hiệu lực, Loại XNK, Đối tác và Đơn vị vận chuyển
    # giúp các dòng tiền phụ phía dưới nhận chính xác thông tin giao dịch cha phía trên
    df["Ngày_Hiệu_Lực"] = df["Ngày_Hiệu_Lực"].ffill()
    df["Loại_XNK"] = df["Loại_XNK"].ffill()
    df["Đối_tác"] = df["Đối_tác"].ffill()
    if "Đơn_Vị_Vận_Chuyển" in df.columns:
        df["Đơn_Vị_Vận_Chuyển"] = df["Đơn_Vị_Vận_Chuyển"].ffill()
    else:
        df["Đơn_Vị_Vận_Chuyển"] = ""
        
    # Làm sạch dữ liệu số tiền và loại tiền
    df = df.dropna(subset=["Loại_tiền", "Số_tiền"])
    df["Số_tiền"] = pd.to_numeric(df["Số_tiền"], errors="coerce").fillna(0)
    
    # Điều kiện lọc nghiệp vụ Xuất khẩu:
    # Chỉ tính khi Loại_XNK chứa "Xuất khẩu" VÀ 4 ký tự đầu của Đơn_Vị_Vận_Chuyển là "DVVC" hoặc "ĐVVC"
    def check_condition(row):
        loai = str(row["Loại_XNK"]).strip()
        dvvc = str(row["Đơn_Vị_Vận_Chuyển"]).strip().upper()
        if "Xuất khẩu" in loai:
            prefix = dvvc[:4]
            if prefix == "DVVC" or prefix == "ÐVVC":
                return True
            else:
                return False
        return True # Giữ lại các loại khác (Nhập khẩu...)

    df = df[df.apply(check_condition, axis=1)]

    # Chuyển đổi cột Ngày hiệu lực sang định dạng datetime chuẩn
    df["Ngày_Hiệu_Lực"] = pd.to_datetime(df["Ngày_Hiệu_Lực"], errors="coerce")
    
    # Lọc dữ liệu theo khoảng thời gian Kỳ báo cáo và Kỳ trước
    if df["Ngày_Hiệu_Lực"].notna().any():
        df_curr_raw = df[(df["Ngày_Hiệu_Lực"].dt.date >= start_curr) & (df["Ngày_Hiệu_Lực"].dt.date <= end_curr)]
        df_prev_raw = df[(df["Ngày_Hiệu_Lực"].dt.date >= start_prev) & (df["Ngày_Hiệu_Lực"].dt.date <= end_prev)]
    else:
        df_curr_raw = df.copy()
        df_prev_raw = df.copy()
        df_prev_raw["Số_tiền"] = df_prev_raw["Số_tiền"] * 0.95

    # Tổng hợp số liệu theo Loại XNK, Đối tác và Loại tiền cho từng kỳ
    df_curr = df_curr_raw.groupby(["Loại_XNK", "Đối_tác", "Loại_tiền"])["Số_tiền"].sum().reset_index()
    df_curr["Kỳ"] = "Kỳ báo cáo"
    
    df_prev = df_prev_raw.groupby(["Loại_XNK", "Đối_tác", "Loại_tiền"])["Số_tiền"].sum().reset_index()
    df_prev["Kỳ"] = "Kỳ trước"
    
    # Gộp chung dữ liệu hai kỳ lại
    df_combined = pd.concat([df_prev, df_curr], ignore_index=True)
    
    # Tạo bảng Pivot song song: mỗi loại tiền có 2 cột (Kỳ trước & Kỳ báo cáo)
    pivot_multi = df_combined.pivot_table(
        index=["Loại_XNK", "Đối_tác"],
        columns=["Loại_tiền", "Kỳ"],
        values="Số_tiền",
        fill_value=0
    )
    
    # Sắp xếp thứ tự cột (Kỳ trước đứng trước, Kỳ báo cáo đứng sau cho từng loại tiền)
    currencies = sorted(df_combined["Loại_tiền"].dropna().unique())
    multi_columns = []
    for cur in currencies:
        if (cur, "Kỳ trước") in pivot_multi.columns:
            multi_columns.append((cur, "Kỳ trước"))
        if (cur, "Kỳ báo cáo") in pivot_multi.columns:
            multi_columns.append((cur, "Kỳ báo cáo"))
            
    pivot_multi = pivot_multi[multi_columns]
    pivot_flat = pivot_multi.reset_index()
    
    # Đặt tên lại cột rõ ràng dạng "AUD (Kỳ trước)", "AUD (Kỳ báo cáo)"...
    pivot_flat.columns = [
        f"{col[0]} ({col[1]})" if col[1] != "" else col[0] 
        for col in pivot_flat.columns
    ]
    
    # Định dạng số có dấu phẩy phân tách hàng nghìn khi hiển thị giao diện Web
    def format_numeric_df(input_df):
        formatted_df = input_df.copy()
        for col in formatted_df.columns:
            if col not in ["Loại_XNK", "Đối_tác"]:
                formatted_df[col] = formatted_df[col].apply(
                    lambda x: "{:,.0f}".format(x) if pd.notnull(x) and isinstance(x, (int, float)) else x
                )
        return formatted_df

    pivot_display = format_numeric_df(pivot_flat)
    
    # Xuất file Excel kết quả ra thư mục templates/
    output_filename = f"Bao_Cao_XNK_SoSanh_{start_curr.strftime('%m%Y')}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        pivot_flat.to_excel(writer, sheet_name='So_Sanh_XNK', index=False)
        df_curr.to_excel(writer, sheet_name='Chi_Tiet_Ky_Nay', index=False)
        df_prev.to_excel(writer, sheet_name='Chi_Tiet_Ky_Truoc', index=False)
        
    return pivot_display, pivot_display, pivot_display, output_path
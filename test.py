import os
import pandas as pd

# 1. Định vị thư mục hiện tại
current_dir = os.path.dirname(os.path.abspath(__file__))
file1_path = os.path.join(current_dir, 'ds_tiepquythugom.xlsx')
file2_path = os.path.join(current_dir, 'Thamso_km.xlsx')
output_path = os.path.join(current_dir, 'test.xlsx')

print("🔄 Đang map dữ liệu: File 1 (Dòng 2) và File 2 (Dòng 1)...")

# 2. Đọc dữ liệu cố định vị trí tiêu đề dòng
df1 = pd.read_excel(file1_path, header=1)  # Cố định File 1 lấy DÒNG 2 làm tiêu đề
df2 = pd.read_excel(file2_path, header=0)  # Cố định File 2 lấy DÒNG 1 làm tiêu đề

# Loại bỏ khoảng trắng thừa ở đầu và cuối tên cột
df1.columns = df1.columns.str.strip()
df2.columns = df2.columns.str.strip()

# --- CỐ ĐỊNH TÊN CỘT CHÍNH XÁC THEO FILE THÔ ---
col_file1_kh = 'Tên khách hàng'
col_file2_pgd = 'Tên PGD'
col_file2_tuyen = 'GDCTUYẾNĐẦUNGÀY'  # Tên viết liền bạn đã sửa ở File 2
col_file2_trong = 'GDCTRONGNGÀY'     # Tên viết liền bạn đã sửa ở File 2

# Kiểm tra an toàn cho File 1 (Dòng 2)
if col_file1_kh not in df1.columns:
    print(f"❌ LỖI FILE 1: Không tìm thấy cột '{col_file1_kh}' ở dòng 2!")
    print(f"Các cột thực tế ở dòng 2 file 1 là: {list(df1.columns)}")
    exit()

# Kiểm tra an toàn cho File 2 (Dòng 1)
missing_cols = [c for c in [col_file2_pgd, col_file2_tuyen, col_file2_trong] if c not in df2.columns]
if missing_cols:
    print(f"❌ LỖI FILE 2: Không tìm thấy cột {missing_cols} ở dòng 1!")
    print(f"Các cột thực tế ở dòng 1 file 2 là: {list(df2.columns)}")
    exit()

# 3. Tạo khóa phụ ở bộ nhớ tạm để map (Giữ nguyên vẹn 100% dữ liệu gốc trong ô)
df1['map_key'] = df1[col_file1_kh].astype(str).str.strip()
df2['map_key'] = df2[col_file2_pgd].astype(str).str.strip()

# Loại bỏ trùng lặp ở file số km (nếu có) để tránh làm tăng số dòng dữ liệu thô của bạn
df2_clean = df2.drop_duplicates(subset=['map_key'])

# 4. Map dữ liệu thô (how='left' để giữ nguyên 100% số dòng của File 1)
merged_df = pd.merge(
    df1,
    df2_clean[['map_key', col_file2_tuyen, col_file2_trong]],
    on='map_key',
    how='left'
)

# 5. Lọc lấy đúng các cột kết quả theo cấu trúc bạn yêu cầu
final_df = merged_df[[col_file1_kh, col_file2_tuyen, col_file2_trong]].copy()
final_df.columns = ['Tên khách hàng', 'GDC tuyến', 'GDC trong ngày']

# 6. Xuất ra file Excel kết quả (.xlsx)
final_df.to_excel(output_path, index=False)
print(f"✅ XỬ LÝ THÀNH CÔNG! File kết quả lưu tại: {output_path}")

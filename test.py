import streamlit as st

st.subheader("💡 Ví dụ Ô Nhập Liệu Tự Động Định Dạng Dấu Phẩy")

# 1. Khởi tạo giá trị trong session_state nếu chưa có
if "nhap_so_tien" not in st.session_state:
    st.session_state.nhap_so_tien = ""

# Hàm xử lý định dạng khi giá trị thay đổi
def format_input_number():
    raw_val = st.session_state.widget_so_tien
    # Lọc bỏ tất cả các ký tự không phải số (để người dùng có thể xóa, gõ số tự nhiên)
    clean_digits = "".join([c for c in raw_val if c.isdigit()])
    
    if clean_digits:
        # Chuyển thành số nguyên và định dạng lại có dấu phẩy hàng ngàn
        formatted_val = "{:,}".format(int(clean_digits))
        st.session_state.nhap_so_tien = formatted_val
        st.session_state.widget_so_tien = formatted_val
    else:
        st.session_state.nhap_so_tien = ""
        st.session_state.widget_so_tien = ""

# 2. Sử dụng st.text_input kết hợp on_change và key
st.text_input(
    "Nhập số tiền / khối lượng:",
    value=st.session_state.nhap_so_tien,
    key="widget_so_tien",
    on_change=format_input_number,
    placeholder="Nhập số vào đây..."
)

# 3. Lấy giá trị thực tế đã làm sạch (dạng số nguyên để tính toán trong code)
raw_string = st.session_state.nhap_so_tien.replace(",", "")
gia_tri_so = int(raw_string) if raw_string.isdigit() else 0

st.write(f"🔢 Giá trị dạng số để tính toán trong Python: `{gia_tri_so}`")
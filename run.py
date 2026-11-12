import sys
import asyncio
from streamlit.web.cli import main

if __name__ == '__main__':
    # Ép Windows dùng cơ chế mạng cũ để không bị hiện lỗi vặt trên Python mới
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    sys.argv = ["streamlit", "run", "app_web.py"]  # Thay bằng tên file chính của bạn
    main()
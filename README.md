# Ứng dụng đánh giá, xếp loại giảng viên 2026

## Chạy trên máy cá nhân
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```
Sau khi chạy, mở địa chỉ localhost mà Streamlit hiển thị.

## Deploy
Có thể deploy trên Streamlit Community Cloud:
1. Đưa `app.py` và `requirements.txt` lên một repository GitHub.
2. Vào https://share.streamlit.io/ và kết nối repository.
3. Chọn `app.py` làm Main file path rồi Deploy.

## Chức năng
- Nhập danh sách giảng viên trực tiếp bằng bảng.
- Tự tính định mức, tỷ lệ hoàn thành, xếp loại viên chức và xếp TNTT.
- Tìm kiếm kết quả cá nhân.
- Tổng hợp kết quả theo 8 đơn vị trong workbook.
- Xuất dữ liệu và kết quả ra Excel.

## Lưu ý quan trọng
Ứng dụng tái hiện các công thức chính từ workbook. Do Excel có các công thức lồng ghép và điều kiện chuyên môn, cần đối chiếu kết quả trên một số hồ sơ mẫu thực tế với file Excel gốc trước khi sử dụng chính thức. Dữ liệu được giữ trong phiên chạy Streamlit; hãy thường xuyên tải Excel về để lưu trữ. Để triển khai dùng chung lâu dài, nên bổ sung cơ sở dữ liệu và đăng nhập/phân quyền.

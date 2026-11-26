# Phiếu đánh giá giảng viên cá nhân 2026

Ứng dụng Streamlit: mỗi giảng viên mở một phiếu riêng, nhập thông tin và xem kết quả tự động.

## Chạy thử trên máy
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy trực tuyến
1. Tạo repository GitHub và tải `app.py`, `requirements.txt` lên.
2. Vào https://share.streamlit.io/ và kết nối repository.
3. Chọn `app.py` làm Main file path rồi Deploy.

## Lưu ý
Ứng dụng tính toán theo các quy tắc/công thức workbook 2026 được cung cấp. Phiếu hiện không lưu dữ liệu cá nhân tập trung sau khi phiên làm việc kết thúc; người dùng tải Excel để lưu hồ sơ. Trước khi dùng chính thức, nên đối chiếu một số trường hợp mẫu với workbook gốc.

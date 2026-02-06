# Auto Video Poster
Tool tự động đăng video lên TikTok, Instagram và Facebook (via Selenium).

## Cài đặt

1. Cài đặt Python 3.8+.
2. Cài đặt thư viện:
   ```bash
   pip install -r requirements.txt
   ```

## Config
Tạo file `config.py` (hoặc chỉnh sửa trực tiếp trong main) với thông tin đăng nhập:
- **Instagram**: Username/Password.
- **TikTok**: `sessionid` cookies (Lấy từ trình duyệt sau khi đăng nhập TikTok > F12 > Application > Cookies).
- **Facebook**: Email/Pass (Sử dụng Selenium để tự động hóa trình duyệt).

## Lưu ý quan trọng
- **An toàn tài khoản**: Việc sử dụng tool tự động (đặc biệt là Instagram/TikTok) có thể bị đánh dấu là spam hoặc bot. Hãy dùng với tần suất thấp.
- **Facebook**: API Facebook cho profile cá nhân không hỗ trợ up video dễ dàng, nên tool dùng Selenium (mở trình duyệt thật).

## Chạy tool
```bash
python main.py
```

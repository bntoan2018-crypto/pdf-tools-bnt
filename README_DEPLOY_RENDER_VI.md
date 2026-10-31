# PDF Tools – Public trên Render (bản tăng cường bảo mật)

## Có trong gói
- Giao diện responsive tiếng Việt, kéo thả file và sắp xếp thứ tự nối PDF.
- Giữ nguyên các thuật toán chuyển đổi PDF hiện tại.
- OCR Tesseract đã được thêm vào Docker image với ngôn ngữ tiếng Việt (`vie`) và tiếng Anh (`eng`).
- CAPTCHA Cloudflare Turnstile kiểm tra phía máy chủ.
- Giới hạn tần suất truy cập bằng Flask-Limiter.
- Giới hạn tải lên mặc định 100 MB.
- Thư mục tạm riêng cho từng yêu cầu; xóa sau khi response đóng và dọn thư mục cũ hơn 1 giờ theo lịch kiểm tra khi có request.
- Trang Chính sách quyền riêng tư và Điều khoản sử dụng (mẫu cần rà soát).

## Trước khi Deploy: tạo khóa CAPTCHA
1. Mở https://dash.cloudflare.com/ và đăng nhập/tạo tài khoản.
2. Vào Turnstile → Add widget.
3. Chọn Managed mode, thêm hostname Render dự kiến (ví dụ `pdf-tools-public.onrender.com`). Có thể thêm hostname sau khi Render cấp URL.
4. Lưu **Site Key** và **Secret Key**.
5. Trong Render, sau khi tạo service, vào Environment và thêm:
   - `TURNSTILE_SITE_KEY` = Site Key
   - `TURNSTILE_SECRET_KEY` = Secret Key
   - `REQUIRE_CAPTCHA` = `true`
6. Nếu chưa cấu hình hai khóa trên, API xử lý sẽ trả lỗi 503 (fail-closed), không cho xử lý file.

## Deploy Render
1. Giải nén gói ZIP.
2. Upload toàn bộ nội dung bên trong thư mục `PDF_Tools_Render_Hardened` lên repository GitHub mới (các file Dockerfile, render.yaml, pdf_tools.py, requirements.txt phải ở root).
3. Render → New → Blueprint → kết nối repository.
4. Xác nhận tạo Web Service `pdf-tools-public`.
5. Chờ Docker build/deploy; kiểm tra `/api/health`.
6. Cấu hình hai biến Turnstile nêu trên trong Environment rồi redeploy.
7. Thử chuyển PDF->Word, OCR và nối PDF trên tài liệu không nhạy cảm.

## Giới hạn và ghi chú vận hành
- CAPTCHA: Cloudflare Turnstile xác minh server-side. Người dùng phải hoàn tất thử thách trước khi gửi file.
- Rate limit hiện là mặc định 8 yêu cầu/phút và 30 yêu cầu/giờ trên mỗi IP. Cấu hình `memory://` phù hợp cho một instance; nếu scale nhiều instance, chuyển sang Redis/Valkey và đặt `RATELIMIT_STORAGE_URI`.
- OCR yêu cầu Tesseract cùng dữ liệu `vie` và `eng`; Dockerfile đã cài các gói này.
- Tệp tạm được dọn sau response và dọn stale theo request định kỳ. Nếu container bị kill đột ngột, không thể đảm bảo xóa ngay; Render có thể giữ filesystem của instance trong vòng đời instance. Không upload tài liệu mật.
- Render filesystem không phải kho lưu trữ bền vững. Không lưu file người dùng hoặc lịch sử trên đó.
- Chưa có quét antivirus, hàng đợi tác vụ, giám sát nâng cao hoặc kiểm thử tải đồng thời quy mô lớn. Cần bổ sung nếu lượng người dùng tăng.
- Privacy Policy và Terms là nội dung khởi đầu, không phải tư vấn pháp lý; cần rà soát trước khi phát hành chính thức.
- Cần kiểm thử thực tế sau deploy; không thể xác nhận OCR/CPU/RAM/hiệu năng cho tới khi chạy trên Render.

## Cấu hình chính
- `MAX_CONTENT_LENGTH`: mặc định 100 MB.
- `TURNSTILE_SITE_KEY`, `TURNSTILE_SECRET_KEY`: khóa Cloudflare Turnstile.
- `REQUIRE_CAPTCHA=true`: khóa xử lý nếu CAPTCHA chưa cấu hình.
- `RATELIMIT_STORAGE_URI`: mặc định `memory://`.

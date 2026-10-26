
PHẦN 1. PHÂN TÍCH: MÔ TẢ SƠ ĐỒ CÂY CẤU TRÚC HEADING CHO FILE SRS
Hệ thống tài liệu đặc tả yêu cầu (SRS) cho tính năng "Lớp học ảo" được phân cấp 
cấu trúc Heading bằng tính năng Navigation Pane trong Microsoft Word như sau:
1. Heading 1: 1. Tổng quan dự án PTIT-Learning
   - Heading 2: 1.1 Mục tiêu hệ thống học trực tuyến
   - Heading 2: 1.2 Phạm vi tài liệu đặc tả
2. Heading 1: 2. Yêu cầu chức năng (Functional Requirements)
   - Heading 2: 2.1 Quản lý tài khoản và phân quyền người dùng
   - Heading 2: 2.2 Module phòng học ảo trực tuyến (Video/Audio Streaming)
   - Heading 2: 2.3 Tính năng chia sẻ màn hình và bảng trắng tương tác
   - Heading 2: 2.4 Quản lý tài liệu bài giảng và bài tập về nhà
3. Heading 1: 3. Yêu cầu phi chức năng (Non-Functional Requirements)
   - Heading 2: 3.1 Hiệu năng và khả năng chịu tải đồng thời
   - Heading 2: 3.2 Bảo mật dữ liệu truyền tải (Mã hóa SSL/TLS)
4. Heading 1: 4. Phụ lục và Tài liệu tham khảo
   - Heading 2: 4.1 Sơ đồ luồng dữ liệu (Data Flow Diagram)
   - Heading 2: 4.2 Thuật ngữ và viết tắt
PHẦN 2. TRIỂN KHAI NỘI DUNG CÁC FILE (WORD, EXCEL, POWERPOINT)
1. Tài liệu Word (SRS - Software Requirements Specification)
   - Cấu trúc: Thiết lập hoàn chỉnh khung SRS với ít nhất 10 tiêu đề cấp 2 (Heading 2) 
     được quản lý trực quan qua Navigation Pane.
   - Thao tác bắt buộc trước khi lưu: Cập nhật lại toàn bộ mục lục tự động 
     (Chọn Update entire table) để đảm bảo số trang chính xác tuyệt đối.
2. Bảng tính Excel (Quản lý 200 khóa học và học viên)
   - Quy tắc chuẩn hóa Database:
     + Hàng đầu tiên (Row 1) là Header Row duy nhất chứa tên các cột (Mã học viên, 
       Họ tên, Khoa, Điểm chuyên cần,...).
     + Tuyệt đối không chèn các hàng tổng cộng (Total/Average) ở giữa vùng dữ liệu.
   - Thao tác lọc đa điều kiện bằng AutoFilter (Ctrl + Shift + L):
     + Bước 1: Chọn toàn bộ bảng dữ liệu, nhấn `Ctrl + Shift + L` để bật bộ lọc.
     + Bước 2: Tại cột "Khoa", chọn điều kiện lọc là `CNTT`.
     + Bước 3: Tại cột "Điểm chuyên cần", chọn bộ lọc số (`Number Filters`) -> 
       Chọn `Less Than` (<) và nhập giá trị `7`.
3. Slide PowerPoint (Giới thiệu tính năng nền tảng)
   - Quy tắc thiết kế trực quan:
     + Áp dụng Combo màu tối chuyên nghiệp: Nền màu Xanh Navy tối (Dark Navy), 
       chữ hiển thị màu trắng để dịu mắt, điểm nhấn (Accent color) dùng màu cam sáng 
       giúp làm nổi bật các thông tin quan trọng.
     + Đảm bảo Contrast Ratio (tỷ lệ tương phản) cao giữa chữ và nền để dễ đọc.
     + Trình bày lộ trình học tập bằng Flowchart rõ ràng trên slide chuẩn 3 vùng.
PHẦN 3. KIỂM THỬ: XỬ LÝ LỖI KHI LỌC DỮ LIỆU TÊN BẰNG HÀM TRIM
Tình huống: Khi sử dụng AutoFilter để lọc danh sách sinh viên theo tên (ví dụ: "Nguyen Van A"), 
mặc dù nhìn thấy tên đó xuất hiện trong bảng nhưng kết quả lọc lại trả về trống.
Nguyên nhân và Cách khắc phục bằng hàm TRIM:
1. Nguyên nhân: 
   - Dữ liệu văn bản trong bảng tính thường bị dính các khoảng trắng thừa ở đầu hoặc cuối 
     do lỗi copy-paste từ các nguồn khác (Ví dụ: `" Nguyen Van A "` hoặc `"Nguyen Van A "`). 
   - Khi lọc tuyệt đối, Excel xem khoảng trắng thừa là một phần của chuỗi ký tự nên không khớp.
2. Cách khắc phục:
   - Bước 1: Tạo một cột phụ trợ bên cạnh (Ví dụ cột phụ ở ô F2).
   - Bước 2: Sử dụng hàm TRIM để làm sạch khoảng trắng thừa: `=TRIM(B2)` (với B2 là ô chứa tên gốc).
   - Bước 3: Kéo công thức xuống cho toàn bộ các dòng dữ liệu còn lại.
   - Bước 4: Copy cột phụ mới xử lý, bấm chuột phải chọn `Paste as Values` đè lại vào cột tên gốc 
     hoặc sử dụng cột sạch này để thực hiện thao tác AutoFilter chính xác.
# 🎓 Academic Auto-Formatter Pro

> Phần mềm desktop thông minh được phát triển bằng **C#, WinUI 3 và OpenXML SDK**, giúp tự động hóa 100% các tiêu chuẩn định dạng báo cáo học tập, bài tập lớn và đồ án tốt nghiệp theo đúng chuẩn học thuật.

---

## ✨ Các Tính Năng Nổi Bật

- **🖼️ Tự động tạo Trang Bìa chuẩn học thuật:**
  * Tự sinh trang bìa đầy đủ bố cục (Tên trường, Khoa, Loại báo cáo, Tên đề tài, Giảng viên hướng dẫn).
  * Hỗ trợ nhúng Logo trực tiếp qua luồng bộ nhớ (Byte Stream).
  * **Chế độ Cá nhân / Nhóm linh hoạt:** Tự động nhận diện danh sách thành viên, kết hợp hệ thống trường thông tin tùy chỉnh (Lớp, Khóa, Hệ) và thuật toán **GridSpan (Hợp nhất ô)** thông minh giúp tên dài không bao giờ bị lỗi xuống dòng.
  * **Vẽ khung viền trang bìa (Page Border):** Tự động tạo khung viền đôi (`BorderValues.Double`) sắc nét riêng cho trang bìa chuẩn phong cách báo cáo quốc gia.
- **📏 Chuẩn hóa Căn lề & Khổ giấy:**
  * Tùy chỉnh lề linh hoạt (Mặc định chuẩn: Trái 3.5cm, Phải/Trên/Dưới 2.0cm).
  * Hỗ trợ chế độ in 2 mặt đối xứng (`MirrorMargins`).
- **✍️ Định dạng Văn bản & Tiểu mục Thông minh:**
  * Ép kiểu chữ đồng bộ toàn văn bản (Times New Roman, Arial...).
  * Tùy chỉnh mức giãn dòng tự động (Line Spacing 1.5 Lines).
  * **Bộ lọc Heading Classifier thông minh:** Sử dụng Regex phân loại chính xác các tiêu đề lớn (Phần, Chương, Lời mở đầu, Kết luận...) và tiểu mục (`1.1`, `1.2.1`), tự động gắn Style `Heading1`, `Heading2`, `Heading3` mà không làm hỏng cấu trúc file.
- **📊 Tối ưu Bảng biểu, Hình vẽ & Phương trình:**
  * Ép toàn bộ bảng và hình ảnh căn giữa, tự động thu nhỏ kích thước hình ảnh vượt quá khung chữ xuống mức an toàn (≤ 75%).
  * Đóng gói tự động các công thức toán học (`OfficeMath`) vào bảng ẩn 3 cột kèm cơ chế đánh số thứ tự phương trình tự động (`SEQ PT`).
- **📑 Sinh Danh mục Tự động:**
  * Tự động chèn mã trường (Field Codes) cho Mục lục (`TOC`), Danh mục hình vẽ (`TOF`) và Danh mục bảng biểu (`TOT`).
- **🛡️ Chống Popup phiền toái:**
  * Tự động vô hiệu hóa hộp thoại *"Update Fields on Open"* đáng ghét của Microsoft Word.

---

## 🛠️ Công Nghệ Sử Dụng

- **Ngôn ngữ:** C# (.NET 8 / .NET 9)
- **Giao diện (UI Framework):** WinUI 3 (Windows App SDK)
- **Thư viện xử lý Word:** OpenXML SDK (`DocumentFormat.OpenXml`)

---

## 🚀 Hướng Dẫn Cài Đặt & Sử Dụng

### 1. Dành cho Người dùng (Chạy trực tiếp file .exe)
1. Tải phiên bản mới nhất tại mục **Releases** của GitHub.
2. Giải nén thư mục và chạy file **`AutoFormatter.exe`**.
3. Kéo thả file báo cáo `.docx` cần chuẩn hóa vào ô giao diện.
4. Cấu hình thông tin trang bìa, lề, font chữ rồi bấm nút **⚡ CHUẨN HÓA & XUẤT FILE**.

### 2. Dành cho Lập trình viên (Chạy từ Source Code)
- Yêu cầu máy tính cài đặt **Visual Studio 2022** kèm theo workload:
  * *Desktop development with C++*
  * *.NET Desktop Development*
  * *Windows App SDK C# Templates*
1. Clone repository về máy:
   ```bash
   git clone [https://github.com/vuminhquan26/academic-auto-formatter.git](https://github.com/vuminhquan26/academic-auto-formatter.git)
Mở file giải pháp AutoFormatter.sln bằng Visual Studio 2022.

Nhấn F5 để build và chạy ứng dụng.

📂 Cấu Trúc Dự Án
Plaintext
AutoFormatter/
│
├── MainWindow.xaml       # Giao diện chính (WinUI 3 XAML)
├── MainWindow.xaml.cs    # Logic điều khiển giao diện & sự kiện
├── DocxEngine.cs         # Động cơ xử lý XML chuyên sâu (OpenXML SDK)
├── FormatConfig.cs       # Các lớp mô hình dữ liệu cấu hình (Data Models)
└── Assets/               # Tài nguyên (Icon, Logo, Ảnh mẫu)
👥 Tác Giả & Đóng Góp
Tác giả: Vũ Minh Quân (Đại học Xây dựng Hà Nội - HUCE)

Mọi đóng góp, báo lỗi (Issues) hoặc đề xuất tính năng mới (Pull Requests) đều rất hoan nghênh!

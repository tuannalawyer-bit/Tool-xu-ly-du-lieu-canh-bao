# Hướng dẫn Sử dụng Hệ thống Tool Phân Tích Tồn Kho Cảnh Báo

Chào mừng bạn đến với hệ thống công cụ phân tích tồn kho tự động. Hệ thống này bao gồm hai công cụ độc lập phục vụ cho hai nhu cầu phân tích khác nhau: **Tool Chi Tiết** và **Tool Tổng Quan**.

---

## 📁 Cấu trúc Thư mục Dự án

```text
├── Chay Tool Chi Tiet.bat                  # File chạy nhanh Tool Chi Tiết (cục bộ)
├── Chay Tool Tong Quan.bat                 # File chạy nhanh Tool Tổng Quan (cục bộ)
├── tool_chi_tiet.py                        # Mã nguồn Python của Tool Chi Tiết
├── tool_tong_quan.py                       # Mã nguồn Python của Tool Tổng Quan
├── dashboard_chi_tiet_template.html        # Giao diện mẫu cho Dashboard Chi Tiết
├── dashboard_tong_quan_template.html       # Giao diện mẫu cho Dashboard Tổng Quan
├── development_log.csv                     # Nhật ký cập nhật tính năng hệ thống
│
└── Tool_PhanTichTonKho_Portable/           # Thư mục Đóng gói Di động (Portable - Chạy ngay)
    ├── Chay_Tool_Chi_Tiet.bat              # Kích đúp để chạy Tool Chi Tiết di động (không cần cài python/thư viện)
    ├── Chay_Tool_Tong_Quan.bat             # Kích đúp để chạy Tool Tổng Quan di động
    ├── tool_chi_tiet.py                    # Bản sao mã nguồn chi tiết
    ├── tool_tong_quan.py                   # Bản sao mã nguồn tổng quan
    ├── dashboard_chi_tiet_template.html    # Bản sao giao diện mẫu chi tiết
    ├── dashboard_tong_quan_template.html   # Bản sao giao diện mẫu tổng quan
    ├── libs/                               # Chứa các thư viện Python offline (openpyxl, pyxlsb, et_xmlfile)
    ├── python_embed/                       # Bộ thông dịch Python 3.10 rút gọn đóng gói sẵn
    └── raw mẫu/                            # Thư mục chứa bảng tra cứu (updated_data_moi.xlsb)
```

---

## 📊 1. Tool Chi Tiết (Detailed Analysis Tool)

### 📖 Giới thiệu
Tự động xử lý các file dữ liệu SAP kiểm tra tồn kho (`.xlsx`, `.csv`), tính toán các chỉ số phân tích, tự động tra cứu hạn sử dụng từ bảng tham chiếu và lọc ra danh sách sản phẩm cần kiểm tra theo đúng công thức chuẩn.

### 🚀 Cách khởi chạy
* **Chạy Portable (Khuyên dùng):** Vào thư mục `Tool_PhanTichTonKho_Portable` kích đúp vào **`Chay_Tool_Chi_Tiet.bat`**.
* **Chạy cục bộ:** Kích đúp vào **`Chay Tool Chi Tiet.bat`** ở thư mục gốc.

### 📋 Quy trình Xử lý
1. **Chọn File đầu vào:** Nhấn nút **"📂 Thêm file"** để chọn một hoặc nhiều file báo cáo tồn kho RSM.
2. **Chọn File XLSB tra cứu:** Mặc định tự tìm file `raw mẫu/updated_data_moi.xlsb`.
3. **Thiết lập Tham số:** Ngày mốc (Y1) và Số tiền tham chiếu (AF1).
4. **Bắt đầu xử lý:** Nhấn **"▶ BẮT ĐẦU XỬ LÝ"**.
5. **Kết quả:** 
   - Tệp Excel chi tiết `*(done).xlsx` chứa toàn bộ 27 cột dữ liệu đã xử lý.
   - Tệp HTML Dashboard `*_dashboard.html` hiển thị biểu đồ và bảng Pivot tương tác kèm bảng chi tiết sản phẩm.

---

## 📈 2. Tool Tổng Quan Chuỗi (Chain Overview Tool)

### 📖 Giới thiệu
Gộp dữ liệu từ nhiều tệp Excel kết quả `*(done).xlsx` để tạo ra một Dashboard duy nhất cho toàn chuỗi. Để tối ưu hóa dung lượng (giảm file HTML từ **80MB+ xuống dưới 500KB**), tool gộp dữ liệu theo nhóm: `[Vùng (RSM), ASM, Cửa hàng, Ngành hàng (MCH2), Phân loại kiểm tra]` và loại bỏ hoàn toàn danh sách chi tiết sản phẩm.

### 🚀 Cách khởi chạy
* **Chạy Portable (Khuyên dùng):** Vào thư mục `Tool_PhanTichTonKho_Portable` kích đúp vào **`Chay_Tool_Tong_Quan.bat`**.
* **Chạy cục bộ:** Kích đúp vào **`Chay Tool Tong Quan.bat`** ở thư mục gốc.

### 📋 Quy trình Xử lý
1. **Chọn File Excel đầu vào:** Nhấn **"📂 Thêm file"** để chọn nhiều tệp `*(done).xlsx` đã qua xử lý ở bước trước.
2. **Tham số:** Ngày hiển thị cập nhật (Định dạng: `DD/MM/YYYY`).
3. **Bắt đầu tổng hợp:** Nhấn **"📊 BẮT ĐẦU TỔNG HỢP"**.
4. **Kết quả:** Sinh ra một file duy nhất **`Bao_cao_tong_quan_chuoi_dashboard.html`** có màu sắc sống động, độ tương phản cao, bảng Pivot tương tác động và biểu đồ phân tích toàn chuỗi cực nhanh.

---

## ⚙️ Logic Tính toán & Quy tắc Phân loại (Tool Chi Tiết)

Chương trình tự động điền các cột từ Y đến AF:
* **Cột Y (Số ngày ko nhập):** Nếu Last GR trống/rỗng, `">90"`, hoặc hiệu ngày `>90` ngày thì điền `">90"`.
* **Cột Z (Date tham khảo):** Tra cứu từ file XLSB.
* **Cột AA (DIO D-15):** `=IF(T=0, 9999, V/T*90)`.
* **Cột AB (DIO/Date):** `=AA/Z`.
* **Cột AC (Note):** Phân loại trạng thái không giao dịch.
* **Cột AD (Note check slow):** `"Check"` nếu `DIO/Date > 5` và `Giá trị tồn > Số tiền tham chiếu`.
* **Cột AE (note hết hạn):** `"Hết hạn"` nếu quá hạn hoặc không nhập hàng mà hạn dùng ngắn.
* **Cột AF (Phân loại kiểm tra):** Lọc theo độ ưu tiên:
  1. 🔴 **Hết hạn**
  2. 🟠 **Nghi vấn tồn ảo**
  3. 🟡 **Non-moving**
  4. ⚪ **Slow moving**

---

## 🕒 Lịch sử Phiên bản
* **v1.4.5 (Chi tiết) / v1.0.0 (Tổng quan) | 09/07/2026:** Tách hệ thống làm 2 công cụ độc lập, thiết kế Dashboard Tổng quan chuỗi gộp dữ liệu tối ưu dung lượng và nâng cấp giao diện Pastel tương phản cao.

# Hướng dẫn Sử dụng Hệ thống Tool Phân Tích Tồn Kho Cảnh Báo

Chào mừng bạn đến với hệ thống công cụ phân tích tồn kho tự động. Phiên bản hiện tại đã được nâng cấp tích hợp hai chức năng vào trong một ứng dụng duy nhất bằng giao diện phân Tab trực quan: **Phân Tích Chi Tiết** và **Tổng Quan Chuỗi**.

---

## 📁 Cấu trúc Thư mục Dự án

```text
├── Chay Tool Do Du Lieu.bat                # File chạy nhanh công cụ Tích hợp (cục bộ)
├── tool_do_du_lieu.py                      # Mã nguồn Python của công cụ Tích hợp (2 Tabs)
├── dashboard_chi_tiet_template.html        # Giao diện mẫu cho Dashboard Chi Tiết (Tab 1)
├── dashboard_tong_quan_template.html       # Giao diện mẫu cho Dashboard Tổng Quan (Tab 2)
├── development_log.csv                     # Nhật ký cập nhật tính năng hệ thống
│
└── Tool_PhanTichTonKho_Portable/           # Thư mục Đóng gói Di động (Portable - Chạy ngay)
    ├── Chay_Tool.bat                       # Kích đúp để chạy nhanh công cụ di động
    ├── tool_do_du_lieu.py                  # Bản sao mã nguồn tích hợp
    ├── dashboard_chi_tiet_template.html    # Bản sao giao diện mẫu chi tiết
    ├── dashboard_tong_quan_template.html   # Bản sao giao diện mẫu tổng quan
    ├── libs/                               # Chứa các thư viện Python offline (openpyxl, pyxlsb, et_xmlfile)
    ├── python_embed/                       # Bộ thông dịch Python 3.10 rút gọn đóng gói sẵn
    └── raw mẫu/                            # Thư mục chứa bảng tra cứu (Store List, Master Article)
```

---

## 📊 1. Giao diện Tab 1: Phân Tích Chi Tiết

### 📖 Giới thiệu
Tự động xử lý các file dữ liệu SAP kiểm tra tồn kho (`.xlsx`, `.csv`), tính toán các chỉ số phân tích, tự động tra cứu hạn sử dụng từ bảng tham chiếu và lọc ra danh sách sản phẩm cần kiểm tra theo đúng công thức chuẩn.

### 🚀 Quy trình Xử lý
1. **Chọn File đầu vào:** Nhấn nút **"📂 Thêm file"** để chọn một hoặc nhiều file báo cáo tồn kho gốc.
2. **Chọn File danh mục & store list:** Chọn đường dẫn đến file `Master article 7.7.xlsx` và `Store List.xlsx`.
3. **Thiết lập Tham số:**
   - Ngày mốc (Y1) và Số tiền tham chiếu (AF1).
   - **Chuỗi báo cáo hiển thị:** Chọn tên tiêu đề chuỗi hiển thị tương ứng (ví dụ: *Tổng quan chuỗi Urban*, *Tổng quan chuỗi WIN*,...).
4. **Bắt đầu xử lý:** Nhấn **"▶ BẮT ĐẦU XỬ LÝ CHI TIẾT"**.
5. **Kết quả:** 
   - Tệp Excel chi tiết `*(done).xlsx` chứa toàn bộ 27 cột dữ liệu đã xử lý.
   - Tệp HTML Dashboard `*_dashboard.html` hiển thị biểu đồ và bảng Pivot tương tác.

---

## 📈 2. Giao diện Tab 2: Tổng Quan Chuỗi

### 📖 Giới thiệu
Gộp dữ liệu từ nhiều tệp Excel kết quả `*(done).xlsx` để tạo ra một Dashboard duy nhất cho toàn chuỗi. Để tối ưu hóa dung lượng (giảm file HTML từ **80MB+ xuống dưới 300KB**), tool gộp dữ liệu theo nhóm: `[Vùng (RSM), ASM, Cửa hàng, Ngành hàng (MCH2), Phân loại kiểm tra]` và loại bỏ hoàn toàn danh sách chi tiết sản phẩm.

### 🚀 Quy trình Xử lý
1. **Chọn File Excel đầu vào:** Nhấn **"📂 Thêm file"** để chọn nhiều tệp `*(done).xlsx` đã qua xử lý ở bước trước.
2. **Tiêu đề & Ngày cập nhật:**
   - **Chuỗi báo cáo hiển thị:** Chọn tiêu đề báo cáo hiển thị của chuỗi từ danh sách thả xuống.
   - Ngày hiển thị cập nhật (Định dạng: `DD/MM/YYYY`).
3. **Bắt đầu tổng hợp:** Nhấn **"📊 BẮT ĐẦU TỔNG HỢP CHUỖI"**.
4. **Kết quả:** Sinh ra một file duy nhất dạng **`dashboard_tong_quan_*.html`** có màu sắc sáng mặc định, bảng Pivot tương tác động và biểu đồ cột hiển thị tooltip thông tin breakdown chi tiết.

---

## ⚙️ Logic Bộ lọc Biểu đồ (Slicer Chart)
* Khi biểu đồ hiển thị ở chế độ mặc định: Nhóm theo **RSM (Vùng)**.
* Khi lọc chọn một RSM cụ thể: Nhóm theo **ASM**.
* **Mới:** Khi lọc chọn một ASM cụ thể (QLKV), biểu đồ sẽ tự động chuyển đổi hiển thị phân rã theo **từng Mã Store (Store Code)** trực thuộc quyền quản lý của ASM đó.

---

## 🕒 Lịch sử Phiên bản
* **v1.5.1 | 09/07/2026:**
  - Tích hợp thành Tab duy nhất (Notebook UI), thêm cấu hình chọn Tên báo cáo linh hoạt cho cả 2 tab.
  - Hỗ trợ định tuyến tự động file đầu vào `.csv` và đọc tệp Master bằng `openpyxl` tránh lỗi nén zip.
  - Nâng cấp bộ lọc biểu đồ ASM hiển thị breakdown theo Mã Store và nhãn chữ gọn gàng.
* **v1.4.5 | 09/07/2026:** Tách hệ thống làm 2 công cụ độc lập, thiết kế Dashboard Tổng quan chuỗi gộp dữ liệu tối ưu dung lượng và nâng cấp giao diện Pastel tương phản cao.

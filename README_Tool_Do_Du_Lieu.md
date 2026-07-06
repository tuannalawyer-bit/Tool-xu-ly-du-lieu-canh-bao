# Hướng dẫn Sử dụng Tool Phân Tích Tồn Kho (Phiên bản Di động)

## 📖 Giới thiệu
Ứng dụng tự động xử lý các file dữ liệu SAP kiểm tra tồn kho (xlsx), tính toán các chỉ số phân tích, tự động tra cứu hạn sử dụng từ bảng tham chiếu và lọc ra danh sách sản phẩm cần kiểm tra theo đúng công thức chuẩn của bạn.

Phiên bản này được đóng gói dạng **Portable (Di động)** trong thư mục `Tool_PhanTichTonKho_Portable`. Bạn có thể sao chép thư mục này sang bất kỳ máy tính Windows khác (có cài Python) và chạy trực tiếp mà không cần cài đặt thêm bất kỳ thư viện ngoài nào.

---

## 📁 Cấu trúc Thư mục Portable
* **`Chay_Tool.bat`**: File khởi chạy chính của chương trình.
* **`tool_do_du_lieu.py`**: Mã nguồn Python chính.
* **`libs/`**: Thư mục chứa các thư viện ngoại tuyến đi kèm (`openpyxl`, `pyxlsb`, `et_xmlfile`).
* **`raw mẫu/`**: Chứa bảng tra cứu mặc định (`updated_data_moi.xlsb`).

---

## 🚀 Cách Khởi chạy trên Máy tính khác

1. Sao chép toàn bộ thư mục `Tool_PhanTichTonKho_Portable` sang máy tính khác.
2. Đảm bảo máy tính đó đã cài đặt **Python** (Phiên bản 3.10 trở lên).
   * *Lưu ý: Nếu chưa cài Python, bạn có thể tải bản cài đặt nhanh tại: [python.org](https://www.python.org/downloads/)*
3. Kích đúp vào file **`Chay_Tool.bat`** để khởi chạy chương trình.

*Không cần kết nối Internet, không cần cấu hình proxy, không cần chạy lệnh cài đặt thư viện (`pip install`...), chương trình sẽ tự động nhận diện và sử dụng các thư viện offline đính kèm trong thư mục `libs`.*

---

## 📋 Quy trình Xử lý trên Giao diện

1. **Chọn File Excel đầu vào:**
   * Nhấn nút **"📂 Thêm file"** để chọn một hoặc nhiều file báo cáo tồn kho (`xlsx`) của RSM. Bạn có thể chọn nhiều file cùng lúc.
   * Để xóa file khỏi danh sách, chọn file đó và nhấn **"✖ Xóa chọn"**.

2. **Chọn File tra cứu hạn sử dụng (XLSB):**
   * Mặc định chương trình tự tìm đến file `raw mẫu/updated_data_moi.xlsb` trong thư mục.
   * Nếu muốn thay đổi file khác, nhấn nút **"📂"** ở bên phải trường nhập để chọn.

3. **Thiết lập Tham số:**
   * **Ngày mốc (Y1):** Định dạng `DD/MM/YYYY`. Mặc định lấy ngày hiện tại. Ví dụ: `01/07/2026`.
   * **Số tiền tham chiếu (AF1):** Ngưỡng giá trị tồn kho để lọc sản phẩm "Slow moving". Mặc định: `100000`.

4. **Bắt đầu xử lý:**
   * Nhấn nút **"▶ BẮT ĐẦU XỬ LÝ"**.
   * Tiến trình và nhật ký chi tiết sẽ hiển thị ở khung **Nhật ký xử lý** bên phải.
   * Kết quả sẽ được lưu thành tệp có hậu tố **`(done)`** cùng thư mục với file gốc (Ví dụ: `Đỗ Khắc Chức.xlsx` -> `Đỗ Khắc Chức (done).xlsx`).

---

## ⚙️ Logic Tính toán & Quy tắc Phân loại (Cột AF)

Chương trình tính toán và điền tự động các cột từ Y đến AF:

* **Cột Y (Số ngày ko nhập):** 
  * Nếu cột Last GR (W) là `">90"`, hoặc trống (`None`), hoặc hiệu ngày `> 90` ngày: Điền chữ `">90"`.
  * Các trường hợp còn lại: Điền số ngày cụ thể (hiệu số ngày giữa Ngày mốc và ngày Last GR).
* **Cột Z (Date tham khảo):** Tra cứu từ sheet `master article` của file XLSB theo mã sản phẩm (cột K).
* **Cột AA (DIO D-15):** Công thức `=IF(T=0, 9999, V/T*90)`.
* **Cột AB (DIO/Date):** Công thức `=AA/Z`.
* **Cột AC (Note):**
  * Nếu Last Sale (X) = `">90"` và Last GR (W) = `">90"`: Điền `"không giao dịch >90 ngày"`.
  * Nếu Last Sale (X) = `">90"` và Last GR (W) là ngày cụ thể: Điền `"Không sale 90 ngày"`.
* **Cột AD (Note check slow):** Điền `"Check"` nếu `DIO/Date > 5` và `Giá trị tồn > Số tiền tham chiếu`.
* **Cột AE (note hết hạn):**
  * Nếu Số ngày ko nhập = `">90"` và Date tham khảo < 90 ngày: Điền `"Hết hạn"`.
  * Nếu Số ngày ko nhập là số ngày cụ thể và lớn hơn Date tham khảo: Điền `"Hết hạn"`.
* **Cột AF (Phân loại kiểm tra):** Lọc và phân loại theo thứ tự ưu tiên:
  1. 🔴 **Hết hạn:** Nếu cột AE có giá trị `"Hết hạn"`.
  2. 🟠 **Nghi vấn tồn ảo:** Nếu cột AC có giá trị `"không giao dịch >90 ngày"`.
  3. 🟡 **Non-moving:** Nếu cột AC có giá trị `"Không sale 90 ngày"`.
  4. ⚪ **Slow moving:** Nếu cột AD có giá trị `"Check"`.

---
*Tool được cập nhật phiên bản v1.0.1 | Bản quyền thuộc dự án | Năm 2026*

# 📦 Quản Lý Đơn Hàng - ReadPDF

Ứng dụng desktop giúp **tự động đọc phiếu đơn hàng từ file PDF** (TikTok Shop, Shopee) và **tổng hợp vào file Excel** để quản lý đơn hàng dễ dàng.

![Python](https://img.shields.io/badge/Python-3.11+-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)

---

## ✨ Tính năng

- 🔍 **Tự động nhận diện** phiếu đơn hàng từ:
  - **TikTok Shop** (vận đơn J&T)
  - **Shopee** (vận đơn SPX - Shopee Express)
- 📄 **Đọc nhiều file PDF** cùng lúc, mỗi trang = 1 phiếu đơn
- 📊 **Trích xuất thông tin**: Mã đơn, Mã vận đơn, Mã SP, Số lượng
- 📝 **Ghi vào file Excel** (`QuanLyDonHang.xlsx`) với đánh STT tự động
- 🔁 **Loại bỏ đơn trùng** tự động theo mã đơn
- 📈 **Thống kê** số đơn theo tháng và tổng cộng
- 🖥️ **Giao diện đồ họa** (GUI) thân thiện, dễ sử dụng

---

## 📁 Cấu trúc dữ liệu Excel

| Cột          | Mô tả                              |
| ------------ | ----------------------------------- |
| STT          | Số thứ tự (đánh tự động)           |
| Nền tảng     | TikTok Shop / Shopee / Khác        |
| Mã đơn       | Mã đơn hàng trích từ PDF           |
| Mã vận đơn   | Mã vận chuyển (J&T / SPX)          |
| Mã SP        | Mã sản phẩm (VD: SP00010)          |
| Số lượng     | Tổng số lượng sản phẩm trong đơn   |
| Ngày nhận    | Ngày import vào hệ thống           |
| Tháng        | Tháng nhận đơn (dạng YYYY-MM)      |
| File PDF     | Tên file PDF gốc                   |

---

## 🚀 Cài đặt & Chạy từ mã nguồn

### Yêu cầu

- Python 3.11 trở lên
- pip (Python package manager)

### Cài đặt

```bash
# Clone repository
git clone https://github.com/<your-username>/ReadPDF.git
cd ReadPDF

# Tạo môi trường ảo (khuyến nghị)
python -m venv venv

# Kích hoạt môi trường ảo
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Cài đặt thư viện
pip install -r requirements.txt
```

### Chạy ứng dụng

```bash
python readpdf.py
```

---

## 📥 Tải file EXE (Windows)

Bạn có thể tải file `.exe` đã build sẵn từ trang [**Releases**](../../releases) của repository.

> File EXE chạy độc lập, **không cần cài Python**.

---

## 🛠️ Hướng dẫn sử dụng

1. **Mở ứng dụng** (chạy file `.exe` hoặc `python readpdf.py`)
2. Nhấn nút **"Chọn file PDF để thêm đơn"**
3. Chọn một hoặc nhiều file PDF chứa phiếu đơn hàng
4. Ứng dụng sẽ tự động:
   - Đọc và trích xuất thông tin từ mỗi trang PDF
   - Bỏ qua các đơn đã tồn tại (tránh trùng lặp)
   - Ghi dữ liệu mới vào file `QuanLyDonHang.xlsx`
5. Nhấn **"Mở file Excel"** để xem kết quả

> ⚠️ **Lưu ý:** Đóng file Excel trước khi thêm đơn mới để tránh lỗi ghi file.

---

## 📦 Thư viện sử dụng

| Thư viện      | Mục đích                         |
| ------------- | -------------------------------- |
| `pdfplumber`  | Đọc và trích xuất nội dung PDF  |
| `pandas`      | Xử lý dữ liệu và ghi file Excel|
| `openpyxl`    | Engine ghi file `.xlsx`          |
| `tkinter`     | Giao diện đồ họa (có sẵn)       |

---

## 🔧 Build thành file EXE

File EXE được tự động build bằng **GitHub Actions** mỗi khi push tag phiên bản mới. Xem file [`.github/workflows/build-exe.yml`](.github/workflows/build-exe.yml) để biết chi tiết.

### Build thủ công trên máy local

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name QuanLyDonHang readpdf.py
```

File `.exe` sẽ nằm trong thư mục `dist/`.

---

## 📄 License

MIT License - Xem file [LICENSE](LICENSE) để biết thêm chi tiết.

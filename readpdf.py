import os
import re
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import pandas as pd
import pdfplumber

# ---------------------------------------------------------------
# Đường dẫn file Excel tổng: nằm cạnh script / file exe / file .app
# ---------------------------------------------------------------
if getattr(sys, "frozen", False):
    _exe = Path(sys.executable)
    # Mac: .../Ten.app/Contents/MacOS/Ten -> lấy thư mục chứa Ten.app
    BASE = _exe.parents[3] if _exe.parent.name == "MacOS" else _exe.parent
else:
    BASE = Path(__file__).parent
MASTER = BASE / "QuanLyDonHang.xlsx"

# Tên cột dùng trong code (không dấu, đừng đổi)
COLS = ["stt", "nen_tang", "ma_don", "ma_van_don", "ma_sp", "so_luong",
        "ngay_nhan", "thang", "file_pdf"]

# Tên hiển thị trong Excel (có dấu, sửa vế phải tùy ý)
HEADERS = {
    "stt": "STT",
    "nen_tang": "Nền tảng",
    "ma_don": "Mã đơn",
    "ma_van_don": "Mã vận đơn",
    "ma_sp": "Mã SP",
    "so_luong": "Số lượng",
    "ngay_nhan": "Ngày nhận",
    "thang": "Tháng",
    "file_pdf": "File PDF",
}

RE_ORDER = re.compile(r"Order ID:\s*(\d{15,20})")
RE_TRACK = re.compile(r"\b(86\d{10})\b")
RE_SKU = re.compile(r"(SP\d{4,5})\s+(\d+)")
RE_QTY = re.compile(r"Qty Total:\s*(\d+)")

# Phiếu Shopee Express (SPX)
RE_SPX_ORDER = re.compile(r"Mã\s*đơn\s*hàng:\s*([A-Z0-9]{8,20})")
RE_SPX_TRACK = re.compile(r"Mã\s*vận\s*đơn:\s*(SPX[A-Z0-9]+)")
RE_SPX_QTY = re.compile(r"Tổng\s*SL\s*sản\s*phẩm:\s*(\d+)")


# ---------------------------------------------------------------
# Xử lý dữ liệu
# ---------------------------------------------------------------
def lay_sp_shopee(page, t):
    """Lấy mã SP và tổng số lượng từ bảng THÔNG TIN ĐƠN HÀNG của phiếu Shopee."""
    skus, qty = [], 0
    try:
        for tb in page.extract_tables():
            for row in tb:
                if not row or len(row) < 6:
                    continue
                stt = (row[0] or "").strip()
                sku = "".join((row[1] or "").split())   # ghép "SP00\n010" -> "SP00010"
                if stt.isdigit() and sku.startswith("SP"):
                    skus.append(sku)
                    sl = (row[5] or "").strip()
                    qty += int(sl) if sl.isdigit() else 0
    except Exception:
        pass
    if not qty:
        m = RE_SPX_QTY.search(t)
        qty = int(m.group(1)) if m else 0
    return ", ".join(skus), qty


def doc_pdf(path):
    """Mỗi trang = 1 phiếu. Nhận diện TikTok Shop hoặc Shopee. Trả về danh sách dict."""
    rows = []
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            t = p.extract_text() or ""

            # --- TikTok Shop (J&T) ---
            m = RE_ORDER.search(t)
            if m:
                tr = RE_TRACK.search(t)
                sk, q = RE_SKU.search(t), RE_QTY.search(t)
                rows.append({
                    "nen_tang": "TikTok Shop",
                    "ma_don": m.group(1),
                    "ma_van_don": tr.group(1) if tr else "",
                    "ma_sp": sk.group(1) if sk else "",
                    "so_luong": int(q.group(1)) if q else 0,
                })
                continue

            # --- Shopee (SPX) ---
            m = RE_SPX_ORDER.search(t)
            if m:
                tr = RE_SPX_TRACK.search(t)
                ma_sp, sl = lay_sp_shopee(p, t)
                rows.append({
                    "nen_tang": "Shopee",
                    "ma_don": m.group(1),
                    "ma_van_don": tr.group(1) if tr else "",
                    "ma_sp": ma_sp,
                    "so_luong": sl,
                })
                continue

            # --- Dự phòng: dãy 18 chữ số ---
            m = re.search(r"\b(\d{18})\b", t)
            if m:
                rows.append({
                    "nen_tang": "Khác",
                    "ma_don": m.group(1),
                    "ma_van_don": "",
                    "ma_sp": "",
                    "so_luong": 0,
                })
    return rows


def xem_thu(path):
    """Dùng để chẩn đoán: trả về số trang và đoạn chữ đầu tiên đọc được."""
    try:
        with pdfplumber.open(path) as pdf:
            t = (pdf.pages[0].extract_text() or "").replace("\n", " ")
            return f"[{len(pdf.pages)} trang] {t[:150]!r}"
    except Exception as e:
        return f"[lỗi đọc: {e}]"


def doc_master():
    """Đọc file Excel tổng, trả về DataFrame với tên cột không dấu."""
    if not MASTER.exists():
        return pd.DataFrame(columns=COLS)
    return pd.read_excel(
        MASTER, dtype={HEADERS["ma_don"]: str, HEADERS["ma_van_don"]: str}
    ).rename(columns={v: k for k, v in HEADERS.items()})


def dem_thang(df):
    thang_nay = datetime.now().strftime("%Y-%m")
    return int((df["thang"] == thang_nay).sum()) if len(df) else 0


def xu_ly(files):
    """Đọc các file PDF, ghi vào Excel tổng.
    Trả về (số đơn mới, số đơn trùng, danh sách lỗi, tổng tháng, tổng tất cả)."""
    old = doc_master()
    seen = set(old["ma_don"].astype(str))
    now = datetime.now()
    new_rows, dup, loi = [], 0, []

    for f in files:
        try:
            don_list = doc_pdf(f)
        except Exception as e:
            loi.append(f"{Path(f).name}: {e}")
            continue
        if not don_list:
            loi.append(f"{Path(f).name}: không đọc được đơn nào\n{xem_thu(f)}")
        for d in don_list:
            if d["ma_don"] in seen:
                dup += 1
                continue
            seen.add(d["ma_don"])
            d.update(ngay_nhan=now.strftime("%d/%m/%Y"),
                     thang=now.strftime("%Y-%m"),
                     file_pdf=Path(f).name)
            new_rows.append(d)

    df = pd.concat([old, pd.DataFrame(new_rows, columns=COLS)], ignore_index=True)
    df["stt"] = range(1, len(df) + 1)   # đánh lại STT liên tục
    df = df[COLS]                        # STT nằm cột đầu tiên
    df.rename(columns=HEADERS).to_excel(MASTER, index=False)  # có thể báo PermissionError
    return len(new_rows), dup, loi, dem_thang(df), len(df)


def mo_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])


# ---------------------------------------------------------------
# Giao diện
# ---------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Quản lý đơn hàng")
        root.geometry("420x330")
        root.resizable(False, False)

        tk.Label(root, text="QUẢN LÝ ĐƠN HÀNG",
                 font=("Arial", 15, "bold")).pack(pady=(18, 4))

        self.lbl_thang = tk.Label(root, text="", font=("Arial", 13))
        self.lbl_thang.pack()
        self.lbl_tong = tk.Label(root, text="", font=("Arial", 10), fg="gray")
        self.lbl_tong.pack(pady=(0, 14))

        self.btn_chon = tk.Button(root, text="Chọn file PDF để thêm đơn",
                                  font=("Arial", 12), width=26, height=2,
                                  command=self.chon_pdf)
        self.btn_chon.pack(pady=4)

        self.btn_mo = tk.Button(root, text="Mở file Excel",
                                font=("Arial", 12), width=26, height=2,
                                command=self.mo_excel)
        self.btn_mo.pack(pady=4)

        self.lbl_kq = tk.Label(root, text="", font=("Arial", 11),
                               wraplength=380, justify="center")
        self.lbl_kq.pack(pady=12)

        self.cap_nhat()

    def cap_nhat(self):
        """Cập nhật số đơn hiển thị."""
        try:
            df = doc_master()
            self.lbl_thang.config(
                text=f"Tháng {datetime.now().strftime('%m/%Y')}: {dem_thang(df)} đơn")
            self.lbl_tong.config(text=f"Tổng từ đầu đến giờ: {len(df)} đơn")
        except Exception:
            self.lbl_thang.config(
                text=f"Tháng {datetime.now().strftime('%m/%Y')}: (chưa đọc được)")
            self.lbl_tong.config(text="Có thể file Excel đang mở")

    def chon_pdf(self):
        files = filedialog.askopenfilenames(
            parent=self.root, title="Chọn file PDF đơn hàng",
            filetypes=[("PDF", "*.pdf")])
        if not files:
            return

        self.btn_chon.config(state="disabled")
        self.lbl_kq.config(text="Đang xử lý, vui lòng chờ...", fg="black")
        self.root.update_idletasks()

        try:
            moi, dup, loi, tong_thang, tong = xu_ly(files)
        except PermissionError:
            self.lbl_kq.config(text="")
            messagebox.showerror(
                "Lỗi", "File Excel đang mở. Hãy đóng file Excel rồi chọn lại.")
        except Exception as e:
            self.lbl_kq.config(text="")
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra:\n{e}")
        else:
            kq = f"Thêm mới: {moi} đơn   |   Trùng (bỏ qua): {dup} đơn"
            if loi:
                kq += "\n\nCần kiểm tra:\n" + "\n".join(loi)
            self.lbl_kq.config(text=kq, fg="red" if loi else "#1a7f37")
            self.cap_nhat()
        finally:
            self.btn_chon.config(state="normal")

    def mo_excel(self):
        if not MASTER.exists():
            messagebox.showinfo(
                "Chưa có dữ liệu",
                "Chưa có file Excel. Hãy chọn file PDF để thêm đơn trước.")
            return
        try:
            mo_file(MASTER)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không mở được file Excel:\n{e}")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

# QUẢN LÝ HỌC ONLINE Ô TÔ

Tool theo dõi học viên **chạy học lý thuyết online** (ô tô B, BSS, BSTĐ, C1, CE…) cho trung tâm đào tạo lái xe:
nạp file *Learning Report* → giờ đã học, %, môn đạt, **dự kiến ngày hoàn thành** (tối đa 8h/ngày, lịch thật),
trạng thái tự động, người chạy, phí giáo viên, báo GV khi học xong, tin nhắn + ảnh thông báo gửi Zalo, thống kê, tự sao lưu.

Phiên bản hiện tại: xem [`version.json`](version.json) · lịch sử: [`PHIEN_BAN.txt`](PHIEN_BAN.txt)

## Cài đặt (Windows)
1. Cài **Python 3** (khuyên dùng 3.11, tick *Add python.exe to PATH*). Không cần cài thêm thư viện – `openpyxl` đã kèm trong `app/vendor`.
2. Tải về: nút **Code → Download ZIP**, giải nén vào ổ D: (vd `D:\QUAN_LY_LT_ONLINE`).
3. Bấm đúp **`CHAY_TOOL.bat`** → Chrome tự mở tool. (Tùy chọn: `TAO_SHORTCUT_KHOI_DONG.bat` để tạo biểu tượng Desktop / tự chạy khi bật máy.)

Hướng dẫn chi tiết: [`HUONG_DAN_SU_DUNG.txt`](HUONG_DAN_SU_DUNG.txt)

## Cập nhật bản mới
- Trong tool: **Cài đặt & Sao lưu → Cập nhật phiên bản → Kiểm tra / Cập nhật ngay**. Tool tự sao lưu dữ liệu, tải bản mới từ kho này, chép đè phần chương trình (giữ nguyên `data/`, `backups/`) và tự khởi động lại.
- Dự phòng khi không mở được tool: tắt cửa sổ đen → chạy **`CAP_NHAT.bat`**.

## Cấu trúc
```
CHAY_TOOL.bat            chạy tool
CAP_NHAT.bat             cập nhật thủ công từ GitHub
app/server.py            máy chủ nội bộ (Python + SQLite)
app/static/index.html    giao diện
app/vendor/              openpyxl kèm sẵn
version.json             số phiên bản mới nhất (tool đọc để biết có bản mới)
data/, backups/          dữ liệu & sao lưu – KHÔNG đưa lên GitHub (.gitignore)
```

## Phát hành phiên bản mới (cho người bảo trì)
1. Sửa code, tăng `VERSION` / `VERSION_DATE` trong `app/server.py`.
2. Cập nhật `version.json` (cùng số phiên bản) và ghi thay đổi vào `PHIEN_BAN.txt`.
3. Commit + push lên nhánh `main` → các máy đang dùng sẽ thấy nút “⬆ Có bản mới”.

> Dữ liệu học viên (CCCD, họ tên…) chỉ nằm trong `data/` trên máy dùng tool, không bao giờ đưa lên GitHub.

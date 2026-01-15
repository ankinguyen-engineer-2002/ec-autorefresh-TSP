# EC Auto Refresh TSP

Project tự động hóa việc làm mới (Refresh) dữ liệu cho các file báo cáo Excel (Power Query / Data Connections) của TAP và MCN.

## 🚀 Tính năng chính

*   **Tự động Refresh:** Hỗ trợ refresh hàng loạt file Excel trong thư mục chỉ định.
*   **Safe Refresh Logic (Quan trọng):**
    *   Sử dụng quy trình: `Copy Local` -> `Refresh` -> `Cut & Move Back`.
    *   **Lợi ích:** Tránh lỗi file bị khóa (file lock) do đồng bộ OneDrive/SharePoint và đặc biệt **giữ nguyên phân quyền (NTFS Permissions)** của file gốc trên server.
*   **Retry Mechanism:** Tự động thử lại 3 lần nếu gặp lỗi khi mở file hoặc refresh.
*   **Thông báo:** Tích hợp Webhook gửi báo cáo kết quả (Thành công/Thất bại) về Power Automate/Chatbot.

## 📂 Cấu trúc dự án

*   **`TAP_refresh.py`**:
    *   Dành cho các báo cáo TAP.
    *   Nguồn: `C:\Users\Admin\NextCommerce\Data - General\TAP custom report`
*   **`MCN_refresh.py`**:
    *   Dành cho các báo cáo MCN.
    *   Nguồn: `C:\Users\Admin\NextCommerce\Data - General\MCN custom report`

## 🛠️ Yêu cầu hệ thống

*   OS: Windows (Bắt buộc).
*   Phần mềm: Microsoft Excel (đã cài đặt và active).
*   Python: 3.x
*   Thư viện Python: `pywin32` (`pip install pywin32`), `requests`.

## 📖 Cách sử dụng

Chạy trực tiếp bằng dòng lệnh hoặc cài đặt vào Task Scheduler/Airflow:

```bash
# Chạy refresh cho TAP
python TAP_refresh.py

# Chạy refresh cho MCN
python MCN_refresh.py
```

## 📝 Nhật ký thay đổi

*   **2026-01-15:**
    *   Tách riêng script cho TAP và MCN.
    *   Cập nhật logic "Move Back" (Cut) để bảo vệ phân quyền file.
    *   Push code lên GitHub.

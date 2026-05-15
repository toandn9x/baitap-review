# Hướng dẫn tạo file Excel bài tập tự đánh giá review lương

Dự án này dùng để tạo file Excel mẫu cho bài tập sau buổi đào tạo tự đánh giá review lương của team kỹ thuật.

## 1. File cần giữ lại

Thư mục chỉ cần 2 file chính:

```text
run.py
.env
```

Ý nghĩa:

| File | Mục đích |
|---|---|
| `run.py` | Script Python tạo file Excel |
| `.env` | File cấu hình toàn bộ nội dung, tiêu chí, điểm số, màu sắc và tên file output |

Khi chạy script, hệ thống sẽ tạo thêm file Excel output:

```text
Mau_bai_tap_sau_dao_tao_review_luong.xlsx
```

File Excel output có thể xóa bất kỳ lúc nào. Chạy lại `python run.py` sẽ tạo lại từ đầu.

---

## 2. Cách chạy

Mở terminal/cmd tại thư mục chứa `run.py`, sau đó chạy:

```bash
python run.py
```

Nếu chạy thành công, terminal sẽ hiện:

```text
Mau_bai_tap_sau_dao_tao_review_luong.xlsx
```

Sau đó mở file Excel vừa tạo để sử dụng.

---

## 3. Lưu ý quan trọng khi chạy lại

Mỗi lần chạy:

```bash
python run.py
```

script sẽ tạo lại file Excel từ đầu.

Điều này có nghĩa là các nội dung người dùng đã nhập trong file Excel cũ sẽ bị reset, bao gồm:

- Điểm tự chấm
- Bằng chứng công việc
- Nhận xét tự đánh giá
- Điểm quản lý
- Ghi chú quản lý
- Các phần tự luận

Nếu muốn giữ bản đã điền, hãy đổi tên hoặc copy file Excel trước khi chạy lại.

Ví dụ:

```text
Mau_bai_tap_sau_dao_tao_review_luong_NguyenVanA.xlsx
```

---

## 4. Lỗi thường gặp

### Lỗi PermissionError

Nếu gặp lỗi:

```text
PermissionError: [Errno 13] Permission denied: 'Mau_bai_tap_sau_dao_tao_review_luong.xlsx'
```

Nguyên nhân thường là file Excel đang mở.

Cách xử lý:

1. Đóng file Excel đang mở.
2. Chạy lại:

```bash
python run.py
```

---

## 5. Cấu trúc file Excel được tạo

File Excel gồm 2 sheet:

### Sheet 1: `Tu danh gia`

Đây là sheet chính để nhân sự điền bài tự đánh giá.

Bao gồm:

- Thông tin cá nhân
- Bảng tiêu chí đánh giá
- Cột tự chấm điểm 1-5
- Cột điểm quy đổi tự tính
- Cột điểm tối đa
- Cột bằng chứng công việc
- Cột nhận xét tự đánh giá
- Cột điểm quản lý
- Cột ghi chú quản lý
- Tổng điểm tự động
- Xếp loại tham khảo tự động
- Các phần tự luận sau đào tạo

### Sheet 2: `Thang diem va checklist`

Sheet tham khảo gồm:

- Thang điểm 1-5
- Danh sách tiêu chí, hệ số, điểm tối đa
- Dải xếp loại
- Checklist bài nộp

---

## 6. Công thức tính điểm

Trong Excel, mỗi tiêu chí có công thức:

```excel
Điểm quy đổi = Hệ số x Điểm tự chấm
```

Ví dụ:

```text
Hệ số = 20
Điểm tự chấm = 4
Điểm quy đổi = 20 x 4 = 80
```

Tổng điểm được tính bằng:

```excel
SUM toàn bộ điểm quy đổi
```

Xếp loại tham khảo được tính theo `RATING_RULES` trong file `.env`.

---

## 7. Ai điền cột nào?

| Cột | Người điền | Ghi chú |
|---|---|---|
| `Tự chấm (1-5)` | Nhân sự | Tự nhập điểm từ 1 đến 5 |
| `Bằng chứng công việc` | Nhân sự | Ghi task, bug, PR, hotfix, tài liệu, hỗ trợ... |
| `Nhận xét tự đánh giá` | Nhân sự | Giải thích lý do tự chấm |
| `Điểm QL` | Quản lý/Leader | Quản lý xác nhận hoặc chấm lại |
| `Ghi chú QL` | Quản lý/Leader | Lý do đồng ý/điều chỉnh, bằng chứng thiếu, kỳ vọng kỳ tới |

---

## 8. Cách chỉnh nội dung trong `.env`

Toàn bộ nội dung chính được cấu hình trong file `.env`.

### Đổi tên file Excel output

```env
OUTPUT_FILE=Mau_bai_tap_sau_dao_tao_review_luong.xlsx
```

Ví dụ đổi thành:

```env
OUTPUT_FILE=Bai_tap_tu_danh_gia.xlsx
```

### Đổi tiêu đề

```env
REPORT_TITLE=BÀI TẬP SAU ĐÀO TẠO - TỰ ĐÁNH GIÁ REVIEW LƯƠNG TEAM KỸ THUẬT
```

### Đổi hướng dẫn

```env
GUIDE_TEXT=Chấm 1-5 cho từng tiêu chí áp dụng...
```

### Đổi header bảng

```env
MAIN_HEADERS=STT|Tiêu chí|Áp dụng cho|Hệ số|Tự chấm (1-5)|Điểm quy đổi|Điểm tối đa|Bằng chứng công việc|Nhận xét tự đánh giá|Điểm QL|Ghi chú QL
```

Các cột được phân cách bằng dấu `|`.

---

## 9. Cách chỉnh tiêu chí, hệ số, điểm tối đa

Tiêu chí nằm trong dòng:

```env
CRITERIA_ROWS=...
```

Mỗi tiêu chí có cấu trúc:

```text
STT|Tên tiêu chí|Hệ số|Điểm tối đa|Áp dụng cho|Mô tả
```

Các tiêu chí được phân cách bằng dấu `;`.

Ví dụ:

```env
1|Khối lượng công việc|16|80|Áp dụng tất cả|Hoàn thành task/module đúng hạn...
```

Nếu muốn sửa hệ số tiêu chí 1 từ `16` thành `18`, sửa thành:

```env
1|Khối lượng công việc|18|90|Áp dụng tất cả|Hoàn thành task/module đúng hạn...
```

Lưu ý:

- `Điểm tối đa` thường bằng `Hệ số x 5`.
- Nếu đổi hệ số, nên đổi cả điểm tối đa cho đúng.

---

## 10. Cách chỉnh dải xếp loại

Dải xếp loại hiển thị ở sheet tham khảo nằm trong:

```env
RATING_ROWS=...
```

Rule dùng để Excel tự tính xếp loại nằm trong:

```env
RATING_RULES=...
```

Cấu trúc:

```text
Tổng điểm tối đa:minScore=label,minScore=label
```

Ví dụ:

```env
500:421=Xuất sắc,341=Giỏi,261=Khá,181=Trung bình,0=Yếu
```

Nghĩa là với thang tối đa 500 điểm:

- Từ 421: Xuất sắc
- Từ 341: Giỏi
- Từ 261: Khá
- Từ 181: Trung bình
- Dưới 181: Yếu

---

## 11. Cách chỉnh màu giao diện Excel

Các màu nằm ở cuối `.env`:

```env
COLOR_NAVY=0F172A
COLOR_BLUE=2563EB
COLOR_PURPLE=7C3AED
COLOR_CYAN=DBEAFE
COLOR_LIGHT=F8FAFC
COLOR_GREEN=DCFCE7
COLOR_YELLOW=FEF9C3
COLOR_RED=FEE2E2
COLOR_ORANGE=FFEDD5
```

Dùng mã màu HEX, không cần dấu `#`.

Ví dụ:

```env
COLOR_BLUE=1D4ED8
```

---

## 12. Quy trình khuyến nghị

1. Chỉnh nội dung trong `.env` nếu cần.
2. Chạy:

```bash
python run.py
```

3. Mở file Excel output để kiểm tra.
4. Gửi file Excel cho nhân sự điền.
5. Nếu cần tạo bản mới, đóng file Excel rồi chạy lại script.

---

## 13. Ghi chú kỹ thuật

Script sử dụng thư viện Python `openpyxl` để tạo Excel.

Nếu máy chưa có `openpyxl`, cài bằng:

```bash
pip install openpyxl
```

Thông thường trong môi trường hiện tại đã có sẵn thư viện này.

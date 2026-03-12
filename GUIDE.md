# Hướng dẫn sử dụng — Ứng dụng Xếp lịch Orientation

> Tài liệu dành cho người dùng cuối. Không yêu cầu kiến thức lập trình.

---

## Mục lục

1. [Khởi động ứng dụng](#1-khởi-động-ứng-dụng)
2. [Nhập dữ liệu (Import)](#2-nhập-dữ-liệu-import)
3. [Định dạng file Excel được hỗ trợ](#3-định-dạng-file-excel-được-hỗ-trợ)
4. [Chỉnh sửa dữ liệu trên giao diện](#4-chỉnh-sửa-dữ-liệu-trên-giao-diện)
5. [Cấu hình nhãn ca (Shift Labels)](#5-cấu-hình-nhãn-ca-shift-labels)
6. [Chạy bộ xếp lịch (Run Solver)](#6-chạy-bộ-xếp-lịch-run-solver)
7. [Xuất kết quả (Export)](#7-xuất-kết-quả-export)
8. [Các lưu ý quan trọng khi sử dụng](#8-các-lưu-ý-quan-trọng-khi-sử-dụng)
9. [Câu hỏi thường gặp (FAQ)](#9-câu-hỏi-thường-gặp-faq)

---

## 1. Khởi động ứng dụng

1. Mở cửa sổ Terminal (Command Prompt hoặc PowerShell).
2. Di chuyển đến thư mục dự án:
   ```
   cd d:\Project\schedule
   ```
3. Chạy lệnh:
   ```
   .venv\Scripts\streamlit run app.py
   ```
4. Trình duyệt sẽ tự mở trang ứng dụng (thường tại `http://localhost:8501`).

---

## 2. Nhập dữ liệu (Import)

Ở **thanh bên trái (Sidebar)**, chọn một trong 4 chế độ nhập:

| Chế độ | Mô tả | Khi nào dùng |
|--------|--------|--------------|
| **Combined workbook** | Một file `.xlsx` duy nhất chứa 3 sheet tên `hosts`, `mentors`, `students` | Khi tất cả dữ liệu nằm trong 1 file |
| **Separate files** | Ba file `.xlsx` riêng lẻ, mỗi file cho một vai trò (Host / Mentor / Student) | Khi dữ liệu được chia theo vai trò |
| **Hybrid** | Một file Combined làm nền + có thể ghi đè từng vai trò bằng file riêng | Khi muốn giữ file chung nhưng cập nhật một vài vai trò |
| **Manual entry** | Không cần file — tự nhập danh sách slot và thêm người trên giao diện | Khi lượng dữ liệu nhỏ hoặc muốn test nhanh |

### Quy trình nhập file

1. Chọn chế độ nhập.
2. Upload file(s) vào ô upload tương ứng.
3. Nhấn nút **"Load"**.
4. Ứng dụng hiển thị thông báo: *"Loaded X hosts, Y mentors, Z students"*.

> **⚠️ Lưu ý:** Nhấn "Load" sẽ **xoá toàn bộ dữ liệu cũ** (bao gồm cả kết quả xếp lịch trước đó). Hãy chắc chắn đã xuất kết quả trước khi load dữ liệu mới.

### Nhập thủ công (Manual entry)

1. Ở Sidebar, nhập danh sách time slot vào ô text (mỗi dòng một slot, ví dụ: `13/6_1`, `13/6_2`, `14/6_1`).
2. Nhấn **"Apply slots"**.
3. Chuyển sang tab Data Editor để thêm người thủ công.

---

## 3. Định dạng file Excel được hỗ trợ

Ứng dụng **tự động nhận diện** định dạng file, bạn không cần chọn thủ công.

### Định dạng A: Checkbox (dấu tích)

- **Cấu trúc:** 3 file riêng lẻ, mỗi file cho 1 vai trò.
- **Tên sheet:** Đặt theo ngày — ví dụ: `13/6`, `14-06`, `14.6` (tự động chuẩn hoá).
- **Cột:**
  - Host: `STT | Tên | Ca 1 | Ca 2 | ...`
  - Mentor/Student: `STT | Ngành | Tên | Ca 1 | Ca 2 | ...`
- **Giá trị ô ca:** `TRUE`, `FALSE`, `1`, `0`, `YES`, `NO`, `✓`, `☑`, `X`

### Định dạng B: Text (nhập văn bản)

- **Cấu trúc:** 1 file chung HOẶC 3 file riêng.
- **Tên sheet:** Đặt theo vai trò — `hosts`, `mentors`, `students` (không phân biệt hoa thường).
- **Cột:** `Tên | Ngành (nếu có) | 13/6 | 14/6 | ...`
- **Giá trị ô ngày:** Danh sách ca dạng văn bản.
  - Hỗ trợ: `ca 1,2,3` | `Không` (= không có ca nào) | `5-12` (dãy) | `9 - 10 - 11` | `2; 3; 6` | `1`

> **⚠️ Lưu ý:** Nếu Excel tự chuyển giá trị như `1,11,12` thành ngày tháng (ví dụ hiển thị `01/11/2012`), ứng dụng sẽ tự động khôi phục lại các số ca chính xác (1, 11, 12). Tuy nhiên, bạn nên kiểm tra lại dữ liệu sau khi import.

---

## 4. Chỉnh sửa dữ liệu trên giao diện

Sau khi import, dữ liệu hiển thị trong tab **"Data Editor"**.

### Cấu trúc giao diện

- **3 tab chính:** Hosts | Mentors | Students
- **Mỗi tab chia theo ngày:** Ví dụ: `13/6` | `14/6`
- **Bảng mỗi ngày gồm:**
  - Cột `Name` (tên)
  - Cột `Major` (ngành — chỉ Mentor và Student)
  - Các cột checkbox theo ca (ví dụ: "Ca 1 (8h00-8h50)")

### Các thao tác chỉnh sửa

| Thao tác | Cách thực hiện |
|----------|----------------|
| Bật/tắt ca | Click vào ô checkbox tương ứng |
| Thêm người | Nhấn nút `+` ở cuối bảng, điền Tên và Ngành |
| Xoá người | Nhấn biểu tượng xoá trên dòng tương ứng |
| Sửa tên / ngành | Click vào ô và sửa trực tiếp |

### Lưu thay đổi — QUAN TRỌNG

> **⚠️ Mọi thay đổi trên bảng chỉ là TẠM THỜI cho đến khi bạn nhấn "Apply Changes".**

- Khi có thay đổi chưa lưu, ứng dụng hiển thị banner:
  *"Unsaved edits in [vai trò] — click Apply Changes to commit all"*
- Nhấn **"Apply Changes"** để xác nhận thay đổi cho **cả 3 vai trò cùng lúc**.
- Nếu bạn chạy Solver mà chưa Apply → kết quả sẽ dựa trên dữ liệu **CŨ**, không phải dữ liệu bạn vừa sửa.

> **💡 Mẹo:** Chuyển qua lại giữa các tab (Hosts ↔ Mentors ↔ Students) sẽ **KHÔNG** làm mất thay đổi chưa lưu. Tất cả đều được giữ cho đến khi bạn nhấn Apply hoặc Load dữ liệu mới.

### Lưu dữ liệu đầu vào

Phía trên bảng có 2 nút download:

- **Save Input (JSON):** Xuất dữ liệu đã nhập dạng JSON (không bao gồm kết quả xếp lịch).
- **Save Input (Excel):** Xuất dạng Excel với 3 sheet (Hosts, Mentors, Students), ô khả dụng đánh dấu `✓`.

---

## 5. Cấu hình nhãn ca (Shift Labels)

Truy cập qua tab **"Shift Labels"** (bên cạnh Data Editor).

### Mặc định

12 ca, từ Ca 1 (8h00-8h50) đến Ca 12 (20h00-20h50).

### Cách tuỳ chỉnh

| Cách | Hướng dẫn |
|------|-----------|
| **Upload JSON** | Nhấn "Upload label mapping (JSON)" → chọn file JSON dạng `{"1": "8am", "2": "9am", ...}` |
| **Sửa trực tiếp** | Sửa bảng trên giao diện (cột Shift # và Label) |
| **Đặt lại mặc định** | Nhấn "Reset defaults" |

Sau khi sửa bảng, nhấn **"Apply Labels"** để đồng bộ.

> **⚠️ Lưu ý:**
> - Nếu bạn **xoá** một số ca khỏi bảng nhãn và nhấn Apply Labels, **tất cả dữ liệu liên quan đến ca đó sẽ bị xoá** — không thể khôi phục.
> - Nếu bạn **thêm** ca mới, ca đó sẽ được thêm vào tất cả các ngày hiện có.
> - Khi số ca thay đổi, ứng dụng hiển thị cảnh báo: *"⚠️ Shift changes detected — Remove/Add shift(s)..."*

---

## 6. Chạy bộ xếp lịch (Run Solver)

### Trước khi chạy

Đảm bảo:
- ✅ Đã có ít nhất 1 Host, 1 Mentor, 1 Student.
- ✅ Có ít nhất 1 time slot (ca).
- ✅ Đã nhấn **"Apply Changes"** nếu có chỉnh sửa.

### Chạy

Nhấn nút **"Run Solver"** (màu xanh, bên dưới Data Editor).

### Kết quả — 4 tab

#### Tab "Schedule"
- Bảng liệt kê tất cả buổi đã xếp: `time_slot | host | mentor | student | major`
- Nếu có nhiều ngày → chia thành sub-tab theo ngày.

#### Tab "Timetable"
- 3 sub-tab theo vai trò: Hosts | Mentors | Students
- Mỗi vai trò hiển thị bảng theo ngày: mỗi người ở mỗi ca làm gì.
  - Host thấy: `Mentor + Student (Ngành)`
  - Mentor thấy: `Student | Host: tên_host`
  - Student thấy: `Mentor | Host: tên_host`

#### Tab "Summary"
- **Thống kê:** Tổng số buổi | Mentor tham gia / tổng | Student được xếp / tổng
- **Cảnh báo:** Danh sách Student chưa được xếp lịch (nếu có)
- **Chi tiết theo Mentor:** Tên, Ngành, Số buổi, Trạng thái (✅ hoặc ❌)
- **Chi tiết theo Ngành:** Tên ngành, Số buổi
- **Kiểm tra ràng buộc:**
  - ✅ Host không bị trùng lịch
  - ✅ Mentor không bị trùng lịch
  - ✅ Student không bị trùng lịch
  - ✅ Mỗi Mentor có ≥ 1 buổi
  - Hiển thị `"ALL CONSTRAINTS SATISFIED"` hoặc chi tiết lỗi

#### Tab "Export"
- Nút tải xuống JSON hoặc Excel (xem mục 7).

### Khi Solver báo lỗi

| Thông báo | Nguyên nhân | Cách xử lý |
|-----------|-------------|-------------|
| *"INFEASIBLE"* | Không tìm được lịch hợp lệ | Kiểm tra: mentor và student có ca chung không? Ngành có khớp không? |
| *"FAIL: Mentor X has 0 sessions"* | Mentor X không được xếp buổi nào | Mỗi mentor **phải** có ít nhất 1 buổi — đây là ràng buộc bắt buộc. Kiểm tra lịch rảnh và ngành của mentor |

---

## 7. Xuất kết quả (Export)

### JSON (`schedule_result.json`)

Chứa toàn bộ dữ liệu đầu vào + kết quả xếp lịch:
- Danh sách time slots, shift labels
- Danh sách hosts, mentors, students (kèm lịch rảnh)
- Kết quả: mỗi buổi gồm `time_slot`, `host`, `mentor`, `student`, `major`

### Excel (`schedule_result.xlsx`)

Gồm nhiều sheet:

- **Sheet theo ngày** (ví dụ: "Ngày 13-6", "Ngày 14-6"):
  - Bảng thời khoá biểu của **Mentor**.
  - Cột: Tên mentor | Ca 1 | Ca 2 | ... (mỗi ô là tên student)

- **Sheet "Tổng hợp" (Summary):**
  - Bảng phân nhóm theo **ngành**, có mã màu.
  - Cột: Ngành (ô gộp, tô màu) | Mentor | Host | Ngày | Ca | Student
  - Header xanh dương, chữ trắng.
  - Độ rộng cột tự co giãn.

---

## 8. Các lưu ý quan trọng khi sử dụng

### 🔴 Lưu ý nghiêm trọng

1. **Nhấn "Apply Changes" trước khi chạy Solver.**
   Nếu không, Solver sẽ dùng dữ liệu cũ, không phải dữ liệu bạn vừa sửa trên giao diện.

2. **Nhấn "Load" sẽ xoá toàn bộ dữ liệu và kết quả hiện tại.**
   Luôn xuất kết quả (Export) trước khi load dữ liệu mới.

3. **Xoá ca khỏi Shift Labels sẽ mất dữ liệu vĩnh viễn.**
   Tất cả lịch rảnh của ca bị xoá sẽ biến mất, không thể hoàn tác.

4. **Mỗi Mentor phải có ít nhất 1 buổi** — đây là ràng buộc cứng của hệ thống.
   Nếu có mentor không thể xếp được (do không khớp ngành hoặc không khớp lịch), Solver sẽ báo *INFEASIBLE*.

### 🟡 Lưu ý vận hành

5. **Trùng tên + cùng ngành = gộp thành 1 người.**
   Nếu file Excel có 2 dòng "Nguyễn Văn A" với cùng ngành "Marketing", hệ thống coi đó là 1 người và gộp lịch rảnh. Nếu muốn tách, hãy phân biệt tên hoặc ngành.

6. **Trùng tên + khác ngành = 2 người riêng biệt.**
   "Nguyễn Văn A — Marketing" và "Nguyễn Văn A — HR" được coi là 2 người khác nhau. Đây là hành vi mong muốn.

7. **Ghép ngành không phân biệt hoa/thường.**
   "marketing", "Marketing", "MARKETING" đều được coi là cùng một ngành.

8. **Mentor/Student đa ngành:**
   Nếu ô ngành chứa nhiều giá trị cách bởi `,`, `;`, `|`, hoặc `/` (ví dụ: "Marketing, Sales"), hệ thống hiểu người đó thuộc nhiều ngành. Student đa ngành sẽ được ưu tiên xếp ít nhất 1 buổi cho mỗi ngành mong muốn.

9. **Sidebar hiển thị cảnh báo người có 0 ca rảnh:**
   Nếu thấy *"Mentors with 0 availability: ..."*, hãy kiểm tra lại file upload hoặc bổ sung ca trong Data Editor.

10. **Mục tiêu xếp lịch:**
    Hệ thống ưu tiên: **(1)** Tối đa số Student được xếp lịch, **(2)** Tối thiểu tổng số buổi (để tránh lãng phí). Không phải mọi Student đều được đảm bảo có lịch — chỉ có thể xếp khi có mentor phù hợp và ca trùng khớp.

### 🟢 Mẹo hữu ích

11. **Kiểm tra dữ liệu sau khi import:**
    Dù ứng dụng tự xử lý nhiều trường hợp đặc biệt (ngày tháng Excel, công thức...), bạn nên mở tab Data Editor để kiểm tra nhanh: tên, ngành, và các ca có đúng không.

12. **Xuất dữ liệu đầu vào (Save Input) trước khi sửa nhiều:**
    Dùng nút "Save Input (JSON)" hoặc "Save Input (Excel)" để tạo bản backup. Nếu sửa nhầm, có thể load lại từ file này.

13. **Dùng tab Summary sau khi chạy Solver:**
    Tab này cho bạn cái nhìn tổng quan: bao nhiêu Student chưa được xếp, mentor nào bị thiếu buổi, ngành nào ít buổi — giúp bạn quyết định có cần điều chỉnh dữ liệu và chạy lại không.

---

## 9. Câu hỏi thường gặp (FAQ)

**Q: Tôi sửa bảng nhưng chạy Solver ra kết quả cũ?**
A: Bạn chưa nhấn **"Apply Changes"**. Mọi thay đổi trên bảng chỉ có hiệu lực sau khi Apply.

**Q: Solver báo "INFEASIBLE" — phải làm sao?**
A: Kiểm tra 3 điều:
  1. Có mentor và student cùng ngành không?
  2. Họ có ca trùng nhau không?
  3. Có đủ host cho các ca đó không?

Thử bỏ bớt mentor bị ràng buộc quá chặt (ít ca rảnh, ngành hiếm) hoặc thêm ca cho người có 0 availability.

**Q: File Excel của tôi bị lỗi khi import?**
A: Kiểm tra:
  - Tên sheet phải đúng chuẩn (theo ngày hoặc theo vai trò).
  - Không có sheet trống hoặc sheet có tên đặc biệt.
  - Hàng header phải nằm ở **dòng 1 hoặc dòng 2** — ứng dụng tự phát hiện nhưng không hỗ trợ header ở dòng 3 trở xuống.

**Q: Tôi có 2 người trùng tên nhưng khác vai trò (1 là Mentor, 1 là Student)?**
A: Hoàn toàn hợp lệ. Hệ thống phân biệt theo vai trò và đảm bảo cùng 1 người không bị xếp vào 2 buổi cùng ca (ràng buộc cross-role).

**Q: Kết quả Excel chỉ hiện lịch Mentor, không có lịch Host và Student?**
A: Đúng thiết kế. File Excel xuất ra tập trung vào **thời khoá biểu Mentor** (vì mentor là trục chính) + sheet Tổng hợp. Nếu cần lịch chi tiết theo Host hoặc Student, xem tab **Timetable** trên giao diện hoặc xuất file JSON.

**Q: Tôi muốn thêm ca mới (ví dụ ca 13)?**
A: Vào tab **Shift Labels** → thêm dòng mới với Shift # = 13 và tên nhãn → nhấn **"Apply Labels"**. Ca mới sẽ xuất hiện trong tất cả các ngày.

---

*Phiên bản tài liệu: 1.0 — Cập nhật theo ứng dụng hiện tại.*

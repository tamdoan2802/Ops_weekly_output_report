---
name: Output_weekly_report
description: Tự động hóa quy trình phân tích và tạo báo cáo nhân sự, năng suất và khối lượng công việc hàng tuần (Weekly HR & Workload Report). Kích hoạt khi người dùng yêu cầu xuất báo cáo tuần, phân tích hiệu suất nhân sự, hoặc tính toán khối lượng công việc (leadtime, backlog, task complexity).
---

# Output Weekly Report (Báo Cáo Năng Suất & Nhân Sự Hàng Tuần)

Skill này được thiết kế để chuẩn hóa và tự động hóa toàn bộ quy trình xuất báo cáo Weekly HR & Workload cho dự án, xử lý phân cấp dữ liệu theo **Team → Customer → Member**.

## 1. Mục Đích Của Skill
Xử lý dữ liệu thô từ nhiều nguồn (Excel files) để tạo ra một file Markdown hoàn chỉnh tổng hợp hiệu suất thiết kế, chất lượng chuyên cần (đi muộn, về sớm, nghỉ phép, OT), Backlog công việc và thông tin về thời gian xử lý (Leadtime & Task Complexity).

## 2. Nguồn Dữ Luận Đầu Vào
Skill yêu cầu các tệp dữ liệu sau nằm trong `G:\My Drive\Dữ liệu nhân sự\Data\`:
- **Workload**: `Workload/Report_Past Month.xlsx` (Hoặc file mới nhất tương ứng). Gồm dữ liệu Tasks, Jobs.
- **Attendance**: `Attendance/HR_Fact_Attendance.xlsx`. Gồm sheets: `FACT_Attendance_Daily` (dữ liệu chấm công hàng ngày), `Req_Leave` (Dữ liệu xin nghỉ phép), `Req_OT` (Dữ liệu xin tăng ca).
- **KPI**: Mapping KPI chuẩn được thiết lập trong script.

## 3. Quy Trình Thực Hiện (Workflow)

Để thực thi workflow xuất báo cáo, Agent cần thực hiện:

### Bước 1: Khởi Chạy Script Python
Chạy script chính để thu thập và xử lý toàn bộ dữ liệu:
```bash
python "G:\My Drive\Dữ liệu nhân sự\.agents\skills\Output_weekly_report\scripts\summarize_weekly_production.py"
```

### Bước 2: Kịch Bản Xử Lý Bên Trong Script
Script sẽ tự động:
1. **Lọc Thời Gian**: Xác định 5 tuần làm việc gần nhất dựa vào ngày hiện tại, xác định tuần đích W0 (tuần hiện tại/vừa qua).
2. **Override Completion Logic**: Tự động xác định Job hoàn thành thông qua Task Check đối với những khách hàng đặc biệt (RES Engineering, Prime Design).
3. **Phân Rã Team -> Customer -> Member**: 
   - Đếm khối lượng công việc (Design, Check, Review & Fix, v.v.).
   - Đếm Units, Sqm, Layouts.
4. **Xử Lý Chuyên Cần (Attendance)**:
   - Tính **Ngày LV (Số ngày làm việc thực tế)** = N_Full + N_Half.
   - Tính **Giờ TT (Giờ thực tế)** dựa trên dữ liệu bảng công chi tiết (`Số giờ làm việc thực tế`).
   - Tính **Giờ TC (Giờ tiêu chuẩn)** = N_Full * 8 + N_Half * 4.
   - Thống kê nghỉ phép, đi muộn, tăng ca (OT).
5. **Backlog & Aging**: Thống kê và phân nhóm job tồn đọng theo nhóm số ngày chờ xử lý.
6. **Task Complexity Profile**: Thống kê số giờ đầu tư thực tế (Invested Seconds) trung vị và trung bình cho từng loại Task.
7. **Kết Xuất Word (.docx)**: Lưu báo cáo vào thư mục `G:\My Drive\Dữ liệu nhân sự\.agents\skills\Ops_weekly_output_report\reports\` (ví dụ: `Weekly_Performance_Report_20260802.docx`).

### Bước 3: Đọc và Trình Bày
- Sau khi script chạy xong, Agent sẽ tìm đường dẫn file báo cáo mới được xuất.
- Đọc file và thông báo tóm tắt cho người dùng kết quả hoặc đính kèm link tới file báo cáo.

## 4. Các File Script Bao Gồm
- `scripts/summarize_weekly_production.py`: Script xử lý chính, chứa toàn bộ logic chuẩn hóa dữ liệu, logic phân cấp thành viên/khách hàng/team và xuất Markdown.
- `scripts/check_fact_attendance.py`: Script công cụ nhỏ hỗ trợ kiểm tra cấu trúc dữ liệu nếu cần debug.

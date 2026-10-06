# QAStation & HFPA Data Processing Tool

Hệ thống tự động hóa toàn diện quy trình xử lý, đối chiếu và lập báo cáo chất lượng theo tài liệu **Standard Operating Procedure (SOP)** chính thức [`Database/2026 FTT&HFPA EN.pdf`](Database/2026%20FTT&HFPA%20EN.pdf) và đặc tả kỹ thuật [`QAStation_Tool_Agent_Spec.md`](QAStation_Tool_Agent_Spec.md).

---

## 1. Nguồn dữ liệu đầu vào (Input Files)

1. **HFPA (Quality Tracking)** — Thư mục `Input/FTT/`:
   - Dữ liệu trạm kiểm tra tại các nhà máy (`JV`, `JV2`, `VH`, `VH2`).
   - Cấu trúc: `InspDate`, `Line` (bao gồm Plant và Line), `ShoeName`, `InspQty`, `FailQty`, `Station`, `Issues`, `IssueQty`.
2. **HFPA (Mes410)** — Thư mục `Input/HFPA/`:
   - Dữ liệu kiểm toán chất lượng từ hệ thống MES410 của 4 nhà máy.
   - Cấu trúc: `Audit Date`, `Plant`, `Line`, `Model 2`, `Audit Sample Size`, `Defective Q'ty(pair)`, `Defect Q'ty`, các cột loại lỗi chi tiết.
3. **Mẫu cơ sở dữ liệu (Database Template)** — Thư mục `Database/`:
   - [`Database/HFPA_Template.xlsx`](Database/HFPA_Template.xlsx): File cơ sở dữ liệu mẫu gồm 3 Sheet: `HFPA`, `FTT`, `BC Grade`.
   - [`Database/HFPA_Template.pptx`](Database/HFPA_Template.pptx): File thuyết trình PowerPoint liên kết dữ liệu.

---

## 2. Quy trình xử lý tự động theo SOP & Đặc tả kỹ thuật

1. **Làm sạch dữ liệu & Tạo Pivot Table 1st** (SOP Trang 1–5, Spec 3.1-3.2, 4.1):
   - Tự động nhận diện và gán cột mã nhà máy `FTY` (`JV`, `JV2`, `VH`, `VH2`).
   - Tạo bảng Pivot #1 với thứ tự các trường chuẩn: `InspDate` $\rightarrow$ `FTY` $\rightarrow$ `Station` $\rightarrow$ **`Plant`** $\rightarrow$ `Line` $\rightarrow$ `Shoename` $\rightarrow$ `InspQty` $\rightarrow$ `FailQty` $\rightarrow$ `DR%`.
   - Bảng phẳng (Tabular form), lặp lại nhãn dòng (Repeat All Item Labels), tắt Subtotals và Grand Totals.

2. **Đối chiếu HFPA (Quality Tracking) vs HFPA (Mes410)** (SOP Trang 6–10, Spec 4.2-4.4):
   - Ánh xạ trường tương đương:
     - `Inspection Qty` $\leftrightarrow$ `Audit Sample Size`
     - `FailQty` $\leftrightarrow$ `Defective Q'ty (pairs)`
     - `InspDate` $\leftrightarrow$ `Audit Date (dd-mmm-yyyy)`
   - Đối chiếu chéo theo từng Nhà máy (`FTY`), từng Xưởng (`Plant`), từng Chuyền (`Line`), từng Ngày (`Date`) và Model.

3. **Hiệu chỉnh sai lệch theo chuẩn Mes410 & Ghi Audit Trail** (SOP Trang 11–15, Spec 4.5, 12, 16):
   - Khi phát hiện sai lệch giữa trạm kiểm tra và MES410, lấy số liệu MES410 làm chuẩn tham chiếu gốc:
     - Cập nhật `InspQty` theo `Audit Sample Size`.
     - Cập nhật `FailQty` theo `Defective Q'ty (pairs)`.
     - Điều chỉnh tỷ lệ các lỗi chi tiết `IssueQty` khớp đúng với MES410.
   - Ghi lại vết thay đổi (Audit Trail) chi tiết tại `logs/reconciliation_audit_trail.csv`.

4. **Lọc Top 5 Model có số lỗi cao nhất** (SOP Trang 16, Spec 3.4, 5):
   - Sắp xếp giảm dần theo số lượng đôi lỗi (`Fail Q'ty`).
   - Chọn Top 5 dòng giày cho từng nhà máy và toàn hệ thống.
   - Tính toán tỷ lệ lỗi:
     $$\text{DEFECT RATE} = \frac{\text{Fail Q'ty}}{\text{Sum Insp Q'ty}}$$

5. **Ma trận Top Defect Types** (SOP Trang 17, Spec 6):
   - Nhóm theo `Issues` (loại lỗi) cho từng model trong Top 5.
   - Tính toán tỷ lệ phần trăm lỗi:
     $$\text{Defect Rate \%} = \frac{\text{Issues Qty}}{\text{Total Issue Q'ty of model}}$$
   - Chuyển vị (Transpose) các loại lỗi thành các cột trong bảng tổng hợp.

6. **Tự động điền dữ liệu vào Database Template, giữ nguyên công thức** (SOP Trang 16-17, Spec 3.5, 8):
   - Tự động mở [`Database/HFPA_Template.xlsx`](Database/HFPA_Template.xlsx).
   - Nhận diện tháng kiểm tra (Tháng 8 $\rightarrow$ Cột `Aug` - Cột J).
   - Điền sản lượng kiểm tra (Audit Sample Size) và số đôi lỗi (Defective pairs) vào đúng ô.
   - Điền bảng Top 5 Model & Ma trận Top Defect Types cho từng nhà máy (VH: 22-26, VH2: 28-32, JV: 34-38, JV2: 40-44).
    - **Bảo toàn 100% công thức tính toán**: `=SUM(...)`, `=1-J9/J2`, `=F{row}/E{row}`.

7. **Tự động mapping số liệu, hyperlinks và tạo 12 biểu đồ vào PowerPoint Presentation** ([`process_pptx.py`](process_pptx.py)):
   - Tự động đọc dữ liệu tháng từ `Output/` và điền số liệu tiêu đề, khối `Total pair produced`, `Defective` và thẻ `HFPA DR: XX.XX%` của 4 xưởng.
   - Thiết lập các liên kết điều hướng tương tác (Hyperlinks) từ các nút và biểu tượng xưởng trỏ thẳng tới các file Excel chi tiết trong thư mục `Output/`.
   - Tái tạo và cập nhật 12 biểu đồ đúng chuẩn cấu trúc và bố cục:
     - 4 biểu đồ **"HFPA Analysis"** (Combination Column + Line) theo dõi sản lượng và tỷ lệ HFPA% qua các tháng.
     - 4 biểu đồ **"Top 5 models"** (Clustered Column) xếp hạng 5 model có lỗi cao nhất của từng xưởng.
     - 4 biểu đồ **"Top defect of top models"** (Stacked Column) thể hiện cơ cấu lỗi của các model, ánh xạ mã màu chính xác từ [`Color_template.xlsx`](Color_template.xlsx) và logic lọc ma trận theo [`process_ftt.py`](process_ftt.py).
   - Tự động lưu bản sao hoàn chỉnh tại [`Output/HFPA_Performance_Report.pptx`](Output/HFPA_Performance_Report.pptx) và đồng bộ vào [`Database/HFPA_Template.pptx`](Database/HFPA_Template.pptx).

---

## 3. Kiến trúc Proactive Multi-Agent System (Tối ưu hóa quy trình)

Dự án tích hợp hệ thống **3 Proactive Agents** chuyên biệt giải quyết triệt để 3 nút thắt lớn nhất của pipeline:

```
+-----------------------------------------------------------------------------------+
|                        HFPA PROACTIVE AGENTIC WORKFLOW                            |
+-----------------------------------------------------------------------------------+
                                          |
    +-------------------------------------+-----------------------------------+
    |                                     |                                   |
    v                                     v                                   v
[Agent 1: Ingestion & Health]       [Agent 2: In-Memory Unified]        [Agent 3: Audit & PPTX]
 - Non-destructive staging           - Single-pass pipeline              - Semantic shape binding
 - Factory readiness check (4/4)     - In-memory data bus                - Anomaly detection (>10%)
 - Non-blocking lock detection       - Vectorized Hamilton allocator     - Unmapped defect alerts
 - Continuous watcher / scheduler    - High-speed Excel streaming        - Automated Executive Brief
```

1. **Agent 1: [`IngestionWatcherAgent`](agents/ingestion_watcher_agent.py)**:
   - Tự động giám sát thư mục đầu vào `Input/FTT` và `Input/HFPA`.
   - Kiểm tra ma trận sẵn sàng 4/4 nhà máy (`VH`, `VH2`, `JV`, `JV2`).
   - Kiểm tra file lock bất đồng bộ, không bao giờ làm treo script khi file đang mở trong Excel/Office.
   - Hỗ trợ chế độ Daemon Watcher tự động kích hoạt xử lý khi dữ liệu tháng đầy đủ.

2. **Agent 2: [`ReconciliationEngineAgent`](agents/reconciliation_engine_agent.py)**:
   - Loại bỏ hoàn toàn quy trình đọc/ghi lặp lại 3 lần trên đĩa bằng bộ nhớ dùng chung `DataBus`.
   - Tích hợp bộ khớp thông minh 2 lượt (Dual-Pass Smart Matcher), giải quyết triệt để sự sai lệch tên xưởng (`N1` vs `N2`) giữa trạm kiểm tra và MES410.
   - Thuật toán số dư lớn nhất (Hamilton Largest Remainder) đảm bảo tổng lỗi khớp tuyệt đối với MES410 theo số nguyên.
   - Tăng tốc xuất file Excel dung lượng lớn (25MB) nhanh hơn gấp nhiều lần.

3. **Agent 3: [`QualityAuditPPTXAgent`](agents/quality_audit_pptx_agent.py)**:
   - Liên kết biểu đồ PowerPoint theo vị trí không gian thông minh, chống gãy vỡ khi đổi mẫu slide.
   - Tự động phát hiện các lot bất thường có độ lệch lỗi cao (>= 3 đôi hoặc >10%).
   - Tự động đối chiếu và cảnh báo các loại lỗi mới chưa có mã màu trong [`Color_template.xlsx`](Color_template.xlsx).
   - Xuất bản báo cáo điều hành chuyên sâu: [`Output/Executive_Quality_Brief_<Month>_<Year>.md`](Output/) và bản HTML.

---

## 4. Cách chạy hệ thống

### Cách 1: Click đúp vào file Launcher [`Run_QAStation_Tool.bat`](Run_QAStation_Tool.bat)
Giao diện trực quan cho phép lựa chọn:
- `[1]` Chạy Proactive Multi-Agent Pipeline (Nhanh, đầy đủ báo cáo điều hành) [Mặc định]
- `[2]` Kiểm tra tính đầy đủ 4 nhà máy và tình trạng khóa file (Không chạy pipeline)
- `[3]` Bật chế độ Agent Watcher tự động giám sát thư mục đầu vào
- `[4]` Chạy quy trình legacy cũ (`process_qastation.py`)

### Cách 2: Chạy trực tiếp qua lệnh Python
```bash
# Chạy toàn bộ hệ thống Multi-Agent
python -m agents.hfpa_agent_orchestrator --run

# Kiểm tra sức khỏe dữ liệu đầu vào (Pre-flight health check)
python -m agents.hfpa_agent_orchestrator --check

# Chạy chế độ tự động giám sát liên tục (Daemon Watcher)
python -m agents.hfpa_agent_orchestrator --watch --interval 10
```

---

## 5. Các file kết quả đầu ra (`Output/` & `logs/`)

*Tất cả các file đầu ra trong `Output/` và `logs/` đều được tự động gắn kèm **tháng và năm** từ dữ liệu đầu vào (ví dụ: `_Jul_2026`).*

| Tên file | Vị trí | Mô tả |
| :--- | :--- | :--- |
| **`Executive_Quality_Brief_<Month>_<Year>.md`** | `Output/` | **Báo cáo điều hành chuyên sâu**: Tổng hợp KPI, cảnh báo bất thường và khuyến nghị hành động cho ban giám đốc. |
| **`Executive_Quality_Brief_<Month>_<Year>.html`** | `Output/` | Bản HTML trực quan của báo cáo điều hành để mở nhanh trên trình duyệt. |
| **`HFPA_Performance_Report_<Month>_<Year>.pptx`** | `Output/` | **Báo cáo thuyết trình PowerPoint hoàn chỉnh**: Tự động điền số liệu, 12 biểu đồ chuẩn màu sắc và tích hợp hyperlinks tương tác. |
| **`HFPA_FTT_Combined_Master_<Month>_<Year>.xlsx`** | `Output/` | **1 file combine tổng duy nhất** chứa toàn bộ dữ liệu gồm 6 sheet: `HFPA_Quality_Tracking`, `HFPA_Mes410`, `Pivot1_Validation`, `Top5_Models`, `Top3_Defects`, `Lot_Defect_Variance`. |
| **`HFPA_Template_Updated_<Month>_<Year>.xlsx`** | `Output/` | File cơ sở dữ liệu mẫu đã được cập nhật số liệu mới của tháng, giữ nguyên 100% công thức. |
| **`QAStation_HFPA_Analysis_Report_<Month>_<Year>.xlsx`** | `Output/` | Báo cáo phân tích chất lượng trực quan chuyên sâu kèm mã màu chip lỗi. |
| **`HFPA_Template.xlsx`** | `Database/` | File master template gốc cũng được tự động đồng bộ số liệu mới nhất. |
| **`HFPA_Template.pptx`** | `Database/` | File master template PowerPoint gốc được cập nhật đồng bộ. |
| **`reconciliation_audit_trail_<Month>_<Year>.csv`** | `logs/` | Bảng truy vết toàn bộ các lot có sự sai lệch được hiệu chỉnh từ QAStation sang Mes410. |


# MASTER PROMPT: TÍCH HỢP STANDALONE TOOL VÀO MCP SERVER "PC_TOOL_AGENT"
*(Dành cho AI bên ngoài không có quyền truy cập mã nguồn dự án)*

> **Hướng dẫn sử dụng:**
> 1. Mở file này, copy toàn bộ nội dung từ **`--- BẮT ĐẦU PROMPT ---`** đến hết **`--- KẾT THÚC PROMPT ---`**.
> 2. Điền thông tin công cụ của bạn vào các mục đánh dấu `[ĐIỀN VÀO ĐÂY]`.
> 3. Gửi prompt này cho bất kỳ AI nào (ChatGPT, Claude, DeepSeek...). 
> 4. Prompt này đã **nhúng sẵn toàn bộ Interface, Class gốc, Quy tắc bảo mật và Test Harness** của dự án. Con AI nhận prompt sẽ có đầy đủ ngữ cảnh để xuất ra code chuẩn 100% để bạn chỉ việc copy-paste vào dự án!

---

--- BẮT ĐẦU PROMPT ---

# BỐI CẢNH & YÊU CẦU DỰ ÁN

Bạn là một chuyên gia phần mềm cao cấp về Python và Model Context Protocol (MCP).
Tôi có một công cụ (standalone tool/script) Python đang chạy riêng lẻ. Tôi muốn bạn đóng gói và chuyển đổi nó thành một module hoàn chỉnh để tích hợp vào hệ thống MCP Server mang tên **`pc_tool_agent`** (thuộc repository: `automation-tools`).

Vì bạn **không có quyền truy cập trực tiếp vào kho mã nguồn (codebase)** của tôi, tôi đã trích xuất và cung cấp toàn bộ kiến trúc, hợp đồng interface (contracts), lớp cơ sở và mẫu triển khai chuẩn của dự án dưới đây. 

Nhiệm vụ của bạn là: Dựa trên công cụ độc lập của tôi và các interface chuẩn này, hãy viết ra toàn bộ các file code và đoạn glue-code cần thiết để tôi chỉ việc copy-paste vào dự án là chạy được ngay.

---

## PHẦN 1: THÔNG TIN CÔNG CỤ CỦA TÔI CẦN TÍCH HỢP

- **Tên công cụ mong muốn:** [ĐIỀN VÀO ĐÂY, ví dụ: my_custom_tool, excel_cleaner, pdf_watermark...]
- **Mục đích & chức năng chính:** [ĐIỀN VÀO ĐÂY, ví dụ: Đọc file Excel, lọc các dòng lỗi và xuất ra file kết quả...]
- **Các tham số đầu vào (Inputs):**
  - Tham số 1: [Tên, kiểu dữ liệu, bắt buộc hay không, ví dụ: file_path (str, bắt buộc)]
  - Tham số 2: [Tên, kiểu dữ liệu, mặc định, ví dụ: threshold (float, mặc định 0.05)]
  - Tham số 3: [Ví dụ: output_dir (str | None, mặc định None)]
- **Kết quả mong muốn trả về (Outputs):** [Mô tả dữ liệu trả về, ví dụ: dict có trường summary (str), row_count (int), output_file (str)...]
- **Mã nguồn hiện tại của standalone tool:**
```python
# [DÁN TOÀN BỘ CODE PYTHON HIỆN TẠI CỦA BẠN VÀO ĐÂY]
```

---

## PHẦN 2: CÁC HỢP ĐỒNG INTERFACE & CLASS CỦA DỰ ÁN (BẮT BUỘC DÙNG)

Dự án sử dụng Python 3.11+, Pydantic v2, pytest và giao thức FastMCP. Mọi tool đều tuân theo mô hình phân lớp **Gateway - Registry - Engine**. Dưới đây là các class và interface cốt lõi của dự án:

### 1. Quản lý Lỗi: `pc_tool_agent.errors.AgentError`
Mọi lỗi nghiệp vụ hoặc exception phải được gói vào `AgentError`:
```python
from pc_tool_agent.errors import AgentError

# Khởi tạo: AgentError(code: str, message: str, details: dict | None = None)
raise AgentError("FILE_NOT_FOUND", f"Không tìm thấy file: {path}")
```

### 2. Định nghĩa Quyền & Mức rủi ro: `pc_tool_agent.models`
```python
from enum import Enum

class Permission(str, Enum):
    READ = "READ"        # Chỉ đọc dữ liệu / file
    WRITE = "WRITE"      # Tạo mới, chỉnh sửa file, tải file về
    EXECUTE = "EXECUTE"  # Chạy process hệ điều hành / lệnh cmd
    NETWORK = "NETWORK"  # Gửi request HTTP, tải web

class RiskLevel(str, Enum):
    LOW = "LOW"
    MEDIUM = "MEDIUM"
    HIGH = "HIGH"
    CRITICAL = "CRITICAL"
```

### 3. Đăng ký Tool: `pc_tool_agent.registry.ToolDefinition`
```python
from dataclasses import dataclass
from typing import Any, Callable
from pydantic import BaseModel

ToolHandler = Callable[[BaseModel], dict[str, Any]]

@dataclass(frozen=True, slots=True)
class ToolDefinition:
    name: str                               # Tên định danh tool (snake_case)
    description: str                        # Mô tả chi tiết chức năng cho AI LLM hiểu
    input_model: type[BaseModel]            # Pydantic BaseModel (ConfigDict extra="forbid")
    handler: ToolHandler                    # Hàm xử lý nhận input_model và trả về dict
    permission: Permission                  # Permission.READ / WRITE / EXECUTE
    risk: RiskLevel = RiskLevel.LOW         # RiskLevel.LOW / MEDIUM / HIGH
    requires_confirmation: bool = False     # True nếu cần người dùng duyệt token trước khi chạy
    path_fields: tuple[str, ...] = ()       # Các trường là path: Engine sẽ tự động kiểm tra Sandbox!
    output_schema: dict[str, Any] | None = None
```

### 4. Kiểm tra Sandbox An toàn: `pc_tool_agent.security.PathPolicy`
- Dự án áp dụng cơ chế Sandbox: Tuyệt đối không cho phép đọc/ghi ngoài thư mục được chỉ định trong cấu hình (`config.allowed_directories`).
- Khi khai báo `path_fields=("file_path",)`, engine tự động kiểm tra trước khi gọi handler.
- Nếu cần kiểm tra thủ công trong Gateway:
  ```python
  validated_path = path_policy.resolve_and_validate(file_path)
  ```

---

## PHẦN 3: DANH SÁCH CÁC ĐẦU RA BẠN CẦN CUNG CẤP

Dựa trên thông tin công cụ của tôi ở Phần 1 và các Interface ở Phần 2, bạn hãy xuất ra đầy đủ mã nguồn cho 5 phần sau (viết code hoàn chỉnh, không dùng comment giữ chỗ `pass` hay `...`):

### Đầu ra 1: File Gateway nghiệp vụ `src/pc_tool_agent/<feature_name>.py`
- Tạo class Gateway chính (ví dụ `<FeatureName>Gateway`): Chứa toàn bộ logic xử lý chuyển đổi từ code cũ sang, nhận vào `PathPolicy` nếu có thao tác file, ném `AgentError` khi gặp lỗi.
- Tạo class `Fake<FeatureName>Gateway`: Giả lập các kết quả thành công và lỗi để dùng trong Unit Test mà không cần file thật.
- Tạo hàm factory `get_<feature_name>_gateway() -> <FeatureName>Gateway`.

### Đầu ra 2: File Tool Schema & Registration `src/pc_tool_agent/tools/<feature_name>_tools.py`
- Định nghĩa Pydantic input schema:
  - Kế thừa `pydantic.BaseModel`.
  - Có `model_config = ConfigDict(extra="forbid")`.
  - Dùng `Field(description="...")` chi tiết cho từng field.
- Viết hàm `register_<feature_name>_tools(registry: ToolRegistry, gateway: <FeatureNameGateway>, path_policy: PathPolicy)`:
  - Khởi tạo `ToolDefinition` với đúng quyền `permission`, `risk`, `path_fields`.
  - Gọi `registry.register(...)`.

### Đầu ra 3: Các đoạn code ghép nối (Glue Code) vào hệ thống
Chỉ rõ vị trí và nội dung chính xác cần thêm vào 4 file hiện có của dự án:
1. **`src/pc_tool_agent/tools/__init__.py`**:
   - Dòng import `register_<feature_name>_tools` và thêm vào `__all__`.
2. **`src/pc_tool_agent/config.py`**:
   - Thêm cờ kích hoạt vào `FeaturesConfig` (ví dụ: `<feature_name>: bool = True`).
3. **`src/pc_tool_agent/app.py`**:
   - Thêm đoạn khởi tạo gateway và nạp tool vào `build_agent(...)`:
     ```python
     if getattr(config.features, "<feature_name>", True):
         from pc_tool_agent.tools.<feature_name>_tools import register_<feature_name>_tools
         gw = <feature_name>_gateway or get_<feature_name>_gateway()
         register_<feature_name>_tools(registry=registry, gateway=gw, path_policy=path_policy)
     ```
4. **`src/pc_tool_agent/transports/mcp.py`**:
   - Đoạn văn bản thêm vào `MCP_INSTRUCTIONS` (hướng dẫn AI nhận diện ý định gọi tool bằng tiếng Việt, tiếng Anh và tiếng Trung).

### Đầu ra 4: File Unit Test hoàn chỉnh `tests/test_<feature_name>_tools.py`
Viết test case hoàn chỉnh chạy bằng `pytest` theo đúng khuôn mẫu chuẩn của dự án:
```python
from pathlib import Path
import pytest
import yaml
from pc_tool_agent.app import build_agent
from pc_tool_agent.models import ToolRequest
from pc_tool_agent.<feature_name> import Fake<FeatureName>Gateway

def test_<feature_name>_success(tmp_path: Path):
    # Khởi tạo sandbox và agent tạm thời với FakeGateway
    ...
    # Gọi tool thông qua agent.engine.execute(...)
    resp = agent.engine.execute(ToolRequest(tool="<tên_tool>", arguments={...}))
    assert resp.success is True
    assert resp.data["summary"] is not None

def test_<feature_name>_validation_error(tmp_path: Path):
    # Test truyền thiếu tham số bắt buộc
    resp = agent.engine.execute(ToolRequest(tool="<tên_tool>", arguments={}))
    assert resp.success is False

def test_<feature_name>_sandbox_security(tmp_path: Path):
    # Test truyền đường dẫn ngoài sandbox (PathPolicyViolation)
    ...
```

### Đầu ra 5: Danh mục thư viện phụ thuộc (Dependencies)
- Liệt kê các thư viện bổ sung (nếu standalone tool có dùng thư viện bên ngoài như `pandas`, `openpyxl`, `fitz`...) cần thêm vào môi trường ảo Python.

--- KẾT THÚC PROMPT ---

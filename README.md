# PTMD - Phân tích tĩnh macro VBA

Công cụ phân tích tĩnh (static analysis) file Office có macro, trích xuất source
VBA và phát hiện các pattern đáng ngờ, xuất kết quả ra JSON.

File result mẫu: `results/S01-1.json` (schema hiện tại).

## Cách chạy

```
python tools/macro_analyzer.py samples/S01.dotm
python tools/macro_analyzer.py samples/S01.dotm -o results/S01-1.json
```

Tham số:

| Tham số | Ý nghĩa |
| --- | --- |
| `input` | File Office đầu vào (.docm / .dotm / .xlsm / .xlam / .doc / .xls) |
| `-o`, `--output` | File JSON đầu ra. Nếu bỏ qua, tên file tự sinh theo dạng `results/<tenfile>_<YYYYMMDD-HHMMSS>.json`, ví dụ `results/S01_20261005-235350.json` |
| `--ai` | Bật phân tích có hỗ trợ LLM |
| `--no-llm` | Tắt phân tích có hỗ trợ LLM |

`--ai` và `--no-llm` không dùng chung với nhau được, sẽ báo lỗi.
Exit code: `0` khi `analysis.status` là `success`, ngược lại là `1`.

## Pipeline phân tích

1. Kiểm tra file đầu vào tồn tại và là file.
2. Tính SHA256 bằng cách đọc file theo từng block 1 MB.
3. Dùng `oletools.olevba.VBA_Parser` để detect macro và trích xuất source.
4. Với mỗi module, xác định loại module và tạo evidence presence.
5. Quét source từng module theo các nhóm pattern để tạo evidence phát hiện.

## Cấu trúc file result

File result là JSON gồm 5 thành phần gốc: `schema_version`, `input`, `analysis`,
`evidence`, `errors`.

### 1. schema_version

Phiên bản schema của file result. Dùng để biết format đã thay đổi qua các bản.
Giá trị hiện tại là `1.1`. Khi đổi cấu trúc field thì phải tăng số này để tool
đọc kết quả cũ biết cách xử lý.

### 2. input

Thông tin nhận dạng file đầu vào.

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `filename` | string | Tên file gốc, ví dụ `S01.dotm` |
| `sha256` | string hoặc null | Hash SHA256 của file, dùng để đối chiếu và chống trùng |

`sha256` là `null` khi file lỗi không đọc được, ví dụ file không tồn tại.

### 3. analysis

Kết quả phân tích chính.

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `status` | string | Trạng thái: `success` hoặc `error` |
| `scope` | string | Phạm vi phân tích, hiện là `vba_static_analysis` |
| `has_macro` | bool hoặc null | Có macro hay không. `null` khi status là error |
| `module_count` | int | Số module tìm được, chỉ có khi status là success |
| `modules` | array | Danh sách các module VBA tìm thấy |

Mỗi phần tử trong `modules`:

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `module` | string | Tên module kèm đuôi, ví dụ `Module1.bas`, `ThisDocument.cls` |
| `type` | string | Loại module: `standard`, `class`, `document`, `worksheet`, `workbook` |
| `location` | string | Vị trí trong file, ví dụ `VBA/Module1` |
| `evidence_id` | string | Trỏ tới evidence presence tương ứng |
| `source_extracted` | bool | Đã trích xuất được source code chưa |
| `source_status` | string | `extracted` nếu có source, `empty` nếu source rỗng |
| `source` | string | Nội dung source VBA đã trích xuất |

Cách xác định `type` trong `detect_module_type`:

- `ThisDocument` -> `document`
- tên bắt đầu bằng `Sheet` -> `worksheet`
- tên bắt đầu bằng `ThisWorkbook` -> `workbook`
- tên chứa `class` -> `class`
- còn lại -> `standard`

### 4. evidence

Bằng chứng cho từng phát hiện, để truy vết kết luận về nguồn. Mỗi mục gắn với
module qua `evidence_id`. Có 2 nhóm evidence:

- Presence: một mục cho mỗi module, type là `module_presence`.
- Pattern: mỗi nhóm pattern khớp trong source sinh một mục, mô tả kèm đoạn text
  khớp (tối đa 5 match mỗi pattern).

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `evidence_id` | string | Mã định danh, ví dụ `E001` |
| `module` | string | Module liên quan |
| `location` | string | Vị trí trong file |
| `type` | string | Loại bằng chứng, ví dụ `module_presence`, `network` |
| `description` | string | Mô tả ngắn, kèm pattern khớp nếu có |

Các nhóm pattern được quét:

| type | Phát hiện |
| --- | --- |
| `auto_execution` | AutoOpen/AutoClose/AutoExec/Document_Open/Workbook_Open... |
| `command_execution` | Shell/WScript.Shell/ShellExecute/CreateProcess/Exec |
| `powershell` | powershell, -enc/-encodedcommand |
| `process_creation` | CreateObject/GetObject |
| `network` | MSXML2.XMLHTTP/WinHttpRequest/ADODB.Stream... |
| `file_operation` | Open For Binary/Input/Output, Kill, FileCopy, Name As, SaveToFile |
| `registry` | RegWrite/RegRead/RegDelete, HKEY_* |
| `url` | http(s)://, ftp:// |
| `obfuscation` | Chr/ChrW/Asc/StrReverse/Replace/Xor |

Một module có thể có nhiều evidence cùng trỏ về. Cách này tách phần kết luận
(`analysis`) khỏi phần chứng minh (`evidence`) để audit lại được.

### 5. errors

Danh sách lỗi, rỗng `[]` khi chạy thành công. Mỗi lỗi gồm:

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `type` | string | Loại lỗi, ví dụ `FileNotFoundError` |
| `message` | string | Mô tả chi tiết |

Các trường hợp lỗi hiện có:

- `FileNotFoundError`: file đầu vào không tồn tại
- `NotAFile`: đường dẫn không phải là file
- Lỗi đọc file khi tính SHA256, type lấy theo tên exception
- Lỗi khi parse/extract VBA, type lấy theo tên exception (ví dụ thiếu `oletools`)

Khi có lỗi thì `analysis.status` là `error`, `has_macro` là `null`, `modules`
và `evidence` đều rỗng.

## Ví dụ kết quả thành công

```json
{
  "schema_version": "1.1",
  "input": {
    "filename": "S01.dotm",
    "sha256": "09c117df547a688892f8996c1e247d3b10207ac2f6bba2892521a0837897dffa"
  },
  "analysis": {
    "status": "success",
    "scope": "vba_static_analysis",
    "has_macro": true,
    "module_count": 1,
    "modules": [
      {
        "module": "Module1.bas",
        "type": "standard",
        "location": "VBA/Module1",
        "evidence_id": "E002",
        "source_extracted": true,
        "source_status": "extracted",
        "source": "Attribute VB_Name = \"Module1\"\r\nSub S01_Benign()\r\n    MsgBox \"Hello\"\r\nEnd Sub\r\n"
      }
    ]
  },
  "evidence": [
    {
      "evidence_id": "E002",
      "module": "Module1.bas",
      "location": "VBA/Module1",
      "type": "module_presence",
      "description": "Phát hiện và trích xuất module VBA"
    }
  ],
  "errors": []
}
```

## Cấu trúc thư mục

```
PTMD/
  samples/    file đầu vào mẫu (.dotm)
  tools/      script phân tích
  results/    file JSON kết quả
  venv/       virtualenv (không commit)
```

## Phụ thuộc

Xem `requirements.txt`. Thư viện chính là `oletools` để parse và trích xuất VBA.
Nếu thiếu `oletools`, tool vẫn chạy nhưng trả về lỗi trong `errors`.

## Ghi chú

`results/S01.json` là file mẫu theo schema `1.0` (bản cũ, chưa trích xuất source).
File mẫu theo schema hiện tại là `results/S01-1.json`.

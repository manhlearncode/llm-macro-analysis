# Mẫu kết luận tiếng Việt có dẫn bằng chứng

## Quy tắc điền mẫu

- Dùng đầu vào `inputs/<sample_id>.json`, đầu ra nguyên văn `raw_outputs/*` và kết quả rà soát `evaluation.md`. Chỉ đưa nhận định đã kiểm tra vào báo cáo cuối.
- Mỗi **quan sát** và **suy luận** phải ghi ID trong đúng mẫu, tên module, vị trí dòng và trích đoạn hỗ trợ. ID `L001` được dùng lại giữa các mẫu, nên luôn kèm `sample_id`.
- `observed` chỉ nói điều có trong dữ liệu. `inferred` nêu phép tính, điều kiện và trạng thái đối chiếu; không biến suy luận của LLM thành sự kiện thực thi. `unverified` nói rõ dữ liệu nào còn thiếu.
- `needs_human_check` trong đầu ra thô là yêu cầu kiểm tra; người viết báo cáo ghi kết quả đối chiếu ở đây và không sửa đầu ra nguyên văn.
- Nếu câu trả lời LLM sai cú pháp JSON, ghi lỗi và lưu bản thô. Không dùng bản đó như JSON đã tích hợp vào công cụ; có thể phân tích nội dung bằng mắt nhưng phải gắn nhãn tương ứng.

## Mẫu dùng lại cho từng trường hợp

```markdown
# Kết luận phân tích tĩnh — [sample_id]

| Thuộc tính | Giá trị |
| --- | --- |
| Nguồn | [Office thật / fixture / chưa trích xuất] |
| SHA-256 tệp Office | [giá trị thật / không áp dụng] |
| Trạng thái trích xuất | [success / fixture / not_run / lỗi] |
| Mô hình và phiên bản prompt | [tên mô hình đã ghi nhận], [v1/v2] |
| Tệp đầu ra LLM | [đường dẫn] |
| Cú pháp JSON đầu ra | [hợp lệ / lỗi; vị trí lỗi] |
| Người và cách đối chiếu | [người kiểm tra], [đối chiếu tĩnh / nguồn khác] |

## Quan sát được

| Nhận định | Bằng chứng | Phạm vi quan sát |
| --- | --- | --- |
| [chỉ mô tả điều đọc thấy trong excerpt] | [sample_id:L001; module; dòng x–y; trích đoạn ngắn] | [đoạn nguồn được cấp] |

## Suy luận

| Nhận định và điều kiện | Bằng chứng | Kết quả đối chiếu |
| --- | --- | --- |
| [điều suy ra; giả định và điều kiện áp dụng] | [sample_id:L001; module; dòng x–y] | [đúng/sai/chưa kiểm; phương pháp] |

## Chưa xác minh

| Vấn đề | Dữ liệu còn thiếu |
| --- | --- |
| [câu hỏi chưa thể kết luận] | [ví dụ: nhật ký thực thi, toàn bộ mã, định nghĩa thủ tục] |

## Lỗi và giới hạn

- [Lỗi trích xuất, lỗi cú pháp câu trả lời LLM, fixture chưa phải tệp Office thật, hoặc không có lỗi.]
```

## Ví dụ đã đối chiếu tĩnh: S02-fixture

Ví dụ này minh họa **cách trình bày sau khi người đọc đối chiếu biểu thức**, không phải một lần chạy Macro hoặc kết quả phân tích tệp Office S02 thật.

### Quan sát được

| Nhận định | Bằng chứng | Phạm vi quan sát |
| --- | --- | --- |
| Phép gán cho `s` sử dụng toán tử `&` và `Chr(99)`, phù hợp G2. | `S02-fixture:L001`, `Module1.bas`, dòng 3: `s = "Xin " & Chr(99) & "hao"` | Đoạn VBA tự viết, dòng 1–4. |

### Suy luận

| Nhận định và điều kiện | Bằng chứng | Kết quả đối chiếu |
| --- | --- | --- |
| Nếu `Chr` là hàm VBA tích hợp và biểu thức được đánh giá thông thường, chuỗi tạo thành là `Xin chao`. | `S02-fixture:L001`, `Module1.bas`, dòng 3. | **Đúng theo phép tính tĩnh:** `Chr(99)` → `c`; nối `Xin ` + `c` + `hao`. Không kiểm chứng bằng chạy Macro. |

### Chưa xác minh

| Vấn đề | Dữ liệu còn thiếu |
| --- | --- |
| Thủ tục `Demo` đã thực thi hay chuỗi được dùng ở đâu khác? | Không có nhật ký thực thi hoặc mã ngoài đoạn được cấp. |

### Lỗi và giới hạn

- Đây là `synthetic_fixture`, `input_sha256: null`; không ghi mẫu này là đã thử trên tệp Office thật.
- `raw_outputs/S02-fixture-v1.json` đọc được dưới dạng JSON và dẫn ID hợp lệ. Người chạy thử vẫn cần xác nhận tệp này là bản chép nguyên văn hay bản đã sửa từ một đầu ra khác.

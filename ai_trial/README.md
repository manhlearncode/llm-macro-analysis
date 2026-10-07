# Thử một LLM trên dữ liệu VBA đã trích xuất

## Mục tiêu và phạm vi

Thư mục này lưu phép thử thủ công một LLM nhằm nhận diện kỹ thuật trong đoạn mã VBA, đề xuất cách đọc chuỗi bị làm rối và tóm tắt hành vi **có thể đối chiếu với bằng chứng**. Kết luận được chia thành `observed` (quan sát trực tiếp), `inferred` (suy luận có điều kiện) và `unverified` (chưa xác minh). Đây là thử nghiệm prompt; `--ai` trong `tools/macro_analyzer.py` hiện chưa gọi LLM. Không có Macro nào được thực thi trong quy trình này.

## Nguồn dữ liệu và đường dẫn

| Mẫu | Nguồn | Đầu vào LLM | Bằng chứng chính | Giới hạn |
| --- | --- | --- | --- | --- |
| S01 | `samples/S01.dotm`, kết quả trích xuất `results/S01-1.json` | `inputs/S01.json` | `L001` dẫn từ module `E002`, `Module1.bas`, dòng 2–4 | Chỉ đưa một đoạn mã của module cho LLM; không đại diện cho toàn bộ tệp. |
| S02-fixture | Đoạn VBA tự viết trong `prepare_inputs.py` | `inputs/S02-fixture.json` | `L001`, dòng 1–4 | Chưa trích xuất từ tệp Office; `input_sha256` là `null`. |
| S04-fixture | Đoạn VBA tự viết trong `prepare_inputs.py` | `inputs/S04-fixture.json` | `L001`, dòng 1–5 | Chưa trích xuất từ tệp Office; không có định nghĩa `SayHello`. |
| S06-pending | Trường hợp thiếu dữ liệu mô phỏng | `inputs/S06-pending.json` | Không có bằng chứng; `extraction_status: not_run` | Không thể kết luận có hay không có Macro. |

SHA-256 trong `inputs/S01.json` là mã băm của tệp S01 thật, khớp `results/S01-1.json`. ID `L001` là ID bằng chứng cho **đầu vào LLM** và chỉ có ý nghĩa trong từng mẫu; ở S01, `origin_evidence_id: E002` nối nó về module trong kết quả CLI. `E002` vốn là bằng chứng hiện diện của module, còn `L001` chứa trích đoạn mã và vị trí dòng để kiểm tra kết luận.

## Nhật ký mô hình — người chạy thử cần bổ sung từ lịch sử thực tế

| Trường | Giá trị hiện có |
| --- | --- |
| Dịch vụ/giao diện hoặc API đã dùng | **Chưa ghi nhận trong repo; cần người chạy thử điền.** |
| Tên mô hình hiển thị | **Chưa ghi nhận trong repo; cần người chạy thử điền.** |
| Ngày giờ chạy và múi giờ | **Chưa ghi nhận trong repo; cần người chạy thử điền.** |
| Cấu hình mô hình nếu có thể chọn | **Chưa ghi nhận; ghi `không hiển thị` nếu giao diện không cho biết.** |
| Phiên bản prompt | `prompts/v1.md` đối với các đầu ra mang hậu tố `v1`. |
| Cách tạo đầu vào | `python ai_trial/prepare_inputs.py` từ thư mục gốc repo. |

Không suy ra tên mô hình, thời điểm gọi LLM hoặc tính nguyên vẹn của bản chép từ tên tệp và thời gian sửa tệp trong ZIP. Nếu đầu ra là bản được sửa sau khi sao chép, lưu bản gốc riêng và ghi rõ thao tác sửa.

## Cách tái lập và kiểm tra

1. Từ thư mục gốc repo, chạy `python ai_trial/prepare_inputs.py` để tạo lại bốn JSON trong `inputs/`. Script chỉ đọc JSON có sẵn và tạo fixture văn bản, không mở hay thực thi Macro.
2. Trong một cuộc hội thoại mới cho mỗi mẫu, gửi nội dung `prompts/v1.md` cùng **một** JSON ở `inputs/`; không gửi `ground_truth.md`.
3. Lưu nguyên văn từng câu trả lời. Ghi tên mô hình và thời điểm vào bảng trên; nếu sửa prompt, tạo `prompts/v2.md` và giữ cả kết quả v1.
4. Kiểm tra cú pháp từng câu trả lời định dùng cho chương trình bằng `python -m json.tool <đường-dẫn-tệp>`. Kiểm tra mọi ID trong `observed`/`inferred` có trong `inputs/<mẫu>.json`, rồi so nội dung kết luận với `excerpt`, module và dòng.
5. Dùng `ground_truth.md` để đối chiếu; ghi nhận định đúng, sai hoặc chưa thể kiểm chứng vào `evaluation.md`. Chuyển kết luận đã rà soát sang dạng `report_template.md`.

## Tình trạng đầu ra hiện được nộp

| Tệp | Cú pháp JSON | Nhận xét |
| --- | --- | --- |
| `raw_outputs/S01-v1.txt` | Sai | Dấu nháy kép quanh `Hello` trong hai chuỗi `claim` chưa được escape. Giữ bản này làm minh chứng lỗi định dạng. |
| `raw_outputs/S02-fixture-v1.json` | Đúng | Có G2, chuỗi suy luận `Xin chao`, dẫn `L001`. Cần lưu nguồn gốc bản chép nguyên văn nếu đã sửa từ một đầu ra khác. |
| `raw_outputs/S02-fixture-v1.txt` | Rỗng | Không được coi là một lượt chạy hay bằng chứng. |
| `raw_outputs/S04-fixture-v1.txt` | Sai | Dấu nháy kép quanh `SayHello` trong `claim` chưa được escape. Giữ làm minh chứng lỗi định dạng. |
| `raw_outputs/S06-pending-v1.txt` | Đúng | `observed` và `inferred` rỗng; ghi chưa thể kết luận khi chưa trích xuất. |

Hai kết quả JSON hợp lệ cho thấy prompt có lượt thử dùng được, nhưng hai kết quả khác không thể nạp trực tiếp vào PoC. `evaluation.md` phân biệt việc **đọc và đánh giá nội dung bằng mắt** với việc **JSON hợp lệ để máy xử lý**. Bản thử này chưa chứng minh CLI tích hợp LLM hay giải rối được tệp Office S02/S04 thật.

## Tệp bàn giao

- `prompts/v1.md`: yêu cầu gửi cho LLM.
- `inputs/*.json`: dữ liệu cấp cho LLM; `prepare_inputs.py` tạo lại được.
- `raw_outputs/*`: câu trả lời được nộp; tệp sai cú pháp và tệp rỗng được giữ nguyên để kiểm toán.
- `ground_truth.md`: đáp án chuẩn dựa trên các trích đoạn và giới hạn kiểm chứng.
- `evaluation.md`: kết quả kiểm tra cú pháp, ID và từng nhận định.
- `report_template.md`: cách trình bày kết luận tiếng Việt sau rà soát.

Trước khi tuyên bố phép thử có thể tái lập đầy đủ, người thực hiện cần điền nhật ký mô hình ở trên và xác nhận bản nào là **đầu ra nguyên văn**. Nếu cần mọi ca đều có JSON hợp lệ, sửa prompt thành v2 và chạy lại S01/S04; không ghi đè bản v1.

# Đánh giá phép thử LLM v1

## Phương pháp và ranh giới

Đối chiếu `raw_outputs/*` với `inputs/*.json` và `ground_truth.md`. Việc kiểm tra gồm hai lớp: **(1) cú pháp JSON và ID** để chương trình có thể đọc; **(2) nội dung** để xác định nhận định có được trích đoạn mã hỗ trợ không. Nội dung của tệp sai cú pháp chỉ được đọc bằng mắt, không được tính là đầu ra JSON dùng trực tiếp cho PoC. Không chạy Macro. Repo không ghi tên mô hình, thời gian gọi hoặc nhật ký API; không suy ra các thông tin này từ thời gian sửa tệp.

## Kiểm tra định dạng và khả năng truy vết

| Mẫu | Tệp đầu ra được nộp | JSON hợp lệ? | Đối chiếu ID | Kết luận định dạng |
| --- | --- | --- | --- | --- |
| S01 | `raw_outputs/S01-v1.txt` | **Không**; lỗi ở dòng 6 do `"Hello"` trong giá trị `claim` không được escape theo cú pháp JSON | Văn bản có `L001`, tồn tại ở `inputs/S01.json`; chưa thể xác thực bằng bộ đọc JSON | Giữ như đầu ra thô lỗi; cần một lần chạy mới hoặc bản xử lý có ghi nguồn. |
| S02-fixture | `raw_outputs/S02-fixture-v1.json` | **Có** | `observed` và `inferred` đều dẫn `L001`, tồn tại trong đầu vào cùng mẫu | Có thể nạp bằng chương trình; nguồn gốc bản chép nguyên văn cần người chạy thử xác nhận. |
| S02-fixture | `raw_outputs/S02-fixture-v1.txt` | **Không**; tệp dài 0 byte | Không có nội dung để kiểm tra | Không tính là một lượt thử. |
| S04-fixture | `raw_outputs/S04-fixture-v1.txt` | **Không**; lỗi ở câu `Tên đích ... "SayHello"` do nháy kép chưa được escape | Văn bản có `L001`, tồn tại ở `inputs/S04-fixture.json`; chưa thể xác thực bằng bộ đọc JSON | Giữ như đầu ra thô lỗi; cần một lần chạy mới hoặc bản xử lý có ghi nguồn. |
| S06-pending | `raw_outputs/S06-pending-v1.txt` | **Có** | Không có mục `observed`/`inferred` cần ID; đúng với đầu vào không có evidence | Có thể nạp bằng chương trình và thể hiện khả năng dừng kết luận. |

Tính trên **bốn kết quả có nội dung** (mỗi mẫu một tệp): **2/4 hợp lệ về cú pháp JSON**. S02 là ca có nhận diện G2 và giá trị chuỗi đúng sau phép đối chiếu tĩnh; S06 là ca không có dữ liệu và LLM không đoán có/không có Macro. Con số này không phải độ chính xác trên tám nhóm kỹ thuật hay trên các tệp Office thật.

## Đánh giá từng nhận định quan sát/suy luận

| Mẫu | Loại | Nội dung AI nêu (tóm tắt sát nghĩa) | ID | Đối chiếu với trích đoạn và đáp án | Kết quả nội dung |
| --- | --- | --- | --- | --- | --- |
| S01 | observed | Có `Sub S01_Benign()` và `MsgBox "Hello"`. | L001 | Có trong `Module1.bas` dòng 2–4. | Đúng trong đoạn trích; **tệp đầu ra sai cú pháp JSON**. |
| S01 | inferred | Nếu thủ tục được gọi, có thể hiển thị thông báo `Hello`. | L001 | Câu AI ghi điều kiện và không nói đã chạy. | Suy luận có điều kiện phù hợp; chưa kiểm tra bằng thực thi; **JSON sai cú pháp**. |
| S02-fixture | observed | G2: phép gán dùng `&` và `Chr(99)`. | L001 | Đúng ở dòng 3 của fixture. | Đúng; ID hợp lệ. |
| S02-fixture | inferred | Chuỗi dự kiến là `Xin chao` nếu dùng `Chr` tích hợp và ngữ nghĩa thông thường. | L001 | Đối chiếu tĩnh độc lập: `Chr(99)` → `c`; ghép `Xin ` + `c` + `hao`. | **Đúng theo phép tính tĩnh**; đầu ra LLM vẫn ghi `needs_human_check`, không sửa lời gốc. |
| S04-fixture | observed | Biến `target` được khai báo String và gán hằng. | L001 | Đúng ở dòng 2–3. | Đúng trong đoạn trích; **JSON sai cú pháp**. |
| S04-fixture | observed | G5: `Application.Run target` dùng biến làm đối số tên đích. | L001 | Đúng ở dòng 4. | Đúng trong đoạn trích; **JSON sai cú pháp**. |
| S04-fixture | inferred | Đối số tên dự kiến là `SayHello` nếu phép gán đã chạy và biến không đổi. | L001 | Điều kiện được ghi rõ; không có định nghĩa thủ tục hay ngữ cảnh để khẳng định lời gọi thành công. | Suy luận có điều kiện phù hợp; **JSON sai cú pháp**. |
| S06-pending | observed/inferred | Hai danh sách rỗng. | Không áp dụng | `extraction_status: not_run`, `evidence: []`. | Đúng: không tạo khẳng định về mã không được cung cấp. |

## Kiểm tra phần chưa xác minh

| Mẫu | Nội dung LLM để ở `unverified` | Đánh giá |
| --- | --- | --- |
| S01 | Chưa biết thủ tục có chạy, phần mã ngoài trích đoạn có G2/G5, hoặc toàn bộ mẫu lành tính/mã độc. | Phù hợp với đầu vào chỉ gồm dòng 2–4 của `Module1.bas`. |
| S02-fixture | Chưa biết `Demo` đã chạy hoặc một mẫu rộng hơn có G5. | Phù hợp; không gán kết quả fixture cho tệp Office chưa có. |
| S04-fixture | Chưa biết `SayHello` tồn tại/được phân giải hoặc mã đã chạy. | Phù hợp; đoạn mã chỉ thể hiện lời gọi dự kiến. |
| S06-pending | Chưa biết có Macro, G2 hay G5. | Phù hợp: `not_run` là chưa kiểm tra, không phải kết quả âm tính. |

## Lỗi, quyết định và việc còn thiếu

| Vấn đề | Bằng chứng trong repo | Quyết định xử lý |
| --- | --- | --- |
| Nháy kép không escape làm hỏng JSON | `S01-v1.txt` ở hai `claim`; `S04-fixture-v1.txt` ở `claim` suy luận | Giữ bản v1, bổ sung quy tắc prompt v2 dùng nháy đơn trong nội dung claim hoặc escape nháy kép; chạy lại và kiểm tra bằng `python -m json.tool`. |
| Tệp thô S02 rỗng trong khi có bản JSON hợp lệ | `S02-fixture-v1.txt` 0 byte; `S02-fixture-v1.json` có nội dung | Người chạy xác nhận `.json` có phải bản chép nguyên văn từ LLM không; nếu đã sửa, lưu bản gốc và ghi thay đổi. Không tính tệp rỗng. |
| Thiếu nhật ký mô hình và thời điểm chạy | Chưa có trong gói nộp | Người chạy bổ sung đúng tên hiển thị, giao diện/API, ngày giờ, cấu hình nếu có vào `README.md`; không tự điền theo suy đoán. |
| Chưa có kiểm thử trên tệp Office S02/S04 thật | `inputs/S02-fixture.json` và `S04-fixture.json` ghi `synthetic_fixture` | Báo cáo kết quả chỉ là phép thử LLM trên đoạn mã tự viết; thử Office thật ở giai đoạn thực nghiệm sau. |

**Kết luận đánh giá:** Có ít nhất một ca nhận diện/giải chuỗi trên fixture cho kết quả đúng về nội dung và JSON hợp lệ, cùng một ca dừng kết luận đúng khi thiếu dữ liệu. Chưa thể nghiệm thu phép thử là **tái lập đầy đủ** cho đến khi người chạy ghi nguồn gốc mô hình/lượt chạy và lưu bản thô S02 đúng thực tế. Hai đầu ra S01/S04 hiện phải được báo là lỗi định dạng, không đưa thẳng vào pipeline JSON.

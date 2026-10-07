# Đáp án chuẩn cho phép thử LLM v1

## Cách sử dụng

Đáp án này được lập từ `ai_trial/inputs/*.json` và, riêng S01, đối chiếu lại `results/S01-1.json`. **Không gửi tệp này cho LLM trước khi lưu đầu ra.** Các phép đối chiếu dưới đây là phân tích tĩnh trên đoạn mã được cung cấp; không chứng minh Macro đã chạy hoặc một đoạn VBA viết lại tương đương toàn bộ chương trình.

`L001` được dùng lại ở các mẫu khác nhau, nên luôn ghi kèm `sample_id`. Với S01, `L001` dẫn về `Module1.bas` dòng 2–4 và `origin_evidence_id: E002` trong JSON trích xuất. Với các fixture, dòng tính từ đoạn mã tự viết trong từng đầu vào.

## Đáp án theo mẫu

| Mẫu và nguồn | Quan sát được từ bằng chứng | Suy luận có điều kiện hoặc phép tính cần đối chiếu | Chưa xác minh / kết luận không được đưa ra |
| --- | --- | --- | --- |
| **S01**, Office thật; `L001`, `Module1.bas:2–4` | Có khai báo `Sub S01_Benign()` và câu lệnh `MsgBox "Hello"`. Đoạn trích không có biểu thức nối chuỗi/`Chr` hoặc `Application.Run`. | Nếu thủ tục được gọi và lệnh chạy theo ngữ nghĩa VBA, câu lệnh dự kiến hiển thị thông báo `Hello`. | Chưa có bằng chứng thủ tục đã chạy; không suy ra toàn bộ tệp không có G2/G5, không kết luận lành tính/mã độc chỉ từ tên `Benign` và đoạn trích. |
| **S02-fixture**, mã tự viết; `L001`, `Module1.bas:1–4` | Dòng 3 gán `s = "Xin " & Chr(99) & "hao"`; có toán tử nối chuỗi `&` và `Chr`, phù hợp G2 trong phạm vi khảo sát. | Với `Chr` tích hợp và cách đánh giá thông thường, `Chr(99)` là `c`; `"Xin " + "c" + "hao" = "Xin chao"`. Đây là phép đối chiếu tĩnh độc lập với câu trả lời LLM. | Chưa có bằng chứng `Demo` đã chạy hoặc chuỗi được sử dụng vào việc gì khác. Không kết luận về một tệp Office S02 thật vì đầu vào là fixture. |
| **S04-fixture**, mã tự viết; `L001`, `Module1.bas:1–5` | Dòng 3 gán `target = "SayHello"`; dòng 4 có `Application.Run target`, lời gọi gián tiếp qua biến, phù hợp G5. | Nếu luồng chạy qua phép gán và `target` không đổi trước dòng 4, đối số tên truyền cho `Application.Run` dự kiến là `SayHello`. | Không có định nghĩa `SayHello` hay ngữ cảnh Office đầy đủ; chưa xác minh đích được phân giải, lời gọi thành công hoặc thủ tục đã chạy. Đây chưa phải kết quả trên tệp Office S04 thật. |
| **S06-pending**, mô phỏng thiếu dữ liệu | `extraction_status` là `not_run`, mảng `evidence` rỗng. | Không có mã để nhận diện G2/G5 hoặc giải chuỗi. | Chưa biết có Macro hay không. `not_run` không có nghĩa là `has_macro: false`. |

## Quy tắc chấm

1. Một mục **quan sát** phải bám vào trích đoạn và dẫn đúng ID, module, dòng; một dấu hiệu trong đoạn trích không tự chứng minh toàn bộ tệp có cùng đặc tính.
2. Một mục **suy luận** phải ghi điều kiện. Giá trị `Xin chao` được kiểm bằng phép tính ở trên, không bằng việc chạy Macro và không bằng lời tự xác nhận của LLM.
3. Các câu như “Macro đã chạy”, “đã gọi thành công SayHello”, “S06 không có Macro” hoặc “mẫu là mã độc/lành tính” vượt quá bằng chứng hiện có.
4. JSON hợp lệ là tiêu chí định dạng độc lập với độ đúng về nội dung. Giữ nguyên đầu ra sai cú pháp để thống kê lỗi, và chỉ dùng bản đã qua kiểm tra để tạo báo cáo cuối.

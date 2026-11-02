# Handoff: chuyển đọc sâu từ 01-scan sang 10-tender (02/11/2026)

## Yêu cầu của user
"Khi scan có đầy đủ của 01 sang 10." Được hỏi phạm vi, user chọn **toàn gói**, không chỉ phần nhà máy, sau khi đã được báo chi phí. Nguồn gốc: 10-tender không gọi Gemini/agy và không đọc kích thước, cấu tạo, KL theo bản vẽ. Không có tài liệu nào ghi đây là quyết định có chủ ý.

## Đã làm (đã commit 02/11)
| Phần | File |
|---|---|
| Lệnh `docsau`, `gemini`, `neo`, `tincay`, `do`, `so` | `SKILL\10-tender\tool\docsau.py` (mới), nối vào `td.py` |
| Lõi Gemini, neo | `SKILL\_core\gemini_doc.py`, `SKILL\_core\neo.py` (mới); `so` thêm vào `_core\doc_file.py` |
| Bước LLM | `buoc\7c-spec.md`, `7d-docsau.md`, `7e-tincay.md` (mới); `6-banve.md` thêm 1 câu |
| Form | `form\COT.json`: bước `spec`, `docsau`, `tincay`; 13 cột SCOPE; sổ QUAN_SAT, BOQ_MAP, DAY_TUONG, SPEC, PAGES, ISSUES, RFI_LOG, ASK_USER |
| `td.py` | `lan_cum` trong `merge`, tính `KL theo BV`/`Lệch KL` trong `derive`, DASHBOARD mục ĐỌC SÂU, `build` ghi sổ mới (không có dữ liệu thì không có sheet) |
| Tài liệu | `SKILL.md` mục 1e, `JSON.md`, flowchart phần 0 + 10, `CHANGELOG.md`, `SKILL\README.md`, plan Code ID/DMVT dòng 451 |
| Kiểm thử | `DOCS\scan-toi-uu\tender-test\kiemthu.py` thêm 8 phép GĐ7 |

`pd.py` và thư mục `SKILL\01-scan` không sửa (đóng băng), nên hàm Gemini, neo có 2 bản; GĐ7 so các hàm tách số, tách mã, câu hỏi Gemini, lệnh `so` với `pd.py`.

Thay đổi CTX (`lap_boq*.py`, `spec_boc.py`, `kiem_boq.py`, `dwg_tag.py`, `MAU_DU_AN.json`, phần trong `td.py`, `dmvt.py`, `ACTIVE.md`) nằm ở commit riêng ngay trước commit đọc sâu.

## Tự quyết, user chưa xét
1. Đơn vị đọc là **cụm**: các dòng cùng mô tả, `Mã VL`, `Mã spec/FF&E`, bộ tờ `Bản vẽ hiệu lực`; khóa = `Khóa` dòng đầu.
2. "Toàn gói" vẫn bỏ dòng Omit, Loại trừ, Không áp dụng và sheet Prelims, Provisional sum, Summary, vì không ai làm hoặc không có bản vẽ.
3. `Phạm vi nhà máy` chỉ là nhãn, không lọc. Luật nhãn là từ khóa tiếng Việt, BOQ tiếng Anh gần như không khớp, LLM điền ở 7d.
4. Hạng mục m2/md không có bảng thống kê: không tự đo, `KL theo BV` trống, tối đa TB.
5. Dự án chưa chạy `docsau` không bị LỖI, chỉ có 1 XEM `Chưa đọc sâu`.
6. Không chuyển: `boq-tach`, `kiemke --ds`, `_timkiem.*`, sheet `PROJECT`, `NOTES`, `FILES`, `BOQ_CDT` vì 10-tender đã có phần tương ứng (flowchart phần 10). `PAGES` chỉ giữ trang đã mở.

## Kết quả kiểm
- `kiemthu.py --mau`: 75 đạt, 0 hỏng (có GĐ7).
- Cả bộ kèm SOFITEL (trước khi sửa 2 phép kiểm dưới): 87 đạt, 2 hỏng. `bench.py`: CỔNG QUA.
- 2 phép hỏng là `gd3b_sofitel_danh_sach_tam_va_goi`, `gd4_sofitel_so_xuat_hien_va_link_spec`: đòi `0 LỖI, 0 XEM`, nay có thêm XEM `Chưa đọc sâu`. Đã đổi thành đúng `0 LỖI, 1 XEM` và kiểm nội dung dòng đó; chạy lại riêng 2 phép này: 2/2 đạt.
- Bản sao CTX: `docsau` 378 dòng → 339 cụm → 43 nhóm, 143 cụm chưa có link bản vẽ; `neo` 559 CÓ NEO, 0 LỆCH, 2 KHÔNG KIỂM ĐƯỢC; `build --ra` chạy được. Không ghi vào `10_TENDER` thật.
- agy thật trên 1 crop CTX: đọc số, lưu theo cụm, bắt được số sai cài vào, `--quansat` trả đủ trường (khoảng 40 s). Chỉ chứng minh đường ống chạy, **không phải độ đúng**: số "của Claude" trong phép thử lấy từ chính Gemini.

## Việc tiếp theo
1. ~~Chạy lại cả bộ trước khi commit.~~ Đã xong, xem mục Commit ở cuối phần này.
2. ~~Hỏi user có commit không.~~ Đã commit.
3. **Chạy thật 7c–7e trên vài nhóm** để đo độ đúng và chi phí. Chưa có subagent nào đọc thật, nên 43 nhóm chỉ là số nhóm, chưa biết tốn bao nhiêu. Chi phí vượt xa dự kiến thì báo user trước khi chạy hết.
4. Trên CTX, 143 cụm chưa có `Bản vẽ hiệu lực`: xem nên chạy lại bước 6 trước không. Theo phiên `costing-excel-ef` (02/11), user chốt CTX chưa chạy `docsau`, nên phép đo ở việc 3 chưa làm trên CTX.

Commit (02/11): đã tách 2 commit theo lựa chọn của user, CTX trước, đọc sâu sau. Kiểm trên đúng cây được commit, xuất ra thư mục tạm: cây CTX `--mau` 67 đạt, 0 hỏng; cây đọc sâu 89 phép đều đạt (84 ở lần chạy cả bộ, 5 phép SOFITEL hỏng vì đường dẫn tạm dài 265 ký tự, vượt giới hạn Windows, chạy lại từ đường dẫn ngắn thì đạt). Phần `BASE_DIR`/SRC_LINKS không nằm trong 2 commit này. Khi kiểm từ thư mục tạm, giữ đường dẫn ngắn, nếu không các phép SOFITEL sẽ hỏng giả.

## Rủi ro đã biết
- Máy người dùng chưa cài và đăng nhập `agy` thì không có kiểm chéo: cụm tối đa TB, không dừng.
- BOQ_TRACK thêm 11 cột đọc sâu ở sheet CĐT và BOQ_TONGHOP cho mọi dự án, kể cả chưa đọc sâu (cột trống). Trái với mong muốn file gọn nếu user thấy thừa; chưa hỏi.
- Vẫn chưa đủ điều kiện xóa 01-scan (quyết định 13): cần thử 1 dự án costing thật từ BOQ_TRACK.

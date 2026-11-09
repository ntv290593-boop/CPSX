# Bảng quyết định — 2026-11-09 12:06

## Các cổng kiểm trước khi chấm

```
[1] ✅ chọn điểm            88 chọn (10 mã điểm) / 13400 bỏ (13400 theo nhóm), 10/10 điểm phích khớp
[2] ✅ oracle độc lập       PASS — 7 hàm oracle độc lập
[3] ✅ vân tay nguồn        khớp mốc
[4] ✅ vân tay hàm          khớp mốc (10 điểm)
[5] ✅ case trùng           116 case, không trùng mã/đầu vào
```


## LỚP LUẬT

USER :  116/116  đạt   ✗ sai 0   ⚠ chờ người 0   ! vỡ 0   · chưa chạy 0
PROBE:    0/0    đạt   (không tính vào kết luận nghiệp vụ)   ✗ sai 0   ⚠ chờ người 0   ! vỡ 0   · chưa chạy 0
  · 33/116 case USER chưa có người gật (nguon khác 'nguoi') — chưa được vào bộ vàng.
  (4.05 giây)

| ID | Vai | Điểm | Nhánh | Đầu vào | Nhận được | Chuẩn | KQ | Vì sao |
|---|---|---|---|---|---|---|---|---|
| — | — | — | — | — | — | — | ✓ | không có dòng nào cần xem |

## Độ phủ lớp luật

```

## vai_tro_doan   (SKILL/10-tender/tool/td.py:528)
   mã điểm      : TD.DOAN.VAITRO
   case         : 24 dựng cho điểm này (USER 24 · PROBE 0) + 92 case khác trong bộ
   đường chạy   : 24 đi qua · 92 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 6/6   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/6   ← chỉ chứng minh code chạy được

## loai_trang   (SKILL/10-tender/tool/td.py:1880)
   mã điểm      : TD.SPEC.LOAI_TRANG
   case         : 10 dựng cho điểm này (USER 10 · PROBE 0) + 106 case khác trong bộ
   đường chạy   : 10 đi qua · 106 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 6/6   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/6   ← chỉ chứng minh code chạy được

## tham_chieu   (SKILL/10-tender/tool/td.py:1175)
   mã điểm      : TD.TCQ.THAM_CHIEU
   case         : 8 dựng cho điểm này (USER 8 · PROBE 0) + 108 case khác trong bộ
   đường chạy   : 29 đi qua · 87 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 2/2   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/2   ← chỉ chứng minh code chạy được

## date_in   (SKILL/10-tender/tool/td.py:203)
   mã điểm      : TD.NGAY.DATE_IN
   case         : 11 dựng cho điểm này (USER 11 · PROBE 0) + 105 case khác trong bộ
   đường chạy   : 11 đi qua · 105 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 2/2   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/2   ← chỉ chứng minh code chạy được

## dot_of   (SKILL/10-tender/tool/td.py:218)
   mã điểm      : TD.NGAY.DOT_OF
   case         : 8 dựng cho điểm này (USER 8 · PROBE 0) + 108 case khác trong bộ
   đường chạy   : 8 đi qua · 108 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 3/3   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/3   ← chỉ chứng minh code chạy được

## ma_vl   (SKILL/10-tender/tool/td.py:84)
   mã điểm      : TD.VL.MA_VL
   case         : 14 dựng cho điểm này (USER 14 · PROBE 0) + 102 case khác trong bộ
   đường chạy   : 33 đi qua · 83 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 2/2   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/2   ← chỉ chứng minh code chạy được

## chuan_ben   (SKILL/10-tender/tool/td.py:2433)
   mã điểm      : TD.BEN.CHUAN_BEN
   case         : 8 dựng cho điểm này (USER 8 · PROBE 0) + 108 case khác trong bộ
   đường chạy   : 8 đi qua · 108 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 2/2   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/2   ← chỉ chứng minh code chạy được

## ma_vl_theo_bang   (SKILL/10-tender/tool/td.py:108)
   mã điểm      : TD.VL.MA_BANG
   case         : 11 dựng cho điểm này (USER 11 · PROBE 0) + 105 case khác trong bộ
   đường chạy   : 11 đi qua · 105 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 2/2   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/2   ← chỉ chứng minh code chạy được

## tap_ma   (SKILL/10-tender/tool/dmvt_xlsx.py:49)
   mã điểm      : DMVT.TAP_MA
   case         : 12 dựng cho điểm này (USER 12 · PROBE 0) + 104 case khác trong bộ
   đường chạy   : 12 đi qua · 104 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 2/2   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/2   ← chỉ chứng minh code chạy được

## doc_ncc   (SKILL/10-tender/tool/spec_boc.py:114)
   mã điểm      : SPEC.DOC_NCC
   case         : 10 dựng cho điểm này (USER 10 · PROBE 0) + 106 case khác trong bộ
   đường chạy   : 10 đi qua · 106 không đi qua · 0 chưa có bằng chứng
   độ tin cậy   : CHẮC — dấu vết từ đường chạy thật
   nhánh phủ (USER) : 2/2   ← con số có ý nghĩa nghiệp vụ
   nhánh phủ (PROBE): 0/2   ← chỉ chứng minh code chạy được

============================================================
Lớp luật đã phủ hết nhánh và hết mốc bằng case USER, đo bằng dấu vết hoặc mã nhánh.
```

## Kết luận lượt chạy

```
[1] ✅ chọn điểm            88 chọn (10 mã điểm) / 13400 bỏ (13400 theo nhóm), 10/10 điểm phích khớp
[2] ✅ oracle độc lập       PASS — 7 hàm oracle độc lập
[3] ✅ vân tay nguồn        khớp mốc
[4] ✅ vân tay hàm          khớp mốc (10 điểm)
[5] ✅ case trùng           116 case, không trùng mã/đầu vào
[6] ✅ chấm                 lớp luật xanh
[7] ✅ độ phủ               phủ đủ theo số đếm
mã thoát 0
```

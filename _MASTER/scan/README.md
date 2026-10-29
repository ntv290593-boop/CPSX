# _MASTER\scan – bo quet ho so

| File | Viec |
|---|---|
| scan_schedules.py | pdfplumber: doc bang thong ke cua tu PDF ban ve CAD theo tieu de cot VI/EN; tra ve trang/bang/dong + do tin cay |
| scan_spec.py | Tach spec MasterFormat (PART, 1.01, A., 1.) theo trang tu text pdftotext -layout; gan nhan submittal/qa/material/warranty |
| build_project_data_NORTON.py | Dung PROJECT_DATA.xlsx (BOQ_TONGHOP, ISSUES 3 trang thai, SPEC, NOTES, _LINKS) – ban NORTON, se tong quat hoa |

Docling (OCR file scan) can tai model tu HuggingFace -> chay tren may Windows, khong chay duoc tren cloud.
Link trong PROJECT_DATA can add-in _MASTER\vba\SRC_LINKS.xlam.

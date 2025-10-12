OSEL Panel Generator v1.0
=========================

오성이엘 패널제작 자동작도 프로그램

사용법:
1. excel_files 폴더에 Excel 파일(.xlsx)을 넣어주세요
2. OSEL_Panel_Generator.exe를 실행하세요
3. done 폴더에 생성된 DXF 파일을 확인하세요

필요한 폴더 구조:
├── OSEL_Panel_Generator.exe (실행 파일)
├── dimstyle/
│   └── style.dxf (스타일 파일)
├── excel_files/
│   └── *.xlsx (Excel 파일들)
└── done/ (자동 생성)
    └── *.dxf (생성된 도면 파일들)

주의사항:
- Excel 파일은 '제작산출결과' 시트가 있어야 합니다
- Excel 파일을 실행 중에는 닫아주세요
- done 폴더는 프로그램이 자동으로 생성합니다

문의: 오성이엘
버전: 2025.10.07
# OSEL Panel Generator - 기술 문서

## 📋 개요

**OSEL Panel Generator**는 오성이엘의 패널제작 자동작도 프로그램으로, Excel 데이터를 기반으로 DXF 형식의 CAD 도면을 자동 생성하는 Python 애플리케이션입니다.

### 🎯 주요 기능
- Excel 파일에서 제작 데이터 자동 읽기
- DXF 형식의 CAD 도면 자동 생성
- 패널 치수, 타공, 레이블링 자동화
- 다중 현장 데이터 처리
- 한글 지원 및 스타일 적용

## 🏗️ 시스템 아키텍처

### 핵심 구성 요소
```
osel_panel.py
├── 전역 변수 관리
├── DXF 초기화 및 스타일 설정
├── Excel 데이터 처리
├── 도면 생성 엔진
├── 파일 I/O 관리
└── 에러 처리 및 로깅
```

### 의존성 라이브러리
```python
# 주요 라이브러리
ezdxf              # DXF 파일 생성 및 조작
openpyxl           # Excel 파일 읽기/쓰기
gooey              # GUI 인터페이스 (현재 미사용)
requests           # HTTP 통신 (로깅용)
datetime           # 날짜/시간 처리
re                 # 정규표현식 (파일명 처리)
glob               # 파일 패턴 매칭
os, sys            # 시스템 인터페이스
```

## 📁 파일 구조

### 프로젝트 디렉토리
```
osel/
├── osel_panel.py           # 메인 애플리케이션
├── dimstyle/
│   └── style.dxf          # DXF 스타일 템플릿
├── excel_files/           # 입력 Excel 파일들
├── done/                  # 출력 DXF 파일들 (자동 생성)
├── release/               # 배포용 실행 파일
│   ├── OSEL_Panel_Generator.exe
│   ├── dimstyle/
│   ├── excel_files/
│   └── README.txt
└── dev.md                 # 이 기술 문서
```

### 실행 파일 구조 (PyInstaller)
```
release/
├── OSEL_Panel_Generator.exe  # 단독 실행 파일 (94.9MB)
├── dimstyle/
│   └── style.dxf            # 스타일 템플릿
├── excel_files/             # 입력 데이터
│   └── *.xlsx
└── done/                    # 출력 결과 (자동 생성)
    └── *.dxf
```

## 🔧 핵심 함수 분석

### 1. 초기화 및 설정

#### `initialize_dxf()`
```python
def initialize_dxf():
    """새로운 DXF 문서를 초기화하고 스타일을 설정하는 함수"""
```
- **목적**: DXF 문서 생성 및 스타일 로드
- **기능**: 
  - 새로운 DXF 문서 생성
  - `style.dxf` 로드 및 스타일 적용
  - 텍스트 스타일 및 치수 스타일 설정
- **중요도**: ⭐⭐⭐⭐⭐ (각 Excel 파일마다 호출)

#### 전역 변수 관리
```python
# 핵심 전역 변수
global_data = {}              # 데이터 저장소
pageCount = 0                 # 페이지 카운터
selected_dimstyle = ''        # 선택된 치수 스타일
text_style_name = ''          # 텍스트 스타일명
```

### 2. Excel 데이터 처리

#### `read_manufacturing_results(sheet)`
```python
def read_manufacturing_results(sheet):
    """제작산출결과 시트에서 데이터를 읽어오는 함수"""
```
- **목적**: Excel에서 제작 데이터 추출
- **입력**: Excel 워크시트 객체
- **출력**: 제작 데이터 딕셔너리 리스트
- **컬럼 매핑**:
  ```python
  column_mapping = {
      "L": "panel_number",      # 패널 번호
      "O": "manufacturing_width", # 제작 폭
      "P": "manufacturing_height", # 제작 높이
      # ... 기타 매핑
  }
  ```

#### `read_excel_value(sheet, cell_address)`
```python
def read_excel_value(sheet, cell_address):
    """Excel 셀에서 값을 안전하게 읽어오는 함수"""
```
- **목적**: Excel 셀 값 안전 추출
- **에러 처리**: None 반환으로 예외 상황 처리

### 3. 도면 생성 엔진

#### `execute_panel()`
```python
def execute_panel():
    """패널 도면을 생성하는 메인 함수"""
```
- **목적**: 패널 도면 생성의 핵심 로직
- **주요 기능**:
  - 현장별 데이터 그룹핑
  - 패널 시퀀스 생성 (1-9번)
  - 도면 요소 그리기 (패널, 치수선, 타공, 텍스트)
  - 테이블 생성

#### 도면 그리기 함수들

##### `draw_rectangle(doc, x, y, width, height, layer)`
```python
def draw_rectangle(doc, x, y, width, height, layer=None):
    """사각형 그리기"""
```

##### `draw_Text(doc, x, y, size, text, layer, alignment)`
```python
def draw_Text(doc, x, y, size, text, layer=None, alignment=TextEntityAlignment.BOTTOM_LEFT):
    """텍스트 그리기 (한글 지원)"""
```

##### `draw_dimension_line(doc, x1, y1, x2, y2, distance, text, layer)`
```python
def draw_dimension_line(doc, x1, y1, x2, y2, distance, text, layer=None):
    """치수선 그리기 (스타일 적용)"""
```

##### `draw_table(doc, start_x, start_y, data, layer)`
```python
def draw_table(doc, start_x, start_y, data, layer=None):
    """테이블 그리기"""
```

### 4. 유틸리티 함수

#### `generate_filename(site_names)`
```python
def generate_filename(site_names):
    """현장명을 기반으로 파일명 생성"""
```
- **목적**: 현장명 조합으로 고유 파일명 생성
- **알고리즘**: 공통 접두사 제거 후 조합

#### `show_custom_error(message)`
```python
def show_custom_error(message):
    """커스텀 에러 메시지 표시"""
```

## 📊 데이터 플로우

### 1. 입력 데이터 처리
```
Excel 파일 (.xlsx)
├── 제작산출결과 시트 읽기
├── 컬럼 매핑 (L→panel_number, O→manufacturing_width, P→manufacturing_height)
├── 데이터 검증 및 정규화
└── 현장별 그룹핑
```

### 2. 도면 생성 프로세스
```
현장별 데이터
├── 패널 시퀀스 생성 (Excel 2-10 → Display 1-9)
├── 도면 요소 배치
│   ├── 패널 사각형
│   ├── 치수선 (수평/수직)
│   ├── 타공 표시
│   ├── 패널 번호
│   └── 테이블
├── 레이어 적용 ('레이져', 'DIM', '0')
└── DXF 파일 출력
```

### 3. 출력 파일 생성
```
done/
├── 현장명_날짜시간.dxf
├── 파일명 중복 방지
└── 자동 폴더 생성
```

## 🎨 DXF 스타일 시스템

### 스타일 파일: `style.dxf`
- **텍스트 스타일**: 한글 폰트 지원 (gulim.ttc, Arial.ttf)
- **치수 스타일**: 전문적인 CAD 치수 표시
- **레이어**: 
  - `'레이져'`: 레이저 절단용 패널
  - `'DIM'`: 치수선
  - `'0'`: 기본 레이어

### 스타일 우선순위
```python
# 텍스트 스타일 우선순위
if 'JKW' in available_textstyles:
    text_style_name = 'JKW'      # gulim.ttc 한글 폰트
elif 'mydim1' in available_textstyles:
    text_style_name = 'mydim1'   # gulim.ttc 한글 폰트
elif 'H' in available_textstyles:
    text_style_name = 'H'
else:
    text_style_name = 'Standard'

# 치수 스타일 우선순위
if 'mydim1' in available_dimstyles:
    selected_dimstyle = 'mydim1'
elif 'over1000dim1' in available_dimstyles:
    selected_dimstyle = 'over1000dim1'
else:
    selected_dimstyle = 'Standard'
```

## 🔄 실행 흐름

### 메인 실행 루프
```python
def main():
    # 1. 초기화
    for file_path in xlsm_files:
        # 2. Excel 파일 로드
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        
        # 3. DXF 문서 초기화 (각 파일마다)
        initialize_dxf()
        
        # 4. 데이터 읽기
        global_data["manufacturing_data"] = read_manufacturing_results(sheet)
        
        # 5. 도면 생성
        execute_panel()
        
        # 6. 파일 저장
        doc.saveas(full_file_path)
```

### 에러 처리 전략
- **파일 I/O 오류**: try-except로 안전 처리
- **Excel 파일 접근**: 파일 잠금 상태 감지
- **DXF 생성 오류**: 폴백 메커니즘 제공
- **사용자 알림**: 명확한 오류 메시지

## 🚀 배포 및 실행

### PyInstaller 설정
```bash
pyinstaller --onefile --name OSEL_Panel_Generator --distpath release --workpath build --specpath . osel_panel.py
```

### 실행 파일 특징
- **크기**: 94.9MB (모든 의존성 포함)
- **독립성**: Python 설치 불필요
- **경로 처리**: 실행 위치 기준 상대 경로
- **자동 폴더 생성**: `done` 폴더 자동 생성

### 사용자 가이드
1. `excel_files` 폴더에 Excel 파일(.xlsx) 배치
2. `OSEL_Panel_Generator.exe` 실행
3. `done` 폴더에서 결과 DXF 파일 확인

## 🔧 개발 환경 설정

### 필수 패키지 설치
```bash
pip install ezdxf openpyxl gooey requests
```

### 개발 도구
- **Python**: 3.10.6+
- **IDE**: Visual Studio Code (권장)
- **버전 관리**: Git
- **빌드 도구**: PyInstaller

### 디버깅 팁
- `print()` 문으로 데이터 플로우 추적
- Excel 파일 구조 변경 시 `column_mapping` 수정
- DXF 스타일 문제 시 `style.dxf` 확인

## 📈 성능 최적화

### 메모리 관리
- 각 Excel 파일마다 DXF 문서 초기화
- 전역 변수 재설정으로 메모리 누수 방지
- 대용량 Excel 파일 처리 시 배치 처리

### 실행 속도
- `data_only=True`로 Excel 계산 공식 제외
- 불필요한 라이브러리 로드 최소화
- 효율적인 파일 I/O 패턴

## 🛠️ 유지보수 가이드

### 코드 수정 시 주의사항
1. **전역 변수**: `initialize_dxf()`에서 재설정 확인
2. **Excel 컬럼**: `column_mapping` 수정 시 데이터 구조 확인
3. **DXF 스타일**: `style.dxf` 변경 시 호환성 확인
4. **파일 경로**: 실행 파일과 스크립트 실행 모두 테스트

### 확장 가능성
- **새로운 도면 요소**: 그리기 함수 추가
- **다른 Excel 형식**: `read_manufacturing_results()` 수정
- **추가 출력 형식**: DXF 외 다른 CAD 형식 지원
- **GUI 인터페이스**: Gooey 라이브러리 활용

## 🐛 알려진 이슈 및 해결방법

### 1. 한글 인코딩 문제
- **증상**: 한글 텍스트가 `???`로 표시
- **해결**: `style.dxf`에서 한글 폰트 스타일 확인

### 2. Excel 파일 접근 오류
- **증상**: "파일을 열 수 없습니다" 오류
- **해결**: Excel 파일이 다른 프로그램에서 열려있지 않은지 확인

### 3. DXF 스타일 미적용
- **증상**: 기본 스타일로 도면 생성
- **해결**: `dimstyle/style.dxf` 파일 존재 및 경로 확인

### 4. 메모리 누수 (다중 파일 처리)
- **증상**: 두 번째 파일에서 첫 번째 파일 데이터 포함
- **해결**: `initialize_dxf()` 함수로 각 파일마다 DXF 문서 초기화

## 📝 변경 이력

### v1.0 (2025-10-07)
- 초기 버전 개발
- Excel → DXF 자동 변환 기능
- 다중 현장 데이터 처리
- PyInstaller 단독 실행 파일 생성
- 한글 지원 및 스타일 적용

### 주요 개선사항
- ✅ DXF 문서 초기화 함수 분리
- ✅ 경로 처리 개선 (실행 파일/스크립트 호환)
- ✅ 에러 처리 강화
- ✅ 메모리 누수 방지
- ✅ 사용자 친화적 파일명 생성

## 📞 지원 및 문의

- **개발사**: 오성이엘
- **버전**: 1.0
- **최종 업데이트**: 2025-10-07
- **문서 버전**: 1.0

---

*이 문서는 OSEL Panel Generator의 기술적 세부사항을 다루며, 개발자와 유지보수 담당자를 위한 참고 자료입니다.*

# 다완 방화쟘 자동작도 시스템 개발 문서

## 개발 일자
2024년 (최종 수정: 현재)

## 주요 수정 사항

### 1. ValueError 해결: load_excel() 함수 반환값 불일치

**문제:**
```
ValueError: not enough values to unpack (expected 18, got 17)
ValueError: too many values to unpack (expected 17)
```

**원인:**
- `load_excel()` 함수는 17개 값을 반환
- `execute_wide()` 함수에서는 18개 값을 받으려 시도 (`FireDoor` 추가)
- 4124번째 줄에서는 17개 값만 받으려 시도

**해결책:**
1. `load_excel()` 함수에 `FireDoor` 처리 추가:
   ```python
   FireDoor = row_data.get("FireDoor", 25)  # 기본값 25 (방화도어)
   return jambType, floorDisplay, material, spec, vcut, OP, poleAngle, JE, JD, HH, MH, HPI_height, U, C, A, grounddig, surang, FireDoor
   ```

2. 모든 `load_excel()` 호출 지점에서 `FireDoor` 받도록 수정:
   - 2294번째 줄: ✅ 이미 수정됨
   - 4124번째 줄: ✅ `FireDoor` 추가

### 2. FireDoor 조건부 층 병합 기능 구현

**요구사항:**
같은 크기여도 `FireDoor` 값이 다르면 층을 합치지 않도록 수정

**구현 위치:**
`aggregate_rows()` 함수 (dawan_jamb.py:2014-2107)

**수정 내용:**

1. **FireDoor 값 추출 추가:**
   ```python
   FireDoor = row_data.get("FireDoor", 25)  # 기본값 25 (방화도어)
   ```

2. **병합 키에 FireDoor 포함:**
   ```python
   # 기존
   key = (jambType, material, spec, vcut, OP, JE, JD, HH, MH,
          HPI_height, U, C, A, grounddig, pangle)

   # 수정 후
   key = (jambType, material, spec, vcut, OP, JE, JD, HH, MH,
          HPI_height, U, C, A, grounddig, pangle, FireDoor)
   ```

3. **키 언패킹 및 결과 생성 수정:**
   ```python
   # 키 언패킹시 FireDoor 포함
   (jambType, material, spec, vcut, OP, JE, JD, HH, MH,
    HPI_height, U, C, A, grounddig, pangle, FireDoor) = key

   # final_list에 FireDoor 필드 추가
   "FireDoor": FireDoor,
   ```

4. **함수 문서 업데이트:**
   ```python
   """
   반환값:
     동일한 (jambType, material, spec, vcut, OP, JE, JD, HH, MH, HPI_height,
             U, C, A, grounddig, poleAngle, FireDoor)
     에 대해서 floorDisplay를 콤마로 합쳐 'floorDisplay'로 두고,
     surang에 합쳐진 개수를 저장한 리스트를 반환.
   """
   ```

**동작 원리:**
- 기존: 크기 관련 속성들만 같으면 층 병합
- 수정 후: 크기 + `FireDoor` 값이 모두 같아야 층 병합

**예시:**
```
1층 방화도어 (FireDoor=25) + 2층 일반도어 (FireDoor=50)
→ 같은 크기여도 별도 처리

1층 방화도어 (FireDoor=25) + 2층 방화도어 (FireDoor=25)
→ 같은 크기면 "1F,2F"로 병합
```

## FireDoor 값 의미
- `25`: 방화도어
- `50`: 일반도어

엑셀 Q열에서 "방화"/"일반" 문자열을 25/50 숫자로 변환하여 처리

## 코드 위치 참조

### 주요 함수들:
- `load_excel()`: 2225-2252줄 (엑셀 데이터 로드)
- `aggregate_rows()`: 2014-2107줄 (층 병합 로직)
- `execute_wide()`: 2256-4117줄 (메인 도면 생성)

### 수정된 파일:
- `dawan_jamb.py`: 메인 애플리케이션 파일

### 테스트 방법:
1. 같은 크기, 다른 FireDoor 값을 가진 데이터로 테스트
2. 생성된 DXF 파일에서 층이 별도로 처리되는지 확인
3. 작업완료/ 디렉터리에서 출력 파일 검증

## 향후 개선 사항
- FireDoor 값 검증 로직 강화
- 엑셀 입력 오류 처리 개선
- 로깅 시스템 추가

---
**개발자**: AI Assistant
**버전**: 1.0
**상태**: 완료
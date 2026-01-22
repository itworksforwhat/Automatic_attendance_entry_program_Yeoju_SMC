# 근태 자동 입력 v3.0 - 완전 사용 설명서

## 📚 목차

1. [시스템 개요](#시스템-개요)
2. [프로그램 구조](#프로그램-구조)
3. [핵심 함수 설명](#핵심-함수-설명)
4. [설정 변경 방법](#설정-변경-방법)
5. [유지보수 가이드](#유지보수-가이드)
6. [문제 해결](#문제-해결)
7. [확장 가이드](#확장-가이드)

---

# 시스템 개요

## 프로그램 목적

원시 근태 데이터(.xls)를 읽어서 근태표 엑셀 파일(.xlsx)에 자동으로 출퇴근 시간을 입력하는 프로그램입니다.

## 주요 기능

1. ✅ **공휴일 자동 감지**: 출근 인원을 분석하여 공휴일/주말 자동 판단
2. ✅ **이전 근무일 찾기**: 공휴일을 건너뛰고 실제 근무일 찾기
3. ✅ **야간 근무 처리**: 12시 이후 출근자 자동 인식 및 처리
4. ✅ **시트 자동 복사**: 이전 근무일 시트를 복사하여 새 시트 생성
5. ✅ **문제 데이터 관리**: 문제가 있는 데이터는 별도 파일로 저장
6. ✅ **.xls 직접 읽기**: xlrd를 사용하여 .xls 파일 직접 처리

## 처리 흐름

```
1. 원시 데이터 로드 (.xls)
   ↓
2. 데이터 분석 (공휴일 감지, 이전 근무일 찾기)
   ↓
3. 데이터 검증 (정상/문제 분류)
   ↓
4. 출퇴근 맵 생성 (오늘/전일)
   ↓
5. 근태표에 입력 (Excel COM)
   ↓
6. 문제 데이터 처리 (별도 파일 생성)
```

---

# 프로그램 구조

## 파일 구성

```
attendance_v3/
├── main.py                 # 메인 실행 파일
├── config.py              # 설정 파일 (★ 여기서 수정)
├── data_analyzer.py       # 데이터 분석 엔진
├── attendance_engine.py   # 출퇴근 로직 엔진
├── excel_com.py           # Excel 제어
├── gui.py                 # GUI (재입력 버튼)
├── models.py              # 데이터 모델
├── logger.py              # 로깅
└── requirements.txt       # 필수 패키지
```

## 각 파일 역할

### 1. main.py
- **역할**: 프로그램 시작점
- **주요 함수**:
  - `main()`: 전체 실행 흐름 제어
  - `_execute()`: 실제 작업 수행

### 2. config.py ⭐ 중요!
- **역할**: 모든 설정 관리
- **수정 빈도**: 높음 (블록 범위, 파일 경로 등)

### 3. data_analyzer.py
- **역할**: 데이터 분석 및 검증
- **주요 기능**:
  - 공휴일 자동 감지
  - 시간 파싱
  - 데이터 검증

### 4. attendance_engine.py
- **역할**: 출퇴근 시간 결정 로직
- **주요 기능**:
  - 케이스별 처리 (오늘 출근만, 야간 근무 등)
  - 출퇴근 시간 계산

### 5. excel_com.py
- **역할**: Excel 파일 제어 (읽기/쓰기)
- **주요 기능**:
  - 시트 복사
  - 셀 값 입력
  - 파일 저장

### 6. models.py
- **역할**: 데이터 구조 정의
- **주요 클래스**:
  - `AttendanceRecord`: 출퇴근 기록
  - `ProcessResult`: 처리 결과
  - `WorkPattern`: 근무 패턴

---

# 핵심 함수 설명

## 1. 데이터 분석 (data_analyzer.py)

### `analyze_work_pattern(df)`
**목적**: 근무 패턴 분석 및 공휴일 감지

**입력**:
- `df`: 원시 데이터 DataFrame

**출력**:
- `WorkPattern`: 근무일/공휴일/주말 정보

**로직**:
```python
1. 날짜별 출근 인원 계산
2. 평균 출근 인원 계산
3. 임계값 = 평균 × 30%
4. 임계값 이하 → 공휴일/주말
5. 이전 근무일 찾기 (공휴일 건너뜀)
```

**예시**:
```python
analyzer = DataAnalyzer(logger)
pattern = analyzer.analyze_work_pattern(df)

print(pattern.workdays)       # [2025-12-24, 2025-12-26]
print(pattern.holidays)       # [2025-12-25]
print(pattern.prev_workday)   # 2025-12-24
```

---

### `validate_data(df, base_date)`
**목적**: 데이터 검증 및 문제 데이터 분류

**입력**:
- `df`: 원시 데이터
- `base_date`: 기준 날짜

**출력**:
- `ValidationResult`: 정상/문제 데이터 목록

**검증 항목**:
1. 출근 시간 형식 오류
2. 퇴근 시간 형식 오류
3. 출근만 있음 (퇴근 누락)
4. 퇴근만 있음 (출근 누락)
5. 퇴근 < 출근 (시간 역전)

**예시**:
```python
result = analyzer.validate_data(df, base_date)

print(f"정상: {len(result.valid_records)}건")
print(f"문제: {len(result.problems)}건")

for problem in result.problems:
    print(f"{problem.name}: {problem.issue}")
```

---

### `_parse_time(value)`
**목적**: 다양한 시간 형식을 datetime으로 변환

**지원 형식**:
1. `"2025/12/26 08:30"` → datetime(2025, 12, 26, 8, 30)
2. `"2025-12-26 08:30"` → datetime(2025, 12, 26, 8, 30)
3. `"08:30"` → datetime(2000, 1, 1, 8, 30)
4. `0.35` (Excel 시리얼) → datetime(1899, 12, 30, 8, 24)
5. `datetime` 객체 → 그대로 반환

**반환값**:
- `(datetime, bool)`: (파싱된 시간, 성공 여부)

**예시**:
```python
dt, ok = analyzer._parse_time("2025/12/26 08:30")
print(dt)  # 2025-12-26 08:30:00
print(ok)  # True

dt, ok = analyzer._parse_time("8시")
print(dt)  # 2000-01-01 08:00:00
print(ok)  # False (형식 오류)
```

---

## 2. 출퇴근 로직 (attendance_engine.py)

### `decide_times(name, today_map, yesterday_map)`
**목적**: 출퇴근 시간 결정 (핵심 로직!)

**입력**:
- `name`: 이름
- `today_map`: 오늘 데이터 맵 `{이름: AttendanceRecord}`
- `yesterday_map`: 전일 데이터 맵

**출력**:
- `ProcessResult`: 출근/퇴근 시간, 날짜, 패턴

**처리 케이스**:

#### 케이스 1: 오늘 출근+퇴근 모두 있음
```python
if cin_today and cout_today:
    return ProcessResult(
        check_in=cin_today.strftime("%H:%M"),
        check_out=cout_today.strftime("%H:%M"),
        base_date=today.date,
        pattern="today_complete"
    )
```

#### 케이스 2: 오늘 출근만 있음
```python
if cin_today and not cout_today:
    if cin_today.hour < 12:  # 주간
        # 전일 퇴근 사용
        return ProcessResult(
            check_in=cin_today.strftime("%H:%M"),    # 오늘 출근
            check_out=cout_yest.strftime("%H:%M"),   # 전일 퇴근
            base_date=dout_yest,
            pattern="today_checkin_with_prev_checkout"
        )
    else:  # 야간
        # 출근만 입력
        return ProcessResult(
            check_in=cin_today.strftime("%H:%M"),
            check_out="",
            base_date=today.date,
            pattern="night_shift_no_checkout"
        )
```

#### 케이스 3: 오늘 퇴근만 있음
```python
if not cin_today and cout_today:
    # 전일 출근 사용 (야간 근무 완료)
    return ProcessResult(
        check_in=cin_yest.strftime("%H:%M"),    # 전일 출근
        check_out=cout_today.strftime("%H:%M"), # 오늘 퇴근
        base_date=yesterday.date,
        pattern="prev_night_shift"
    )
```

#### 케이스 4: 오늘 데이터 없음
```python
if not cin_today and not cout_today:
    if cin_yest and cout_yest:
        if cin_yest.hour >= 12:  # 야간 근무자
            # 전일 출근+퇴근 모두 사용
            return ProcessResult(
                check_in=cin_yest.strftime("%H:%M"),
                check_out=cout_yest.strftime("%H:%M"),
                base_date=dout_yest,
                pattern="prev_night_shift_complete"
            )
        else:  # 주간 근무자 결근
            # 퇴근만 사용
            return ProcessResult(
                check_in="",
                check_out=cout_yest.strftime("%H:%M"),
                base_date=dout_yest,
                pattern="absent_with_prev_checkout"
            )
```

**예시**:
```python
engine = AttendanceEngine(pattern, logger)
result = engine.decide_times("홍길동", today_map, yesterday_map)

print(result.check_in)   # "08:30"
print(result.check_out)  # "17:45"
print(result.pattern)    # "today_checkin_with_prev_checkout"
```

---

## 3. Excel 제어 (excel_com.py)

### `prepare_sheet(sheet_name, clear_ranges)`
**목적**: 시트 복사 및 셀 지우기

**입력**:
- `sheet_name`: 새 시트 이름 (예: "25.12.26")
- `clear_ranges`: 지울 범위 리스트 (예: ["D9:E11", "K9:L18"])

**동작**:
1. 마지막 시트 복사
2. 이름을 `sheet_name`으로 변경
3. `clear_ranges` 범위의 값 삭제
4. 파일 저장

**예시**:
```python
excel = ExcelCOMHandler(logger)
excel.open("근태표.xlsx")
excel.prepare_sheet("25.12.26", ["D9:E11", "K9:L18"])
excel.close()
```

---

### `write_attendance(blocks, today_map, yesterday_map, engine)`
**목적**: 출퇴근 데이터 입력

**입력**:
- `blocks`: 블록 리스트 `[(이름범위, 출근범위, 퇴근범위), ...]`
- `today_map`: 오늘 맵
- `yesterday_map`: 전일 맵
- `engine`: AttendanceEngine 인스턴스

**동작**:
1. 각 블록의 이름 범위 순회
2. `decide_times()`로 출퇴근 시간 결정
3. 출근/퇴근 셀에 값 입력 (텍스트 형식: `'08:30`)
4. 파일 저장

**예시**:
```python
blocks = [
    ("C9:C11", "D9:D11", "E9:E11"),    # 블록 1
    ("J9:J18", "K9:K18", "L9:L18"),   # 블록 2
]

excel.write_attendance(blocks, today_map, yesterday_map, engine)
```

---

# 설정 변경 방법

## config.py 주요 설정

### 1. 파일 경로 설정

```python
# 원시 데이터 파일 경로
RAW_DATA_FILE = r"C:/Users/관리부서브/KJH/코딩/12.26.xls"

# 근태표 파일 경로
ATTENDANCE_FILES = {
    "여주": r"C:/Users/관리부서브/KJH/코딩/일일근태보고-여주(202512).xlsx",
    "SMC": r"C:/Users/관리부서브/KJH/코딩/일일근태보고-SMC(202512).xlsx",
}
```

**변경 방법**:
1. 파일 경로를 복사
2. `r"복사한_경로"` 형식으로 입력
3. 역슬래시(`\`) 또는 슬래시(`/`) 모두 가능

---

### 2. 여주 근태표 블록 설정

```python
YEOJU_BLOCKS = [
    ("C9:C11", "D9:D11", "E9:E11"),      # 개발팀
    ("J9:J18", "K9:K18", "L9:L18"),     # 생산1과
    ("Q9:Q18", "R9:R18", "S9:S18"),     # 생산2과
    # ... 더 많은 블록
]
```

**블록 구조**:
- `(이름범위, 출근범위, 퇴근범위)`
- 예: `("C9:C11", "D9:D11", "E9:E11")`
  - C9:C11 = 이름
  - D9:D11 = 출근 시간
  - E9:E11 = 퇴근 시간

**새 블록 추가 방법**:
```python
YEOJU_BLOCKS = [
    # 기존 블록들...
    
    # 새 블록 추가
    ("AB9:AB15", "AC9:AC15", "AD9:AD15"),  # 새로운 팀
]
```

---

### 3. 셀 지우기 범위 설정

```python
CLEAR_RANGES_YEOJU = [
    "D9:E11",   # 개발팀 출퇴근
    "G9:G11",   # 개발팀 잔업
    "K9:L18",   # 생산1과 출퇴근
    # ... 더 많은 범위
]
```

**범위 추가 방법**:
```python
CLEAR_RANGES_YEOJU = [
    # 기존 범위들...
    
    "AC9:AD15",  # 새 팀 출퇴근
    "AF9:AF15",  # 새 팀 잔업
]
```

---

### 4. 공휴일 감지 설정

```python
# 공휴일 임계값 (평균 출근 인원의 %)
HOLIDAY_THRESHOLD = 0.3  # 30%

# 최소 출근 인원 (이하면 무조건 공휴일)
MIN_ATTENDANCE = 5
```

**조정 방법**:
- `HOLIDAY_THRESHOLD` 낮추기 → 더 쉽게 공휴일 인식
- `HOLIDAY_THRESHOLD` 높이기 → 더 엄격하게 판단

**예시**:
```python
# 평균 40명 출근하는 회사
# HOLIDAY_THRESHOLD = 0.3 → 12명 이하면 공휴일
# HOLIDAY_THRESHOLD = 0.5 → 20명 이하면 공휴일
```

---

### 5. 컬럼명 설정

```python
# 원시 데이터 컬럼명 (매핑됨)
COL_DATE = '근무일자'
COL_NAME = '이름'
COL_IN_RAW = '출근시간'
COL_OUT_RAW = '퇴근시간'
```

**원시 데이터 컬럼명이 다른 경우**:
```python
# 예: 출근시간 → 출근
COL_IN_RAW = '출근'
COL_OUT_RAW = '퇴근'
```

---

# 유지보수 가이드

## 1. 새로운 팀/부서 추가

### 단계 1: 블록 확인
1. 근태표 엑셀 열기
2. 새 팀의 이름/출근/퇴근 범위 확인
3. 예: 이름=C20:C25, 출근=D20:D25, 퇴근=E20:E25

### 단계 2: config.py 수정
```python
YEOJU_BLOCKS = [
    # 기존 블록들...
    
    # 새 팀 추가
    ("C20:C25", "D20:D25", "E20:E25"),  # 신규팀
]

CLEAR_RANGES_YEOJU = [
    # 기존 범위들...
    
    "D20:E25",  # 신규팀 출퇴근
    "G20:G25",  # 신규팀 잔업 (있다면)
]
```

### 단계 3: 테스트
```bash
python main.py
```

---

## 2. 근태표 양식 변경

### 컬럼 위치가 변경된 경우

**Before**:
```
| 이름(C) | 출근(D) | 퇴근(E) |
```

**After**:
```
| 이름(C) | 잔업(D) | 출근(E) | 퇴근(F) |
```

**수정**:
```python
# config.py
YEOJU_BLOCKS = [
    # Before
    ("C9:C11", "D9:D11", "E9:E11"),  # ❌
    
    # After
    ("C9:C11", "E9:E11", "F9:F11"),  # ✅
]
```

---

## 3. 원시 데이터 형식 변경

### 시간 형식이 변경된 경우

현재 지원 형식:
1. `2025/12/26 08:30`
2. `2025-12-26 08:30`
3. `08:30`
4. Excel 시리얼 (float)

**새 형식 추가 방법**:

`data_analyzer.py` → `_parse_time()` 함수 수정:

```python
def _parse_time(self, value):
    # ... 기존 코드 ...
    
    if isinstance(value, str):
        value = value.strip()
        
        try:
            # 새 형식 추가
            if '.' in value and '/' not in value:
                # "08.30" 형식
                parts = value.split('.')
                hour = int(parts[0])
                minute = int(parts[1])
                return datetime(2000, 1, 1, hour, minute), True
```

---

## 4. 로직 변경

### 퇴근 시간 우선순위 변경

**현재**: 전일 퇴근 사용
**변경**: 평균 퇴근 시간 사용

`attendance_engine.py` → `decide_times()` 함수 수정:

```python
# Before
if cout_yest:
    return ProcessResult(
        check_in=cin_today.strftime("%H:%M"),
        check_out=cout_yest.strftime("%H:%M"),  # 전일
        ...
    )

# After
default_checkout = "17:30"  # 기본 퇴근 시간
if cout_yest:
    return ProcessResult(
        check_in=cin_today.strftime("%H:%M"),
        check_out=default_checkout,  # 고정값
        ...
    )
```

---

## 5. 새 근태표 추가

### 예: 천안 공장 추가

**단계 1**: `config.py`에 설정 추가

```python
# 파일 경로
ATTENDANCE_FILES = {
    "여주": r"...",
    "SMC": r"...",
    "천안": r"C:/근태표/천안(202512).xlsx",  # 추가
}

# 블록 정의
CHEONAN_BLOCKS = [
    ("C9:C15", "D9:D15", "E9:E15"),
    ("J9:J20", "K9:K20", "L9:L20"),
]

# 지울 범위
CLEAR_RANGES_CHEONAN = [
    "D9:E15", "K9:L20",
]
```

**단계 2**: `main.py`에 처리 추가

```python
def _execute(self):
    # ... 기존 코드 ...
    
    # 여주 처리
    self._process_attendance("여주", config.YEOJU_BLOCKS, config.CLEAR_RANGES_YEOJU)
    
    # SMC 처리
    self._process_attendance("SMC", config.SMC_BLOCKS, config.CLEAR_RANGES_SMC)
    
    # 천안 처리 추가
    self._process_attendance("천안", config.CHEONAN_BLOCKS, config.CLEAR_RANGES_CHEONAN)
```

---

# 문제 해결

## 문제 1: "입력: 0건"

### 원인
이름 매칭 실패

### 해결
1. 원시 데이터와 근태표의 이름이 정확히 일치하는지 확인
2. 공백, 특수문자 확인
3. 로그에서 `'xxx': 원시 데이터에서 찾을 수 없음` 확인

### 예시
```
근태표: "홍길동"
원시 데이터: "홍 길동"  ❌

→ 원시 데이터에서 이름 수정 필요
```

---

## 문제 2: 시간이 "7:46:00 AM" 형식으로 표시

### 원인
Excel이 텍스트를 시간으로 자동 변환

### 해결
✅ 이미 수정됨 (작은따옴표 `'` 추가)

현재 버전은 `'08:30` 형식으로 입력하여 텍스트 강제

---

## 문제 3: 공휴일이 감지되지 않음

### 원인
임계값이 너무 낮음

### 해결
`config.py`에서 `HOLIDAY_THRESHOLD` 조정

```python
# 더 쉽게 감지
HOLIDAY_THRESHOLD = 0.5  # 50%

# 더 엄격하게
HOLIDAY_THRESHOLD = 0.2  # 20%
```

---

## 문제 4: 특정 인원의 시간이 틀림

### 원인
1. 원시 데이터가 잘못됨
2. 로직 오류
3. 이름 매칭 오류

### 해결
1. 로그 확인:
```
[DEBUG]   처리 중: 'xxx'
[DEBUG]     출퇴근 시간: cin_today=..., cout_today=...
```

2. 원시 데이터 직접 확인
3. 문제_데이터_확인.xlsx 파일 확인

---

## 문제 5: 파일을 열 수 없음

### 원인
Excel 파일이 이미 열려있음

### 해결
1. 모든 Excel 파일 닫기
2. 프로그램 재실행

---

## 문제 6: 시트 복사가 안 됨

### 원인
마지막 시트 이름 형식 불일치

### 해결
마지막 시트 이름이 "YY.MM.DD" 형식인지 확인
- 예: "25.12.24" ✅
- 예: "2025-12-24" ❌

---

# 확장 가이드

## 1. GUI 개선

현재: 재입력 버튼만 있음

**확장 아이디어**:
```python
# gui.py 수정

import tkinter as tk
from tkinter import ttk, filedialog

class AttendanceGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("근태 자동 입력 v3.0")
        
        # 파일 선택 버튼 추가
        ttk.Button(self.window, text="원시 데이터 선택", 
                   command=self.select_raw_file).pack()
        
        # 진행 표시줄 추가
        self.progress = ttk.Progressbar(self.window, length=300)
        self.progress.pack()
    
    def select_raw_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        # 파일 경로 저장
```

---

## 2. 이메일 알림

처리 완료 시 이메일 발송

```python
# main.py에 추가

import smtplib
from email.mime.text import MIMEText

def send_completion_email(processed_count):
    msg = MIMEText(f"근태 입력 완료: {processed_count}명 처리됨")
    msg['Subject'] = "근태 자동 입력 완료"
    msg['From'] = "sender@company.com"
    msg['To'] = "manager@company.com"
    
    with smtplib.SMTP('smtp.company.com') as server:
        server.send_message(msg)

# _execute() 끝에 추가
send_completion_email(processed)
```

---

## 3. 데이터베이스 연동

처리 이력을 DB에 저장

```python
# 새 파일: database.py

import sqlite3
from datetime import datetime

class AttendanceDB:
    def __init__(self, db_path="attendance.db"):
        self.conn = sqlite3.connect(db_path)
        self._create_tables()
    
    def _create_tables(self):
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY,
                date TEXT,
                name TEXT,
                check_in TEXT,
                check_out TEXT,
                pattern TEXT,
                created_at TEXT
            )
        ''')
    
    def save_record(self, date, name, check_in, check_out, pattern):
        self.conn.execute(
            'INSERT INTO history VALUES (?, ?, ?, ?, ?, ?, ?)',
            (None, date, name, check_in, check_out, pattern, 
             datetime.now().isoformat())
        )
        self.conn.commit()
```

---

## 4. 웹 인터페이스

Flask 웹 앱으로 변환

```python
# 새 파일: web_app.py

from flask import Flask, render_template, request, jsonify
import main

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    base_date = request.form['date']
    result = main.execute_with_date(base_date)
    return jsonify(result)

if __name__ == '__main__':
    app.run(debug=True)
```

---

## 5. 통계 및 분석

출근율, 지각율 등 통계

```python
# 새 파일: statistics.py

from datetime import datetime
from collections import defaultdict

class AttendanceStats:
    def __init__(self, records):
        self.records = records
    
    def calculate_attendance_rate(self):
        """출근율 계산"""
        total = len(self.records)
        present = sum(1 for r in self.records if r.check_in)
        return (present / total) * 100
    
    def calculate_late_rate(self):
        """지각율 계산 (9시 기준)"""
        total = sum(1 for r in self.records if r.check_in)
        late = sum(1 for r in self.records 
                  if r.check_in and r.check_in.hour >= 9)
        return (late / total) * 100 if total > 0 else 0
    
    def get_average_work_hours(self):
        """평균 근무 시간"""
        hours = []
        for r in self.records:
            if r.check_in and r.check_out:
                diff = (r.check_out - r.check_in).total_seconds() / 3600
                hours.append(diff)
        return sum(hours) / len(hours) if hours else 0
```

---

# 부록

## A. 데이터 흐름도

```
원시 데이터 (.xls)
    ↓
[load_raw_data]
    ↓
DataFrame
    ↓
[analyze_work_pattern] ← 공휴일 감지
    ↓
WorkPattern (이전 근무일, 공휴일 목록)
    ↓
[validate_data] ← 데이터 검증
    ↓
ValidationResult (정상/문제)
    ↓
[create_maps] ← 오늘/전일 맵 생성
    ↓
{이름: AttendanceRecord}
    ↓
[decide_times] ← 출퇴근 시간 결정
    ↓
ProcessResult (출근, 퇴근, 패턴)
    ↓
[write_attendance] ← Excel 입력
    ↓
근태표 파일 (.xlsx)
```

---

## B. 케이스별 처리 요약

| 케이스 | 오늘 출근 | 오늘 퇴근 | 전일 출근 | 전일 퇴근 | 입력 출근 | 입력 퇴근 | 패턴 |
|--------|----------|----------|----------|----------|----------|----------|------|
| 1 | ✅ | ✅ | - | - | 오늘 | 오늘 | today_complete |
| 2-1 | ✅(주간) | ❌ | - | ✅ | 오늘 | 전일 | today_checkin_with_prev_checkout |
| 2-2 | ✅(야간) | ❌ | - | - | 오늘 | 빈칸 | night_shift_no_checkout |
| 3 | ❌ | ✅ | ✅ | - | 전일 | 오늘 | prev_night_shift |
| 4-1 | ❌ | ❌ | ✅(야간) | ✅ | 전일 | 전일 | prev_night_shift_complete |
| 4-2 | ❌ | ❌ | ✅(주간) | ✅ | 빈칸 | 전일 | absent_with_prev_checkout |
| 5 | ❌ | ❌ | ❌ | ❌ | 빈칸 | 빈칸 | no_data |

---

## C. 에러 코드

| 코드 | 의미 | 해결 방법 |
|------|------|----------|
| E001 | 파일을 찾을 수 없음 | 파일 경로 확인 |
| E002 | 컬럼을 찾을 수 없음 | 컬럼명 매핑 확인 |
| E003 | 시간 파싱 실패 | 시간 형식 확인 |
| E004 | Excel COM 오류 | Excel 파일 닫기 |
| E005 | 이름 매칭 실패 | 이름 일치 확인 |

---

## D. 성능 최적화

### 대용량 데이터 처리

```python
# data_analyzer.py

# Before
for idx, row in df.iterrows():  # 느림
    ...

# After
for row in df.itertuples():  # 빠름
    ...
```

### Excel 입력 최적화

```python
# excel_com.py

# Before (개별 입력)
for name in names:
    cell.Value = value  # 매번 COM 호출

# After (배치 입력)
values = [[v1], [v2], [v3]]
range.Value = values  # 한 번에 입력
```

---

## E. 테스트

### 단위 테스트 예시

```python
# test_analyzer.py

import unittest
from data_analyzer import DataAnalyzer
from datetime import datetime

class TestDataAnalyzer(unittest.TestCase):
    def setUp(self):
        self.analyzer = DataAnalyzer(None)
    
    def test_parse_time_with_slash(self):
        dt, ok = self.analyzer._parse_time("2025/12/26 08:30")
        self.assertEqual(dt.hour, 8)
        self.assertEqual(dt.minute, 30)
        self.assertTrue(ok)
    
    def test_parse_time_with_hyphen(self):
        dt, ok = self.analyzer._parse_time("2025-12-26 08:30")
        self.assertEqual(dt.hour, 8)
        self.assertTrue(ok)

if __name__ == '__main__':
    unittest.main()
```

---

## F. 버전 관리

### 변경 이력

**v3.0** (2025-12-26)
- 전면 재작성
- 공휴일 자동 감지
- 야간 근무 처리
- .xls 직접 지원

**v2.0** (이전)
- 기본 기능 구현

---

## G. 연락처

**문제 발생 시**:
1. 로그 파일 확인 (`attendance.log`)
2. 문제_데이터_확인.xlsx 확인
3. 개발자에게 문의

---

**문서 버전**: v3.0
**최종 수정**: 2025-12-26
**작성자**: Claude AI

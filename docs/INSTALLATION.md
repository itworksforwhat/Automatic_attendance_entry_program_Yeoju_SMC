# 설치 가이드

> 프로그램 설치하고 돌리는 법

## 시스템 요구사항

### 필수
- **Windows 10 이상** (맥/리눅스 안됨)
- **Python 3.8 이상**
- **Microsoft Excel** (COM 인터페이스 사용)

### 권장
- Python 3.10
- Excel 2016 이상

## 설치 단계

### 1. Python 설치

Python이 없으면 먼저 설치:

1. https://python.org 접속
2. Download → Windows용 최신 버전 다운로드
3. 설치 시 **"Add Python to PATH"** 체크 (중요!)
4. Install Now 클릭

확인:
```bash
python --version
```

`Python 3.x.x` 나오면 성공

### 2. 프로젝트 클론

#### Git 있으면
```bash
git clone https://github.com/itworksforwhat/Automatic_attendance_entry_program_Yeoju_SMC.git
cd Automatic_attendance_entry_program_Yeoju_SMC
```

#### Git 없으면
1. GitHub에서 Code → Download ZIP
2. 압축 풀기
3. 명령 프롬프트에서 해당 폴더로 이동

### 3. 패키지 설치

```bash
cd src
pip install -r requirements.txt
```

#### 설치되는 패키지
- `pandas`: 데이터 처리
- `openpyxl`: Excel 파일 읽기
- `pywin32`: Excel COM 자동화
- `xlrd`: .xls 파일 지원

### 4. pywin32 설정 (중요!)

```bash
python -m pywin32_postinstall -install
```

이거 안 하면 Excel COM 오류남.

### 5. 실행 테스트

```bash
python main.py
```

GUI 창 뜨면 성공!

## 문제 해결

### "pip를 찾을 수 없습니다"

Python 설치 시 PATH 추가 안 한 경우:

1. Python 재설치
2. "Add Python to PATH" 체크
3. 또는 수동으로 환경 변수 추가

### "pywin32를 찾을 수 없습니다"

```bash
pip install --upgrade pywin32
python -m pywin32_postinstall -install
```

### "Excel을 열 수 없습니다"

- Excel 설치 확인
- 관리자 권한으로 실행
```bash
# 관리자 권한 PowerShell에서
python main.py
```

### "ModuleNotFoundError: No module named 'xxx'"

해당 패키지 설치:
```bash
pip install xxx
```

또는 전체 재설치:
```bash
pip install -r requirements.txt --force-reinstall
```

### "xlrd.biffh.XLRDError: Excel xlsx file; not supported"

```bash
pip uninstall xlrd
pip install xlrd==2.0.1
```

### COM 오류 (pywintypes.com_error)

1. Excel 완전히 종료
2. 프로그램 재시작
3. 안 되면 PC 재부팅

## 가상환경 사용 (선택사항)

깔끔하게 설치하고 싶으면:

### 가상환경 생성
```bash
python -m venv venv
```

### 가상환경 활성화

**PowerShell:**
```powershell
.\venv\Scripts\Activate.ps1
```

**CMD:**
```cmd
venv\Scripts\activate.bat
```

### 패키지 설치
```bash
cd src
pip install -r requirements.txt
```

### 실행
```bash
python main.py
```

### 비활성화
```bash
deactivate
```

## PowerShell 실행 정책 오류

```
& : 이 시스템에서 스크립트를 실행할 수 없으므로...
```

해결:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

또는 Python 직접 실행:
```powershell
& python src/main.py
```

## 업데이트

### Git 사용
```bash
git pull origin claude/overtime-tracking-system-PljmE
```

### 수동
1. 최신 버전 다운로드
2. 기존 파일 백업
3. 새 파일로 교체

## 확인 사항

설치 완료 후 확인:

✅ Python 버전 확인: `python --version`
✅ pip 작동 확인: `pip --version`
✅ 패키지 설치 확인: `pip list`
✅ 프로그램 실행 확인: `python main.py`
✅ GUI 창 정상 표시 확인

모두 성공하면 설치 완료!

## 다음 단계

[사용자 가이드](USER_GUIDE.md)를 읽고 프로그램 사용법 익히기

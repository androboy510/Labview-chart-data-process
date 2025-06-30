# 🚀 프로그램 실행 가이드

## 📋 빠른 시작

### 기본 실행 방법
```bash
# 1. 프로젝트 폴더로 이동
cd C:\Users\andro\Desktop\Data_process

# 2. 프로그램 실행
python src/excel_processor_gui_v4.py
```

## 🛠️ 설치 및 설정

### 1. 의존성 설치
```bash
pip install -r requirements.txt
```

### 2. 드래그 앤 드롭 기능 (선택사항)
```bash
pip install tkinterdnd2
```

## 🎯 사용 방법

### 기본 사용법
1. **파일 선택**: "파일 선택" 버튼 클릭 또는 파일을 드래그 앤 드롭
2. **자동 처리**: 선택한 파일이 자동으로 처리됩니다
3. **결과 확인**: 미리보기에서 처리된 데이터 확인
4. **자동 저장**: 원본 파일 위치에 `_processed.xlsx` 확장자로 저장

### 고급 기능
- **다중 파일 처리**: 여러 파일을 동시에 선택하여 일괄 처리
- **진행률 모니터링**: 실시간 진행 상황 확인
- **상태 표시**: 처리 완료 및 오류 상태를 상태바에서 확인

## 🔧 실행 파일 생성 (선택사항)

### PyInstaller 설치
```bash
pip install pyinstaller
```

### 실행 파일 빌드
```bash
python -m PyInstaller --onefile --windowed src/excel_processor_gui_v4.py
```

### 실행 파일 사용
- `dist/excel_processor_gui_v4.exe` 파일을 더블클릭하여 실행

## 📁 지원 파일 형식

### 입력 파일
- `.xlsx` (Excel 파일)
- `.csv` (CSV 파일)

### 출력 파일
- `.xlsx` (처리된 Excel 파일)

## 🚨 문제 해결

### 일반적인 오류
1. **Python이 설치되지 않은 경우**
   - [Python 공식 사이트](https://www.python.org/downloads/)에서 다운로드

2. **패키지 설치 오류**
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

3. **드래그 앤 드롭이 작동하지 않는 경우**
   ```bash
   pip install tkinterdnd2
   ```

### 디버깅
- 터미널에서 실행하여 오류 메시지 확인
- 상태바에서 처리 상태 확인

## 📞 지원

문제가 발생하면 [GitHub Issues](https://github.com/androboy510/Labview-chart-data-process/issues)에 문의해 주세요.

---

**최종 업데이트**: 2025-06-30 
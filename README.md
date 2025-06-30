# Labview Waveform Chart Data Processor

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

> **Labview Waveform Chart Data Processor**는 Labview에서 생성된 웨이브폼 차트 데이터를 자동으로 처리하고 분석하는 전문 도구입니다.

## 🚀 주요 기능

- **📊 자동 데이터 처리**: Labview 웨이브폼 차트 데이터의 자동 정리 및 분석
- **🎨 모던 GUI**: 애플/스타트업 스타일의 직관적이고 세련된 사용자 인터페이스
- **⚡ 빠른 처리**: 다중 파일 대기열 시스템으로 효율적인 배치 처리
- **🖱️ 드래그 앤 드롭**: 파일을 직접 드래그하여 간편한 데이터 업로드
- **💾 스마트 저장**: 원본 파일 위치에 자동 저장으로 파일 관리 편의성
- **📈 실시간 진행률**: 처리 과정을 실시간으로 모니터링

## 📋 요구사항

- **Python**: 3.8 이상
- **운영체제**: Windows 10/11
- **메모리**: 최소 4GB RAM 권장

## 🛠️ 설치 방법

### 1. 저장소 클론
```bash
git clone https://github.com/androboy510/Labview-chart-data-process.git
cd Labview-chart-data-process
```

### 2. 의존성 설치
```bash
pip install -r requirements.txt
```

### 3. 실행
```bash
python src/excel_processor_gui_v4.py
```

## 📦 실행 파일 다운로드

최신 릴리즈에서 실행 파일(.exe)을 다운로드하여 바로 사용할 수 있습니다:

1. [Releases](https://github.com/androboy510/Labview-chart-data-process/releases) 페이지 방문
2. 최신 버전의 `excel_processor_gui_v4.exe` 다운로드
3. 다운로드한 파일을 더블클릭하여 실행

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

## 📁 프로젝트 구조

```
Labview-chart-data-process/
├── src/                    # 소스 코드
│   ├── excel_processor_gui_v4.py    # 메인 GUI 애플리케이션
│   └── excel_processor.py           # 데이터 처리 엔진
├── docs/                   # 문서
│   ├── READMEv2.md         # 상세 사용법
│   ├── PROJECT_SUMMARY.md  # 프로젝트 요약
│   └── TASKS.md            # 개발 작업 목록
├── examples/               # 예제 파일
│   └── create_sample_data.py
├── requirements.txt        # Python 의존성
├── .gitignore             # Git 무시 파일
└── README.md              # 프로젝트 개요
```

## 🔧 기술 스택

- **Frontend**: tkinter, tkinterdnd2
- **Data Processing**: pandas, numpy
- **File I/O**: openpyxl, xlsxwriter
- **Build Tool**: PyInstaller
- **Version Control**: Git

## 🚀 개발 환경 설정

### 로컬 개발
```bash
# 가상환경 생성
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 의존성 설치
pip install -r requirements.txt

# 개발 모드 실행
python src/excel_processor_gui_v4.py
```

### 실행 파일 빌드
```bash
# PyInstaller 설치
pip install pyinstaller

# 실행 파일 생성
pyinstaller --onefile --windowed src/excel_processor_gui_v4.py
```

## 📊 데이터 처리 기능

### 자동 처리 항목
- **헤더 정리**: 불필요한 공백 및 특수문자 제거
- **샘플 데이터 처리**: 샘플 열의 자동 인식 및 처리
- **시간 열 추가**: 샘플 데이터 기반 시간 축 자동 생성
- **데이터 정규화**: 일관된 형식으로 데이터 표준화

### 지원 파일 형식
- **입력**: `.xlsx`, `.csv`
- **출력**: `.xlsx` (처리된 데이터)

## 🤝 기여하기

1. 이 저장소를 포크합니다
2. 새로운 기능 브랜치를 생성합니다 (`git checkout -b feature/AmazingFeature`)
3. 변경사항을 커밋합니다 (`git commit -m 'Add some AmazingFeature'`)
4. 브랜치에 푸시합니다 (`git push origin feature/AmazingFeature`)
5. Pull Request를 생성합니다

## 📝 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 📞 지원

- **이슈 리포트**: [GitHub Issues](https://github.com/androboy510/Labview-chart-data-process/issues)
- **문서**: [docs/](docs/) 폴더 참조
- **예제**: [examples/](examples/) 폴더 참조

## 🔄 업데이트 내역

최신 업데이트는 [docs/READMEv2.md](docs/READMEv2.md)에서 확인할 수 있습니다.

---

**개발자**: [androboy510](https://github.com/androboy510)  
**최종 업데이트**: 2025-06-30 
# SCU 스토리보드(SB) 자동 검수 시스템 (SB-Review-Automation)

서울사이버대학교(SCU) 콘텐츠 설계 및 품질검수 지침과 표준 SB 양식에 맞춘 자동 검수 및 품질 보증 패키지입니다.

---

## 1. 주요 기능 및 특징

- **2-Track 하이브리드 검수**:
  - **Track 1 (정량 검사 - Python)**: 슬라이드 규격(16:9, 33.867x19.05cm), 본문 폰트 크기(20pt Bold 준수 여부), 화면 텍스트 이탈(오버플로우), 슬라이드 노트 출처 누락 등 자동 스캔
  - **Track 2 (정성 검수 - LLM Agent)**: 원고(HWP/HWPX/DOCX/PDF/PPTX) 대조 내용 누락·왜곡 검증, 교수설계 논리(학습목표 행동동사, 퀴즈 정답/해설 일치, AI 추천 학습정리 작성) 점검
- **1쌍(Pair) 동시 산출물 생성**:
  - `output/SB_검수결과_<파일명>_YYMMDD_HHMM.md`: 표준 4대 항목(총평, 내용 검수, 레이아웃 검수, 저작권/출처 점검) 마크다운 보고서
  - `output/SB_검수결과_<파일명>_YYMMDD_HHMM.pptx`: 보고서 지적사항을 파워포인트 네이티브 '메모(Comment)'로 자동 주입한 결과물
- **LCDMS 연동 자동 다운로드**:
  - 서울사이버대학교 LCDMS와 연동하여 과목별/주차별 원고 및 SB 초안을 안전하게 자동 수집 (Read-Only 모드)

---

## 2. 프로젝트 폴더 구조

```text
SB-Review-Automation/
├── .agents/                    # 에이전트 스킬 설정
├── SB-Design/
│   ├── contents/               # 과목별 메타데이터
│   ├── criteria/               # 품질검수 기준 및 프롬프트 템플릿
│   ├── rendered_slides/        # 슬라이드 시각 교차 검증용 렌더링 임시 디렉토리
│   ├── scripts/                # 핵심 검수 스크립트 모듈
│   │   ├── doc_parser.py       # 원고 파서 (HWP, HWPX, DOCX, PDF, PPTX)
│   │   ├── sb_parser.py        # SB 파서
│   │   ├── pptx_layout_checker.py # 레이아웃 검사기
│   │   ├── report_generator.py # 보고서 생성기
│   │   ├── add_comments_to_pptx.py # PPTX 네이티브 메모 주입기
│   │   ├── lcdms_downloader.py # LCDMS 원고/SB 다운로더
│   │   └── slide_renderer.py   # 슬라이드 렌더러
│   └── run_sb_audit.py         # SB 검수 오케스트레이터
├── input/                      # 검수 대상 원고 및 SB 파일 보관 폴더
├── output/                     # 생성된 검수 보고서(.md) 및 메모 반영 PPTX(.pptx)
├── setting/                    # LCDMS 과목 URL 및 세션 설정
├── AGENTS.md                   # 에이전트 지침
├── GEMINI.md                   # 상세 검수 가이드라인 및 워크플로우 규칙
└── README.md
```

---

## 3. 설치 및 실행 방법

### 1) 환경 설정 및 의존성 설치
```bash
pip install python-pptx python-docx olefile pdfplumber requests beautifulsoup4
```

### 2) LCDMS 자료 다운로드 (선택 사항)
1. `setting/contents_set_url.md`에 과목명과 URL 등록
2. `setting/lcdms_cookie.txt`에 LCDMS `tongken` 쿠키 등록
3. 다운로드 스크립트 실행:
   ```bash
   python SB-Design/scripts/lcdms_downloader.py <과목명> <주차>
   # 예: python SB-Design/scripts/lcdms_downloader.py 머신러닝 12
   ```

### 3) 검수 실행
1. `input/` 폴더에 검수 대상 원고 파일과 SB 파일을 배치합니다.
2. 터미널 또는 에이전트를 통해 검수를 실행합니다:
   ```bash
   python SB-Design/run_sb_audit.py
   ```
3. 에이전트 환경에서는 **"SB검수해줘"** 한마디로 자동 실행됩니다.
4. `output/` 폴더에 생성된 보고서(.md)와 검수메모 PPTX(.pptx)를 확인합니다.

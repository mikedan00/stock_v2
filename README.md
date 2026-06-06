# 📊 주식 AI 브리핑 시스템 v2

HuggingFace Router (Gemma-4 26B) 기반 업그레이드 버전

---

## ✨ v2 신기능

| 기능 | 설명 |
|------|------|
| 🔗 URL 입력 | 뉴스/분석 사이트 URL → 텍스트 추출 → RAG 반영 |
| 📎 파일 업로드 | Excel, PDF, Word, PPT, TXT, HWP, 이미지 → 텍스트 추출 → RAG 반영 |
| 🧠 RAG 분석 | TF-IDF 기반 인메모리 벡터 검색, 관련 컨텍스트 자동 주입 |
| 👤 5가지 페르소나 | A~E 전문가 관점 선택 (전문/기관/외인/보수/공격적) |
| 📄 Word 내보내기 | 전체 리포트 → 전문 서식 .docx |
| 📊 PPT 내보내기 | 요약본 → 다크 프리미엄 디자인 .pptx (차트 포함) |

---

## 🚀 빠른 시작

```bash
# 1. 압축 해제
unzip stock_briefing_v2.zip && cd stock_v2

# 2. Python 패키지 설치
pip install -r requirements.txt

# 3. Node.js 패키지 설치 (Word/PPT 내보내기용)
npm install -g docx pptxgenjs

# 4. 환경변수 설정
cp .env.example .env
# HF_TOKEN, GMAIL_USER, GMAIL_APP_PASSWORD 입력

# 5. 실행
streamlit run app.py
```

---

## 🗂️ 파일 구조

```
stock_v2/
├── app.py                    ← Streamlit 메인 (6탭 UI)
├── config.py                 ← 환경변수·모델 설정
├── stock_data.py             ← pykrx+yfinance 주가 수집 (종목명 지원)
├── news_fetcher.py           ← 뉴스 수집+번역
├── llm_engine.py             ← HF Router API 연동
├── report_generator.py       ← 페르소나+RAG 통합 리포트
├── analyst_personas.py       ← A~E 5가지 분석가 페르소나
├── email_sender.py           ← Gmail SMTP 발송
├── requirements.txt
├── .env.example
├── uploaders/
│   ├── file_reader.py        ← Excel/PDF/Word/PPT/TXT/HWP/이미지 읽기
│   └── url_reader.py         ← URL 텍스트 추출
├── rag/
│   └── rag_engine.py         ← TF-IDF 인메모리 RAG 엔진
└── exporters/
    ├── word_exporter.py      ← .docx 내보내기 (Node.js docx)
    └── ppt_exporter.py       ← .pptx 내보내기 (Node.js pptxgenjs)
```

---

## 👤 5가지 AI 분석가 페르소나

| 키 | 이름 | 특징 |
|----|------|------|
| A 🏆 | 20년 경험 전문 트레이더 | 기술적+기본적 통합, 진입/청산가 중심 |
| B 🏦 | 기관투자 트레이더 | 수급·밸류에이션·매크로 중심 |
| C 🌏 | 외국인투자 트레이더 | 환율·글로벌 피어·MSCI 중심 |
| D 🛡️ | 보수적 안정성 트레이더 | 원금보전·배당·저변동성 |
| E 🚀 | 적극적 위험감수 트레이더 | 모멘텀·테마주·단기 급등 |

---

## 📎 지원 파일 형식

| 형식 | 확장자 | 비고 |
|------|--------|------|
| Excel | .xlsx, .xls | 시트별 전체 추출 |
| PDF | .pdf | pdfplumber (최대 30페이지) |
| Word | .docx, .doc | 본문+표 추출 |
| PowerPoint | .pptx, .ppt | 슬라이드별 텍스트 |
| 텍스트 | .txt, .md, .csv | UTF-8/EUC-KR 자동 감지 |
| 한글 | .hwp | olefile 기반 (제한적) |
| 이미지 | .png, .jpg | OCR (pytesseract 설치 시) |

---

## ⚙️ 모델 목록

```
google/gemma-4-26B-A4B-it:deepinfra  ← 기본값
google/gemma-4-26B-A4B-it:novita
google/gemma-4-31B-it:deepinfra
google/gemma-4-31B-it:together
Qwen/Qwen3.5-9B:together
Qwen/Qwen2.5-7B-Instruct:together
```

---

## 📝 주의사항

- Word/PPT 내보내기는 **Node.js** 필수: `npm install -g docx pptxgenjs`
- HWP 파일은 `pip install olefile` 필요 (완전하지 않을 수 있음)
- 이미지 OCR은 `pip install pytesseract` + Tesseract-OCR 설치 필요
- RAG 인덱스는 Streamlit 세션 내에서만 유지 (재시작 시 초기화)
- 본 리포트는 AI 분석 기반이며 **투자 결정은 본인 책임**입니다

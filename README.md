# 📑 DART 자동 다운로더 (Streamlit, Persist 버전)

**OpenDART API**를 활용해 기업의 공시 ZIP 파일을 자동으로 수집하고,  
요약 엑셀·CSV 파일 및 묶음 ZIP을 한 번에 다운로드할 수 있는 Streamlit 앱입니다.  
공시정보를 일일이 찾을 필요 없이, **클릭 한 번으로 연도별 공시 ZIP 전체를 내려받을 수 있습니다.**

---

## 🌐 실행 데모

🔗 [https://sjhvske89mknqc52qn9tmk.streamlit.app](https://sjhvske89mknqc52qn9tmk.streamlit.app)

> ⚠️ 실제 동작에는 개인 OpenDART API Key가 필요합니다.

---

## 🚀 실행 예시

- **메인화면:**  
  [![메인화면](https://drive.google.com/uc?export=view&id=1ca7wZRwcdSepcrxczrdi8d3Ec-6g6uTG)](https://drive.google.com/uc?export=view&id=1ca7wZRwcdSepcrxczrdi8d3Ec-6g6uTG)

- **결과표:**  
  [![결과표](https://drive.google.com/uc?export=view&id=1BQ8VL6PBISbpNbmrSaCsA6nk-fEBx2oC)](https://drive.google.com/uc?export=view&id=1BQ8VL6PBISbpNbmrSaCsA6nk-fEBx2oC)

> 예시: **삼성전자서비스 (2023)** 공시 10건 자동 다운로드 결과  
> 회사 보안정책상 예시 이미지는 Google Drive를 통해 공유되었습니다.

---

## ⚙️ 주요 기능

### 🔍 회사 검색
- 정확일치 + 부분일치 동시 검색  
- 상장사(`stock_code` 존재) 우선 정렬  
- 동일명 기업(예: 삼성전자 / 삼성전자서비스) 구분 가능  
- 검색 결과 개수 자동 표시  
- 첫 항목 자동 선택 방지 (`index=None`)  

### 🧾 공시 ZIP 자동 다운로드
- 선택한 기업의 **해당 연도 전체 공시 ZIP 자동 수집**
- 페이지네이션(`page_no`) 자동 처리  
- 각 ZIP 파일명:  
  ```
  제출일_보고서명_접수번호.zip
  ```
- 모든 ZIP을 하나로 묶은 **일괄 ZIP 다운로드** 제공  
- 각 항목에는 **DART 원문 바로가기 링크 포함**

### 📊 요약 파일 자동 생성
- 수집된 공시 목록을 엑셀(`.xlsx`) 및 CSV(`.csv`)로 제공  
- 주요 컬럼:  
  - 제출일, 보고서명, 접수번호, 공시제목, DART 링크, ZIP 파일명 등  
- Streamlit 내에서 즉시 다운로드 가능

### 🧠 결과 유지 기능 (Persist 모드)
- Streamlit의 rerun(재실행) 후에도 결과가 유지됨  
- 세션 상태(`st.session_state`)에 다음 항목 저장:
  ```python
  ["last_df", "excel_bytes", "csv_bytes", "zip_bundle_bytes",
   "last_corp_name", "last_corp_code", "last_year", "last_count"]
  ```
- rerun이 발생해도 표, 다운로드 버튼, 기업정보가 그대로 유지되어  
  파일을 여러 번 내려받거나 비교할 때 매우 편리

### 🧹 유틸리티
- **캐시 비우기 버튼**: `@st.cache_data.clear()`로 즉시 초기화  
- **검색 실패 시**: 마지막 정상 마스터(`corp_master_cache`) 자동 재사용  
- **에러 대응 강화**:
  - ZIP이 아닌 HTML/JSON/XML 응답일 경우 상세 원인 안내  
  - API Key 오류 또는 호출 한도 초과 시 명확한 메시지 표시  

---

## 🧩 파일 구조

```
dart_auto_downloader/
├── app.py                # Streamlit 메인 앱
├── requirements.txt      # 필요 패키지 목록
├── README.md             # 설명서 (현재 파일)
└── assets/               # (선택) 이미지 또는 로고
```

---

## 🧰 설치 및 실행

### 1️⃣ 패키지 설치
```bash
pip install streamlit pandas requests openpyxl
```

### 2️⃣ 앱 실행
```bash
streamlit run app.py
```

### 3️⃣ API Key 설정
- [OpenDART](https://opendart.fss.or.kr/) 회원가입 후 **인증키 발급**
- 실행 후 좌측 사이드바에 API Key 입력

---

## 💡 사용 방법

1. **API Key 입력**
2. **검색어 입력 → [회사 검색] 클릭**
3. 검색 결과 중 원하는 기업 선택
4. **연도 입력 → [공시 ZIP 다운로드] 클릭**
5. 다운로드 완료 후:
   - 📊 요약 엑셀 받기  
   - 📄 요약 CSV 받기  
   - 🧾 ZIP 일괄 다운로드  

---

## 🧱 내부 구조 개요

| 구성요소 | 설명 |
|-----------|------|
| `fetch_corp_master()` | 법인코드 마스터 ZIP(XML) 다운로드 및 캐싱 |
| `search_companies()` | 회사명 검색 (정확일치 + 부분일치) |
| `fetch_list()` | 연도별 공시 목록 수집 (페이지네이션 자동) |
| `download_zip_bytes()` | 개별 공시 ZIP 다운로드 |
| `make_excel_bytes()` | DataFrame → 엑셀 바이트 변환 |
| `render_results()` | 세션 결과 표 및 다운로드 버튼 렌더링 |
| `clear_results()` | 세션 상태 초기화 |

---

## 🧩 변경 요약 (vs 안정화 버전)

| 구분 | 안정화 버전 | Persist 버전 |
|------|--------------|---------------|
| 세션 유지 | ❌ 새로고침 시 초기화 | ✅ rerun 후에도 결과 유지 |
| UI | 단일 실행 중심 | 이전 결과 자동 렌더링 |
| 캐시 관리 | 수동 초기화 불편 | 사이드바에서 즉시 초기화 가능 |
| 오류 처리 | ZIP 여부만 확인 | HTML/JSON/XML 상세 메시지 추가 |
| 코드 품질 | 일부 중복 존재 | 정리 및 버전 헤더 업데이트 (v1.4) |

---

## ⚠️ 주의사항

- OpenDART API 호출 한도: **1일 약 10,000회**
- 회사명 검색 시 대소문자·공백 무시됨  
- 네트워크 오류 시 잠시 대기 후 재시도  
- 대량 다운로드 시 API 예의상 약간의 대기(`sleep 0.06~0.08s`)

---

## 🧑‍💻 라이선스

MIT License  
Copyright (c) 2025

---

## 📚 참고 자료

- OpenDART API 문서: [https://opendart.fss.or.kr/guide/main.do](https://opendart.fss.or.kr/guide/main.do)
- Streamlit 공식 문서: [https://docs.streamlit.io](https://docs.streamlit.io)

# 📑 DART 자동 다운로더 (Streamlit)

**OpenDART API**를 이용해 기업의 공시 ZIP 파일을 자동으로 수집하고,  
요약 엑셀·CSV 파일과 함께 한 번에 다운로드할 수 있는 Streamlit 웹앱입니다.  
공시정보를 매번 수동으로 찾는 대신 클릭 한 번으로 ZIP 묶음을 내려받을 수 있습니다.

---

## 🌐 실행 데모

🔗 [https://sjhvske89mkq52qn9tmk.streamlit.app](https://sjhvske89mkq52qn9tmk.streamlit.app)

---

## 🚀 실행 결과 예시

메인화면 : https://drive.google.com/uc?export=view&id=1ca7wZRwcdSepcrxczrdi8d3Ec-6g6uTG,
결과표 : https://drive.google.com/uc?export=view&id=1BQ8VL6PBISbpNbmrSaCsA6nk-fEBx2oC

> 예시: **삼성전자서비스 (2023)** 검색 후 10건의 공시 ZIP 자동 다운로드 결과입니다.  
> 회사 보안정책상 이미지 파일은 Google Drive를 통해 공유되었습니다.

---

## ⚙️ 주요 기능

### 🔍 회사 검색
- **정확일치 + 부분일치** 동시 지원  
- 상장사 우선 정렬  
- 동일명 기업(예: 삼성전자 / 삼성전자서비스)도 구분 선택 가능  

### 🧾 공시 ZIP 자동 다운로드
- 선택한 기업의 해당 연도 전체 공시 ZIP 자동 수집  
- ZIP 파일명: `제출일_보고서명_접수번호.zip`  
- 각 항목에 **DART 원문 링크 자동 포함**

### 📊 요약 파일 생성
- 엑셀/CSV 형식으로 정리된 요약표 제공  
- 기업명, 보고서명, 접수번호, 제출일, ZIP 파일명, DART 링크 등 포함  
- Streamlit 내에서 즉시 다운로드 가능  

### 🧠 결과 유지 기능
- Streamlit의 rerun(재실행)에도 결과 테이블과 버튼이 유지되어  
  파일을 여러 번 다운로드하거나 비교하기 편리  

---

## 🧩 파일 구조


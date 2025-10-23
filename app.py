import io
import re
import time
import zipfile
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional

import streamlit as st

# ====== 기존 모듈 상수 ======
API_HOST = "https://opendart.fss.or.kr"
CORPCODE_API = f"{API_HOST}/api/corpCode.xml"
LIST_API    = f"{API_HOST}/api/list.json"
DOC_API     = f"{API_HOST}/api/document.xml"

S = requests.Session()
S.headers.update({"User-Agent": "dart-auto-downloader/streamlit/1.0"})

# ====== 유틸 ======
def sanitize_filename(name: str) -> str:
    if not name:
        name = "unknown_report"
    bad = r'\\/:*?"<>|'
    for ch in bad:
        name = name.replace(ch, "_")
    name = "_".join(name.split())
    return name[:120]

@st.cache_data(show_spinner=False)
def fetch_corp_master(api_key: str) -> pd.DataFrame:
    """법인코드 마스터 다운로드 + 파싱 (캐시)"""
    r = S.get(CORPCODE_API, params={"crtfc_key": api_key}, timeout=60)
    r.raise_for_status()
    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        # CORPCODE.xml 하나뿐
        with zf.open(zf.namelist()[0]) as fp:
            import xml.etree.ElementTree as ET
            root = ET.parse(fp).getroot()
    rows = []
    for el in root.findall(".//list"):
        rows.append({
            "corp_code": el.findtext("corp_code") or "",
            "corp_name": el.findtext("corp_name") or "",
            "stock_code": el.findtext("stock_code") or "",
        })
    return pd.DataFrame(rows)

def search_companies(master: pd.DataFrame, query: str) -> pd.DataFrame:
    """회사명 검색 (정확일치 + 부분일치 동시 표출)"""
    q = (query or "").strip()
    if not q:
        return master.head(0)
    m = master.copy()
    m["__norm"] = m["corp_name"].fillna("").str.replace(r"\s+", "", regex=True)
    qn = re.sub(r"\s+", "", q)
    mask_exact = m["__norm"].str.casefold() == qn.casefold()
    exact = m[mask_exact].copy()
    exact["__rank"] = 0
    mask_part = m["__norm"].str.contains(re.escape(qn), case=False, regex=True)
    part = m[mask_part & (~mask_exact)].copy()
    part["__rank"] = 1
    res = pd.concat([exact, part], ignore_index=True)
    res["__listed"] = res["stock_code"].fillna("").ne("")
    res = res.sort_values(by=["__rank", "__listed", "corp_name"], ascending=[True, False, True])
    return res.head(200).drop(columns=["__norm","__rank","__listed"], errors="ignore")

def fetch_list(api_key: str, corp_code: str, year: str) -> List[Dict]:
    out = []
    page_no = 1
    bgn_de = f"{year}0101"
    end_de = f"{year}1231"
    while True:
        params = {
            "crtfc_key": api_key,
            "corp_code": corp_code,
            "bgn_de": bgn_de,
            "end_de": end_de,
            "page_no": page_no,
            "page_count": 100
        }
        r = S.get(LIST_API, params=params, timeout=60)
        r.raise_for_status()
        data = r.json()
        if data.get("status") != "000":
            raise RuntimeError(f"list.json 오류: {data.get('message')}")
        items = data.get("list") or []
        out.extend(items)
        total = int(data.get("total_count", 0))
        if len(out) >= total or not items:
            break
        page_no += 1
        time.sleep(0.1)
    return out

def is_zip(content: bytes) -> bool:
    return len(content) > 1000 and content[:4] == b"PK\x03\x04"

def download_zip(api_key: str, rcept_no: str, rcept_dt: str, report_nm: str) -> Optional[bytes]:
    """ZIP 다운로드 (파일명은 호출부에서 결정, 여기서는 바이트만 반환)"""
    params = {"crtfc_key": api_key, "rcept_no": rcept_no}
    r = S.get(DOC_API, params=params, timeout=60)
    content = r.content or b""
    if r.status_code == 200 and is_zip(content):
        return content
    return None

def make_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf.read()

# ====== Streamlit UI ======
st.set_page_config(page_title="DART 자동 다운로더 (Streamlit)", page_icon="📑", layout="wide")
st.title("📑 DART 자동 다운로더 (Streamlit)")

with st.sidebar:
    st.markdown("### 설정")
    api_key = st.text_input("OpenDART API Key", type="password", help="https://opendart.fss.or.kr/")
    default_year = datetime.now().year
    year = st.text_input("다운로드 연도 (YYYY)", value=str(default_year))
    query = st.text_input("회사명 검색 (부분일치 가능)", value="")
    run_search = st.button("회사 검색")

col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown("#### 검색 및 선택")
    if api_key and run_search:
        try:
            master = fetch_corp_master(api_key)
            cand = search_companies(master, query)
            if cand.empty:
                st.warning("검색 결과가 없습니다.")
            else:
                # 사용자 선택 박스
                # 표시는 '회사명 (corp_code / 주식코드)' 형태
                cand["_label"] = cand.apply(
                    lambda r: f"{r['corp_name']} (corp_code:{r['corp_code']}" +
                              (f", 주식코드:{r['stock_code']}" if r['stock_code'] else "") + ")",
                    axis=1
                )
                idx = st.selectbox("회사 선택", range(len(cand)), format_func=lambda i: cand.iloc[i]["_label"])
                # 선택 상태 공유
                st.session_state["selected_company"] = {
                    "corp_code": cand.iloc[idx]["corp_code"],
                    "corp_name": cand.iloc[idx]["corp_name"],
                }
        except Exception as e:
            st.error(f"회사 검색 오류: {e}")

    # 실행 버튼
    run_download = st.button("공시 ZIP 다운로드 & 요약 생성")

with col_right:
    st.markdown("#### 결과")
    placeholder_table = st.empty()
    placeholder_dl = st.empty()

# ====== 액션 처리 ======
if run_download:
    if not api_key:
        st.error("API Key를 입력하세요.")
    elif "selected_company" not in st.session_state:
        st.error("회사를 먼저 검색/선택하세요.")
    elif not (len(year) == 4 and year.isdigit()):
        st.error("연도 형식이 올바르지 않습니다. 예: 2024")
    else:
        corp_code = st.session_state["selected_company"]["corp_code"]
        corp_name = st.session_state["selected_company"]["corp_name"]
        try:
            with st.spinner("공시 목록 수집 중…"):
                items = fetch_list(api_key, corp_code, year)
            if not items:
                st.info("해당 연도에 수집할 공시가 없습니다.")
            else:
                summary = []
                # 다운로드 파일을 하나의 zip(묶음)으로 제공하기 위해 메모리 ZIP 빌더
                bundle_buf = io.BytesIO()
                with zipfile.ZipFile(bundle_buf, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
                    for idx, it in enumerate(items, start=1):
                        rcept_no  = it.get("rcept_no", "")
                        rcept_dt  = it.get("rcept_dt", "")
                        report_nm = it.get("report_nm") or it.get("rpt_nm") or "unknown_report"

                        safe_zip_name = sanitize_filename(f"{rcept_dt}_{report_nm}_{rcept_no}") + ".zip"
                        # 이미 존재 여부는 클라우드에선 의미가 적으므로 건너뜀
                        content = download_zip(api_key, rcept_no, rcept_dt, report_nm)
                        if content:
                            bundle.writestr(safe_zip_name, content)
                            zip_saved = safe_zip_name
                        else:
                            zip_saved = f"{rcept_no}.zip (다운로드 실패)"

                        summary.append({
                            "기업명": corp_name,
                            "corp_code": corp_code,
                            "보고서명": report_nm,
                            "접수번호": rcept_no,
                            "제출일": rcept_dt,
                            "ZIP저장파일": zip_saved,
                            "DART링크": f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}"
                        })
                        # 서버 과부하 방지
                        time.sleep(0.1)

                df = pd.DataFrame(summary).sort_values(["제출일", "보고서명"], ascending=[False, True])

                # 표 렌더
                with placeholder_table:
                    st.dataframe(df, use_container_width=True)

                # 다운로드 버튼들(엑셀/CSV/ZIP묶음)
                excel_bytes = make_excel(df)
                csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
                bundle_buf.seek(0)
                zip_bundle_bytes = bundle_buf.read()

                with placeholder_dl:
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.download_button(
                            "📊 요약 엑셀 받기",
                            data=excel_bytes,
                            file_name=f"공시ZIP요약_{year}_{sanitize_filename(corp_name)}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with c2:
                        st.download_button(
                            "📄 요약 CSV 받기",
                            data=csv_bytes,
                            file_name=f"공시ZIP요약_{year}_{sanitize_filename(corp_name)}.csv",
                            mime="text/csv"
                        )
                    with c3:
                        st.download_button(
                            "🧾 ZIP 일괄 다운로드",
                            data=zip_bundle_bytes,
                            file_name=f"DART_{year}_{sanitize_filename(corp_name)}_{corp_code}_ZIP묶음.zip",
                            mime="application/zip"
                        )
                st.success(f"완료! 총 {len(df)}건 처리했습니다.")
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

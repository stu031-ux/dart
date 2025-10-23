# -*- coding: utf-8 -*-
"""
DART 자동 다운로더 (Streamlit, 안정화 버전)
- 회사 검색(정확일치 우선 + 부분일치)
- 연도별 공시 목록 수집
- 각 공시 ZIP 다운로드 → 묶음 ZIP으로 제공
- ZIP 파일명: 제출일_보고서명_접수번호.zip
- 요약(엑셀/CSV) + DART 바로가기 링크 포함
- 안정화:
  * corpCode.xml이 ZIP이 아닐 때 원인 메시지 노출
  * selectbox에 '레코드(dict) 자체' 사용 + index=None (첫 항목 자동선택 방지)
  * 새 검색 시 이전 선택값 초기화
"""

import os
import io
import re
import time
import zipfile
import requests
import pandas as pd
from datetime import datetime
from typing import List, Dict, Optional

import streamlit as st

# ==============================
# Streamlit Page Config
# ==============================
st.set_page_config(page_title="DART 자동 다운로더", page_icon="📑", layout="wide")
st.title("📑 DART 자동 다운로더")

# ==============================
# OpenDART Endpoints / Session
# ==============================
API_HOST = "https://opendart.fss.or.kr"
CORPCODE_API = f"{API_HOST}/api/corpCode.xml"
LIST_API    = f"{API_HOST}/api/list.json"
DOC_API     = f"{API_HOST}/api/document.xml"

S = requests.Session()
S.headers.update({"User-Agent": "dart-auto-downloader/streamlit/1.2"})

# ==============================
# Utils
# ==============================
def sanitize_filename(name: str) -> str:
    """Windows/macOS/Linux 공통 안전 파일명으로 정리"""
    if not name:
        name = "unknown_report"
    bad = r'\\/:*?"<>|'
    for ch in bad:
        name = name.replace(ch, "_")
    name = "_".join(name.split())
    return name[:120]

def is_zip(content: bytes) -> bool:
    return len(content) > 4 and content[:4] == b"PK\x03\x04"

@st.cache_data(show_spinner=False)
def fetch_corp_master(api_key: str) -> pd.DataFrame:
    """법인코드 마스터 ZIP(xml) 다운로드 후 DataFrame으로 반환 (캐시)
       - 응답이 ZIP이 아닐 경우, 에러 메시지를 추출해 안내
    """
    key = (api_key or "").strip()
    if not key:
        raise RuntimeError("API Key가 비어 있습니다. 올바른 키를 입력하세요.")

    try:
        r = S.get(CORPCODE_API, params={"crtfc_key": key}, timeout=60)
    except requests.RequestException as e:
        raise RuntimeError(f"네트워크 오류: {e}")

    content = r.content or b""
    content_type = r.headers.get("Content-Type", "")
    if not (is_zip(content) or "zip" in content_type.lower()):
        # ZIP이 아니면 JSON/XML/HTML일 수 있음 → 가능한 메시지 추출
        txt = ""
        try:
            txt = content.decode("utf-8", errors="ignore")
        except Exception:
            pass

        # JSON(status/message)
        try:
            import json
            j = json.loads(txt)
            status = j.get("status"); message = j.get("message")
            if status or message:
                raise RuntimeError(f"OpenDART 오류(status={status}): {message}")
        except Exception:
            pass

        # XML(<message>..</message>)
        try:
            import xml.etree.ElementTree as ET
            root = ET.fromstring(txt)
            msg = root.findtext(".//message") or root.findtext(".//msg") or ""
            if msg.strip():
                raise RuntimeError(f"OpenDART 오류: {msg.strip()}")
        except Exception:
            pass

        hint = f" (HTTP {r.status_code})" if r.status_code and r.status_code != 200 else ""
        if "html" in content_type.lower():
            raise RuntimeError(f"OpenDART에서 ZIP이 아닌 HTML 응답을 반환했습니다{hint}. 잠시 후 다시 시도하거나 API Key/한도를 확인하세요.")
        raise RuntimeError("OpenDART에서 ZIP이 아닌 응답을 반환했습니다. API Key/요청 상태를 확인하세요.")

    # 정상 ZIP 처리
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            with zf.open(zf.namelist()[0]) as fp:
                import xml.etree.ElementTree as ET
                root = ET.parse(fp).getroot()
    except zipfile.BadZipFile:
        raise RuntimeError("받은 파일이 ZIP 형식이 아닙니다. 잠시 후 다시 시도하거나 API Key를 확인하세요.")
    except Exception as e:
        raise RuntimeError(f"ZIP 처리 중 오류: {e}")

    rows = []
    for el in root.findall(".//list"):
        rows.append({
            "corp_code": el.findtext("corp_code") or "",
            "corp_name": el.findtext("corp_name") or "",
            "stock_code": el.findtext("stock_code") or "",
        })
    return pd.DataFrame(rows)

def search_companies(master: pd.DataFrame, query: str) -> pd.DataFrame:
    """회사명 검색 (정확일치 + 부분일치 동시 표출, 정확일치 우선)"""
    q = (query or "").strip()
    if not q:
        return master.head(0)

    m = master.copy()
    m["__norm"] = m["corp_name"].fillna("").str.replace(r"\s+", "", regex=True)
    qn = re.sub(r"\s+", "", q)

    # 정확일치
    mask_exact = m["__norm"].str.casefold() == qn.casefold()
    exact = m[mask_exact].copy()
    exact["__rank"] = 0

    # 부분일치
    mask_part = m["__norm"].str.contains(re.escape(qn), case=False, regex=True)
    part = m[mask_part & (~mask_exact)].copy()
    part["__rank"] = 1

    # 합치고 정렬 (상장사 우선)
    res = pd.concat([exact, part], ignore_index=True)
    res["__listed"] = res["stock_code"].fillna("").ne("")
    res = res.sort_values(by=["__rank", "__listed", "corp_name"], ascending=[True, False, True])
    return res.head(200).drop(columns=["__norm", "__rank", "__listed"], errors="ignore")

def fetch_list(api_key: str, corp_code: str, year: str) -> List[Dict]:
    """연도별 공시 목록 수집(list.json 페이지네이션 처리)"""
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
            "page_count": 100,
        }
        r = S.get(LIST_API, params=params, timeout=60)
        r.raise_for_status()
        data = r.json()
        if data.get("status") != "000":
            raise RuntimeError(f"list.json 오류: {data.get('message')}")
        items = data.get("list") or []
        out.extend(items)
        total = int((data.get("total_count") or 0))
        if len(out) >= total or not items:
            break
        page_no += 1
        time.sleep(0.08)  # API 예의상 약간 대기
    return out

def download_zip_bytes(api_key: str, rcept_no: str) -> Optional[bytes]:
    """document.xml ZIP 원문(바이트) 반환"""
    params = {"crtfc_key": api_key, "rcept_no": rcept_no}
    r = S.get(DOC_API, params=params, timeout=60)
    content = r.content or b""
    if r.status_code == 200 and is_zip(content):
        return content
    return None

def make_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    return buf.read()

# ==============================
# Sidebar (입력)
# ==============================
st.sidebar.markdown("### ⚙️ 설정")
api_key = st.sidebar.text_input(
    "OpenDART API Key",
    type="password",
    value=os.getenv("OPENDART_API_KEY", ""),
    help="https://opendart.fss.or.kr/ 에서 키를 발급받아 입력하세요.",
).strip()

year_default = str(datetime.now().year)
year = st.sidebar.text_input("다운로드 연도 (YYYY)", value=year_default)
query = st.sidebar.text_input("회사명 검색 (부분일치 가능)", value="")
exact_only = st.sidebar.checkbox("정확히 일치한 회사만 보기", value=False)
run_search = st.sidebar.button("회사 검색")

# =======

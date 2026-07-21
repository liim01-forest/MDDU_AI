from __future__ import annotations

import tempfile
import contextlib
import importlib.util
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from io import StringIO
from pathlib import Path
from typing import Any, Callable
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import requests
import streamlit as st

from emedi_downloader import (
    COMPANY_COLUMNS,
    ITEM_COLUMNS,
    SearchResult,
    append_industry_sheets,
    build_session,
    collect_all_rows,
    default_company_params,
    default_item_params,
    download_excel,
    safe_filename,
    search_page,
)


APP_TITLE = "의료기기 인허가 정보 수집기"

FDA_510K_URL = "https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfPMN/pmn.cfm"
FDA_PMA_URL = "https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfPMA/pma.cfm"
FDA_TPLC_URL = "https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfTPLC/tplc.cfm"
FDA_MAUDE_URL = "https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfMAUDE/TextSearch.cfm"
OPENFDA_510K_URL = "https://api.fda.gov/device/510k.json"
OPENFDA_PMA_URL = "https://api.fda.gov/device/pma.json"
OPENFDA_CLASSIFICATION_URL = "https://api.fda.gov/device/classification.json"
OPENFDA_MAUDE_URL = "https://api.fda.gov/device/event.json"


STANDARD_COLUMNS = [
    "source",
    "country_or_region",
    "data_type",
    "product_name",
    "manufacturer",
    "approval_or_submission_no",
    "product_code_or_udi",
    "device_class",
    "approval_date",
    "status",
    "raw_category",
    "source_url",
    "collected_at",
]

FDA_RESULT_COLUMNS = [
    "인증 여부",
    "Device Name",
    "Product Code",
    "Regulation Description",
    "Subsequent Product Code",
    "Class",
    "510(k) Number",
    "Decision Date",
    "원본 사이트 링크",
]


@dataclass(frozen=True)
class SearchFilters:
    keyword: str
    manufacturer: str
    approval_no: str
    start_date: date | None
    end_date: date | None
    max_pages: int
    request_delay: float


def collected_at() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sample_mfds(filters: SearchFilters) -> pd.DataFrame:
    keyword = filters.keyword or "혈압계"
    manufacturer = filters.manufacturer or "메디케어"
    now = collected_at()
    rows = [
        {
            "source": "MFDS",
            "country_or_region": "KR",
            "data_type": "품목허가",
            "product_name": f"{keyword} 자동 측정 시스템",
            "manufacturer": manufacturer,
            "approval_or_submission_no": filters.approval_no or "제허 24-0001호",
            "product_code_or_udi": "A23010.01",
            "device_class": "2등급",
            "approval_date": "2024-03-18",
            "status": "허가",
            "raw_category": "의료기기 품목정보",
            "source_url": "https://emed.mfds.go.kr",
            "collected_at": now,
        },
        {
            "source": "MFDS",
            "country_or_region": "KR",
            "data_type": "업체정보",
            "product_name": f"{keyword} 모니터링 소프트웨어",
            "manufacturer": manufacturer,
            "approval_or_submission_no": "제인 23-0420호",
            "product_code_or_udi": "E06020.01",
            "device_class": "2등급",
            "approval_date": "2023-11-02",
            "status": "인증",
            "raw_category": "업체/제품 검색",
            "source_url": "https://emed.mfds.go.kr",
            "collected_at": now,
        },
    ]
    return pd.DataFrame(rows, columns=STANDARD_COLUMNS)


def normalize_condition(values: dict[str, Any], preserve_empty_keys: set[str] | None = None) -> dict[str, str]:
    preserve_empty_keys = preserve_empty_keys or set()
    result: dict[str, str] = {}
    for key, value in values.items():
        text = "" if value is None else str(value).strip()
        if text or key in preserve_empty_keys:
            result[key] = text
    return result


def has_real_condition(condition: dict[str, str]) -> bool:
    return any(key != "chkGroup" and bool(value) for key, value in condition.items())


def mfds_yes_no(key: str) -> str:
    return st.radio(
        "여부",
        ["", "Y", "N"],
        index=0,
        format_func=lambda value: {"": "전체", "Y": "예", "N": "아니오"}[value],
        horizontal=True,
        key=key,
        label_visibility="collapsed",
    )


def mfds_industry_select(key: str) -> str:
    return st.selectbox(
        "업종",
        ["1|2|21|22", "1", "2", "21", "22"],
        format_func=lambda value: {
            "1|2|21|22": "전체",
            "1": "제조업",
            "2": "수입업",
            "21": "체외진단제조업",
            "22": "체외진단수입업",
        }[value],
        key=key,
        label_visibility="collapsed",
    )


def mfds_business_state_select(key: str) -> str:
    return st.selectbox(
        "업상태",
        ["", "0", "1", "2", "3", "4"],
        format_func=lambda value: {"": "전체", "0": "정상", "1": "폐업", "2": "휴업", "3": "재개", "4": "취소"}[value],
        key=key,
        label_visibility="collapsed",
    )


def mfds_date_range(prefix: str, label: str) -> tuple[date | None, date | None]:
    start_col, end_col = st.columns(2)
    start = start_col.date_input(f"{label} 시작", value=None, key=f"{prefix}_from")
    end = end_col.date_input(f"{label} 종료", value=None, key=f"{prefix}_to")
    return start, end


def render_mfds_item_form() -> dict[str, str]:
    st.markdown("#### 품목검색")
    cols = st.columns(4)
    query2 = cols[0].text_input("명칭", placeholder="제품명, 품목명, 모델명", key="mfds_item_query2")
    udidi_code = cols[1].text_input("UDI 코드", key="mfds_item_udidi")
    grade = cols[2].selectbox("품목등급", ["0", "1", "2", "3", "4"], format_func=lambda value: "전체" if value == "0" else f"{value}등급", key="mfds_item_grade")
    item_state = cols[3].selectbox("품목상태", ["", "정상", "취소", "취하", "양도", "만료"], format_func=lambda value: "전체" if value == "" else value, key="mfds_item_state")

    cols = st.columns(4)
    item_no = cols[0].text_input("품목허가번호", placeholder="예: 제허00-000호", key="mfds_item_no")
    entp_name = cols[1].text_input("업체명", key="mfds_item_entp_name")
    with cols[2]:
        indty_cd = mfds_industry_select("mfds_item_industry")
    with cols[3]:
        tcsbiz_state = mfds_business_state_select("mfds_item_business_state")

    cols = st.columns(4)
    mdentp_prmno = cols[0].text_input("업허가번호", key="mfds_item_mdentp_prmno")
    mnfacr_nm = cols[1].text_input("제조원", key="mfds_item_mnfacr_nm")
    type_name = cols[2].text_input("모델명", key="mfds_item_type_name")
    brand_name = cols[3].text_input("제품명", key="mfds_item_brand_name")

    cols = st.columns(4)
    item_name = cols[0].text_input("품목명", key="mfds_item_name")
    query = cols[1].text_input("문구검색", placeholder="업체명, 제품명, 품목명, 모델명, 제조원", key="mfds_item_query")
    rcprslry_cd = cols[2].text_input("요양급여코드", key="mfds_item_rcprslry_cd")
    md_clsf_no = cols[3].text_input("품목분류번호", key="mfds_item_md_clsf_no")

    cols = st.columns(2)
    with cols[0]:
        prm_from, prm_to = mfds_date_range("mfds_item_prm_date", "품목허가일자")
    with cols[1]:
        valid_from, valid_to = mfds_date_range("mfds_item_valid_date", "품목유효기간")

    cols = st.columns(4)
    with cols[0]:
        st.caption("요양급여대상여부")
        rcprslry_trgt = mfds_yes_no("mfds_item_rcprslry_trgt")
    with cols[1]:
        st.caption("추적관리대상여부")
        trace_manage = mfds_yes_no("mfds_item_trace_manage")
    with cols[2]:
        st.caption("수출용에포함")
        xprtpp_yn = mfds_yes_no("mfds_item_xprtpp_yn")
    with cols[3]:
        st.caption("인체이식용여부")
        hmnbd_yn = mfds_yes_no("mfds_item_hmnbd_yn")

    chk_group = st.radio(
        "조회조건",
        ["GROUP_BY_FIELD_01", ""],
        index=0,
        format_func=lambda value: "허가번호 단위" if value == "GROUP_BY_FIELD_01" else "모델명 단위",
        horizontal=True,
        key="mfds_item_chk_group",
    )

    return normalize_condition(
        {
            "query2": query2,
            "udidiCode": udidi_code,
            "itemName": item_name,
            "itemNoFullname": item_no,
            "grade": "" if grade == "0" else grade,
            "itemState": item_state,
            "entpName": entp_name,
            "mdentpPrmno": mdentp_prmno,
            "typeName": type_name,
            "brandName": brand_name,
            "indtyCd": "" if indty_cd == "1|2|21|22" else indty_cd,
            "tcsbizRsmptSeCd": tcsbiz_state,
            "mnfacrNm": mnfacr_nm,
            "query": query,
            "rcprslryCdInptvl": rcprslry_cd,
            "mdClsfNo": md_clsf_no,
            "prdlPrmDtFrom": prm_from.isoformat() if prm_from else "",
            "prdlPrmDtTo": prm_to.isoformat() if prm_to else "",
            "validDateFrom": valid_from.isoformat() if valid_from else "",
            "validDateTo": valid_to.isoformat() if valid_to else "",
            "rcprslryTrgtYn": rcprslry_trgt,
            "traceManageTargetYn": trace_manage,
            "xprtppYn": xprtpp_yn,
            "hmnbdTspnttyMdYn": hmnbd_yn,
            "chkGroup": chk_group,
        },
        preserve_empty_keys={"chkGroup"},
    )


def render_mfds_company_form() -> dict[str, str]:
    st.markdown("#### 업체검색")
    cols = st.columns(4)
    with cols[0]:
        indty_cd = mfds_industry_select("mfds_company_industry")
    with cols[1]:
        tcsbiz_state = mfds_business_state_select("mfds_company_business_state")
    entp_name = cols[2].text_input("업체명", key="mfds_company_entp_name")
    mdentp_prmno = cols[3].text_input("업허가번호", key="mfds_company_mdentp_prmno")

    cols = st.columns([1, 2, 1])
    rprsv_nm = cols[0].text_input("대표자", key="mfds_company_rprsv_nm")
    with cols[1]:
        date_from, date_to = mfds_date_range("mfds_company_prm_date", "업허가일자")

    return normalize_condition(
        {
            "entpName": entp_name,
            "mdentpPrmno": mdentp_prmno,
            "rprsvNm": rprsv_nm,
            "indtyCd": "" if indty_cd == "1|2|21|22" else indty_cd,
            "tcsbizRsmptSeCd": tcsbiz_state,
            "entpPrmDtFrom": date_from.isoformat() if date_from else "",
            "entpPrmDtTo": date_to.isoformat() if date_to else "",
        }
    )


def mfds_rows_to_table(rows: list[dict[str, Any]], tab: str) -> pd.DataFrame:
    columns = COMPANY_COLUMNS if tab == "company" else ITEM_COLUMNS
    visible = columns[:8] if tab == "company" else columns[:12]
    return pd.DataFrame([{label: row.get(key, "") for key, label in visible} for row in rows])


def mfds_rows_to_standard(rows: list[dict[str, Any]], tab: str) -> pd.DataFrame:
    mapped_rows = []
    now = collected_at()
    for row in rows:
        if tab == "company":
            mapped_rows.append(
                {
                    "source": "MFDS",
                    "country_or_region": "KR",
                    "data_type": "업체정보",
                    "product_name": "",
                    "manufacturer": row.get("ENTP_NAME", ""),
                    "approval_or_submission_no": row.get("MDENTP_PRMNO", ""),
                    "product_code_or_udi": "",
                    "device_class": "",
                    "approval_date": row.get("ENTP_PRM_DT", ""),
                    "status": row.get("TCSBIZ_RSMPT_SE_CD", ""),
                    "raw_category": row.get("INDTY_CD_NM", ""),
                    "source_url": "https://emedi.mfds.go.kr/search/data/list",
                    "collected_at": now,
                }
            )
        else:
            mapped_rows.append(
                {
                    "source": "MFDS",
                    "country_or_region": "KR",
                    "data_type": "품목정보",
                    "product_name": row.get("ITEM_NAME", "") or row.get("BRAND_NAME", ""),
                    "manufacturer": row.get("ENTP_NAME", ""),
                    "approval_or_submission_no": row.get("ITEM_NO_FULLNAME", ""),
                    "product_code_or_udi": row.get("UDIDICD", "") or row.get("MD_CLSF_NO", ""),
                    "device_class": row.get("GRADE", ""),
                    "approval_date": row.get("PRDL_PRM_DT", ""),
                    "status": row.get("ITEM_STATE", ""),
                    "raw_category": row.get("CLSFNO_NM_REGVL", "") or row.get("INDTY_CD_NM", ""),
                    "source_url": "https://emedi.mfds.go.kr/search/data/list",
                    "collected_at": now,
                }
            )
    return pd.DataFrame(mapped_rows, columns=STANDARD_COLUMNS)


def run_mfds_search(tab: str, condition: dict[str, str], page: int = 1) -> None:
    result = search_page(build_session(), tab, condition, page)
    st.session_state.mfds_active_tab = tab
    st.session_state.mfds_condition = condition
    st.session_state.mfds_page = page
    st.session_state.mfds_total = result.total_count
    st.session_state.mfds_raw_rows = result.excel_rows
    st.session_state.mfds_df = mfds_rows_to_standard(result.excel_rows, tab)
    add_log("MFDS", "success", f"eMedi {('업체검색' if tab == 'company' else '품목검색')} {page}페이지 조회", len(result.excel_rows))


def mfds_state_date(key: str) -> str:
    value = st.session_state.get(key)
    return value.isoformat() if hasattr(value, "isoformat") else ""


def current_mfds_item_aligned_condition() -> dict[str, str]:
    indty_cd = st.session_state.get("mfds_item_industry_aligned", "1|2|21|22")
    grade = st.session_state.get("mfds_item_grade_aligned", "0")
    return normalize_condition(
        {
            "query2": st.session_state.get("mfds_item_query2_aligned", ""),
            "udidiCode": st.session_state.get("mfds_item_udidi_aligned", ""),
            "itemName": st.session_state.get("mfds_item_name_aligned", ""),
            "itemNoFullname": st.session_state.get("mfds_item_no_aligned", ""),
            "grade": "" if grade == "0" else grade,
            "itemState": st.session_state.get("mfds_item_state_aligned", ""),
            "entpName": st.session_state.get("mfds_item_entp_name_aligned", ""),
            "mdentpPrmno": st.session_state.get("mfds_item_mdentp_prmno_aligned", ""),
            "typeName": st.session_state.get("mfds_item_type_name_aligned", ""),
            "brandName": st.session_state.get("mfds_item_brand_name_aligned", ""),
            "indtyCd": "" if indty_cd == "1|2|21|22" else indty_cd,
            "tcsbizRsmptSeCd": st.session_state.get("mfds_item_business_state_aligned", ""),
            "mnfacrNm": st.session_state.get("mfds_item_mnfacr_nm_aligned", ""),
            "query": st.session_state.get("mfds_item_query_aligned", ""),
            "rcprslryCdInptvl": st.session_state.get("mfds_item_rcprslry_cd_aligned", ""),
            "mdClsfNo": st.session_state.get("mfds_item_md_clsf_no_aligned", ""),
            "prdlPrmDtFrom": mfds_state_date("mfds_item_prm_from_aligned"),
            "prdlPrmDtTo": mfds_state_date("mfds_item_prm_to_aligned"),
            "validDateFrom": mfds_state_date("mfds_item_valid_from_aligned"),
            "validDateTo": mfds_state_date("mfds_item_valid_to_aligned"),
            "rcprslryTrgtYn": st.session_state.get("mfds_item_rcprslry_trgt_aligned", ""),
            "traceManageTargetYn": st.session_state.get("mfds_item_trace_manage_aligned", ""),
            "xprtppYn": st.session_state.get("mfds_item_xprtpp_yn_aligned", ""),
            "hmnbdTspnttyMdYn": st.session_state.get("mfds_item_hmnbd_yn_aligned", ""),
            "chkGroup": st.session_state.get("mfds_item_chk_group_aligned", "GROUP_BY_FIELD_01"),
        },
        preserve_empty_keys={"chkGroup"},
    )


def current_mfds_company_compact_condition() -> dict[str, str]:
    indty_cd = st.session_state.get("mfds_company_industry_compact", "1|2|21|22")
    return normalize_condition(
        {
            "entpName": st.session_state.get("mfds_company_entp_name_compact", ""),
            "mdentpPrmno": st.session_state.get("mfds_company_mdentp_prmno_compact", ""),
            "rprsvNm": st.session_state.get("mfds_company_rprsv_nm_compact", ""),
            "indtyCd": "" if indty_cd == "1|2|21|22" else indty_cd,
            "tcsbizRsmptSeCd": st.session_state.get("mfds_company_business_state_compact", ""),
            "entpPrmDtFrom": mfds_state_date("mfds_company_prm_from_compact"),
            "entpPrmDtTo": mfds_state_date("mfds_company_prm_to_compact"),
        }
    )


def submit_mfds_enter_search(tab: str, condition_factory: Callable[[], dict[str, str]]) -> None:
    condition = condition_factory()
    if not has_real_condition(condition):
        return
    try:
        run_mfds_search(tab, condition, 1)
        st.session_state.pop(f"mfds_{tab}_enter_error", None)
    except Exception as exc:
        add_log("MFDS", "error", str(exc), 0)
        st.session_state[f"mfds_{tab}_enter_error"] = str(exc)


def make_mfds_download(condition: dict[str, str], tab: str, delay: float, max_pages: int | None) -> tuple[str, bytes, int, int, list[dict[str, Any]]]:
    session = build_session()
    progress_bar = st.progress(0)
    progress_text = st.empty()

    def progress(page: int, total_pages: int, row_count: int) -> None:
        progress_bar.progress(min(page / total_pages, 1.0) if total_pages else 1.0)
        progress_text.info(f"식약처 데이터 수집 중: {page:,}/{total_pages:,}페이지, 누적 {row_count:,}건")

    total_count, rows = collect_all_rows(session, tab, condition, delay=delay, max_pages=max_pages, progress=progress)
    if not rows:
        raise RuntimeError("검색 결과가 없습니다.")

    stem = safe_filename("_".join(f"{key}-{value}" for key, value in condition.items())) or "emedi_result"
    filename = f"{stem}.xlsx"
    params = default_company_params() if tab == "company" else default_item_params()
    params.update(condition)
    result = SearchResult(tab=tab, total_count=total_count, excel_rows=rows, params=params)

    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = Path(temp_dir) / filename
        download_excel(session, result, output_path)
        append_industry_sheets(output_path, rows)
        progress_bar.progress(1.0)
        progress_text.success(f"엑셀 파일 생성 완료: {len(rows):,}/{total_count:,}건")
        return filename, output_path.read_bytes(), total_count, len(rows), rows


def count_unique_mfds_companies(rows: list[dict[str, Any]]) -> dict[str, int]:
    buckets: dict[str, set[tuple[str, str]]] = {"제조업": set(), "수입업": set()}
    for row in rows:
        industry = str(row.get("INDTY_CD_NM") or "").strip()
        if industry not in buckets:
            continue
        stable_id = str(row.get("MDENTP_SNO") or "").strip()
        permit_no = str(row.get("MDENTP_PRMNO") or "").strip()
        name = str(row.get("ENTP_NAME") or "").strip()
        key = ("sno", stable_id) if stable_id else ("permit_name", f"{permit_no}|{name}")
        if key[1].strip("|"):
            buckets[industry].add(key)
    return {industry: len(values) for industry, values in buckets.items()}


def sample_fda(filters: SearchFilters, data_type: str) -> pd.DataFrame:
    keyword = filters.keyword or "Blood Pressure Monitor"
    manufacturer = filters.manufacturer or "MediCare Inc."
    now = collected_at()
    rows = [
        {
            "source": "FDA",
            "country_or_region": "US",
            "data_type": data_type,
            "product_name": keyword,
            "manufacturer": manufacturer,
            "approval_or_submission_no": filters.approval_no or "K241234",
            "product_code_or_udi": "DXN",
            "device_class": "Class II",
            "approval_date": "2024-07-12",
            "status": "Cleared",
            "raw_category": "Premarket Notification",
            "source_url": "https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfpmn/pmn.cfm",
            "collected_at": now,
        },
        {
            "source": "FDA",
            "country_or_region": "US",
            "data_type": "Product Classification",
            "product_name": f"{keyword} accessories",
            "manufacturer": manufacturer,
            "approval_or_submission_no": "21 CFR 870.1130",
            "product_code_or_udi": "DSK",
            "device_class": "Class II",
            "approval_date": "",
            "status": "Active",
            "raw_category": "Product Code Classification",
            "source_url": "https://www.accessdata.fda.gov/scripts/cdrh/cfdocs/cfpcd/classification.cfm",
            "collected_at": now,
        },
    ]
    return pd.DataFrame(rows, columns=STANDARD_COLUMNS)


def fda_escape(value: str) -> str:
    return value.strip().replace('"', '\\"')


def fda_api_get(url: str, params: dict[str, Any]) -> dict[str, Any]:
    response = requests.get(url, params=params, timeout=30)
    if response.status_code == 404:
        return {"results": [], "meta": {"results": {"total": 0}}}
    response.raise_for_status()
    return response.json()


def fda_join_search(parts: list[str]) -> str:
    return " AND ".join(part for part in parts if part)


def build_fda_510k_search(
    product_code: str, device_name: str, k_number: str, start_date: date, end_date: date
) -> str:
    parts = []
    if product_code:
        parts.append(f"product_code:{fda_escape(product_code).upper()}")
    if k_number:
        parts.append(f"k_number:{fda_escape(k_number).upper()}")
    if device_name:
        parts.append(f'device_name:"{fda_escape(device_name)}"')
    parts.append(f"decision_date:[{start_date:%Y%m%d} TO {end_date:%Y%m%d}]")
    return fda_join_search(parts)


def build_fda_pma_search(
    product_code: str, device_name: str, pma_number: str, start_date: date, end_date: date
) -> str:
    parts = []
    if product_code:
        parts.append(f"product_code:{fda_escape(product_code).upper()}")
    if pma_number:
        parts.append(f"pma_number:{fda_escape(pma_number).upper()}")
    if device_name:
        text = fda_escape(device_name)
        parts.append(f'(trade_name:"{text}" OR generic_name:"{text}")')
    parts.append(f"decision_date:[{start_date:%Y%m%d} TO {end_date:%Y%m%d}]")
    return fda_join_search(parts)


def build_fda_classification_search(product_code: str, regulation_number: str) -> str:
    parts = []
    if product_code:
        parts.append(f"product_code:{fda_escape(product_code).upper()}")
    if regulation_number:
        parts.append(f'regulation_number:"{fda_escape(regulation_number)}"')
    return fda_join_search(parts)


def fda_classification_lookup(product_codes: list[str], regulation_number: str = "") -> dict[str, dict[str, Any]]:
    lookup: dict[str, dict[str, Any]] = {}
    codes = [code.strip().upper() for code in product_codes if code and code.strip()]
    if not codes and regulation_number:
        search = build_fda_classification_search("", regulation_number)
        payload = fda_api_get(OPENFDA_CLASSIFICATION_URL, {"search": search, "limit": 20})
        for row in payload.get("results", []):
            code = str(row.get("product_code", "")).upper()
            if code:
                lookup[code] = row
        return lookup

    for code in sorted(set(codes)):
        search = build_fda_classification_search(code, "")
        payload = fda_api_get(OPENFDA_CLASSIFICATION_URL, {"search": search, "limit": 1})
        rows = payload.get("results", [])
        if rows:
            lookup[code] = rows[0]
    return lookup


def fda_certification_status(row: dict[str, Any], source: str) -> str:
    decision = str(row.get("decision_description") or row.get("decision_code") or "").strip()
    if source == "510(k)":
        if "Substantially Equivalent" in decision or row.get("decision_code", "").startswith("SE"):
            return "Cleared"
        return decision or "Unknown"
    if source == "PMA":
        code = str(row.get("decision_code") or "").upper()
        if code == "APPR":
            return "Approved"
        return decision or code or "Unknown"
    return decision or "Unknown"


def fda_source_link(row: dict[str, Any], source: str, product_code: str) -> str:
    if source == "510(k)" and row.get("k_number"):
        return f"{FDA_510K_URL}?ID={row.get('k_number')}"
    if source == "PMA" and row.get("pma_number"):
        return f"{FDA_PMA_URL}?ID={row.get('pma_number')}"
    if product_code:
        return f"{FDA_TPLC_URL}?id={product_code}"
    return FDA_510K_URL


def fda_result_row(row: dict[str, Any], source: str, classification: dict[str, Any] | None) -> dict[str, str]:
    classification = classification or {}
    device_name = row.get("device_name") or row.get("trade_name") or row.get("generic_name") or classification.get("device_name") or ""
    product_code = str(row.get("product_code") or classification.get("product_code") or "").upper()
    regulation_description = classification.get("device_name") or row.get("advisory_committee_description") or ""
    device_class = classification.get("device_class") or ""
    submission_number = row.get("k_number") if source == "510(k)" else row.get("pma_number", "")
    return {
        "인증 여부": fda_certification_status(row, source),
        "Device Name": str(device_name),
        "Product Code": product_code,
        "Regulation Description": str(regulation_description),
        "Subsequent Product Code": "",
        "Class": f"Class {device_class}" if device_class else "",
        "510(k) Number": str(submission_number or ""),
        "Decision Date": str(row.get("decision_date") or ""),
        "원본 사이트 링크": fda_source_link(row, source, product_code),
    }


def fda_results_to_standard(rows: pd.DataFrame) -> pd.DataFrame:
    mapped_rows = []
    now = collected_at()
    for _, row in rows.iterrows():
        mapped_rows.append(
            {
                "source": "FDA",
                "country_or_region": "US",
                "data_type": "510(k)/PMA/TPLC",
                "product_name": row.get("Device Name", ""),
                "manufacturer": "",
                "approval_or_submission_no": row.get("510(k) Number", ""),
                "product_code_or_udi": row.get("Product Code", ""),
                "device_class": row.get("Class", ""),
                "approval_date": row.get("Decision Date", ""),
                "status": row.get("인증 여부", ""),
                "raw_category": row.get("Regulation Description", ""),
                "source_url": FDA_510K_URL,
                "collected_at": now,
            }
        )
    return pd.DataFrame(mapped_rows, columns=STANDARD_COLUMNS)


def search_fda_sources(
    include_510k: bool,
    include_pma: bool,
    include_tplc: bool,
    include_maude: bool,
    product_code: str,
    device_name: str,
    k_number: str,
    pma_number: str,
    regulation_number: str,
    start_date: date,
    end_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, int]]:
    limit = 100
    raw_rows: list[tuple[str, dict[str, Any]]] = []
    source_counts = {"510(k)": 0, "PMA": 0, "TPLC": 0, "MAUDE": 0}

    if include_510k:
        search = build_fda_510k_search(product_code, device_name, k_number, start_date, end_date)
        params = {"limit": limit, "sort": "decision_date:desc"}
        if search:
            params["search"] = search
        payload = fda_api_get(OPENFDA_510K_URL, params)
        for row in payload.get("results", []):
            raw_rows.append(("510(k)", row))
        source_counts["510(k)"] = len(payload.get("results", []))

    if include_pma:
        search = build_fda_pma_search(product_code, device_name, pma_number, start_date, end_date)
        params = {"limit": limit, "sort": "decision_date:desc"}
        if search:
            params["search"] = search
        payload = fda_api_get(OPENFDA_PMA_URL, params)
        for row in payload.get("results", []):
            raw_rows.append(("PMA", row))
        source_counts["PMA"] = len(payload.get("results", []))

    codes = [row.get("product_code", "") for _, row in raw_rows]
    if product_code:
        codes.append(product_code)
    classification_lookup = fda_classification_lookup(codes, regulation_number) if include_tplc else {}
    source_counts["TPLC"] = len(classification_lookup)

    result_rows = []
    for source, row in raw_rows:
        code = str(row.get("product_code") or "").upper()
        result_rows.append(fda_result_row(row, source, classification_lookup.get(code)))

    if include_tplc and not raw_rows:
        for classification in classification_lookup.values():
            result_rows.append(fda_result_row(classification, "TPLC", classification))

    maude_df = pd.DataFrame()
    if include_maude:
        maude_parts = []
        if product_code:
            maude_parts.append(f"device.device_report_product_code:{fda_escape(product_code).upper()}")
        if device_name:
            maude_parts.append(f'device.brand_name:"{fda_escape(device_name)}"')
        maude_parts.append(f"date_received:[{start_date:%Y%m%d} TO {end_date:%Y%m%d}]")
        maude_search = fda_join_search(maude_parts)
        params = {"limit": min(limit, 25)}
        if maude_search:
            params["search"] = maude_search
        maude_payload = fda_api_get(OPENFDA_MAUDE_URL, params)
        maude_rows = maude_payload.get("results", [])
        source_counts["MAUDE"] = len(maude_rows)
        maude_df = pd.DataFrame(
            [
                {
                    "MDR Report Key": row.get("mdr_report_key", ""),
                    "Event Type": row.get("event_type", ""),
                    "Report Date": row.get("date_received", ""),
                    "Source Type": ", ".join(row.get("source_type", []) if isinstance(row.get("source_type"), list) else []),
                }
                for row in maude_rows
            ]
        )

    fda_result_df = pd.DataFrame(result_rows, columns=FDA_RESULT_COLUMNS).drop_duplicates().reset_index(drop=True)
    return fda_result_df, maude_df, source_counts


def sample_eudamed(filters: SearchFilters) -> pd.DataFrame:
    keyword = filters.keyword or "Blood pressure monitor"
    manufacturer = filters.manufacturer or "MediCare Europe GmbH"
    now = collected_at()
    rows = [
        {
            "source": "EUDAMED",
            "country_or_region": "EU",
            "data_type": "UDI/Device",
            "product_name": keyword,
            "manufacturer": manufacturer,
            "approval_or_submission_no": "DE-MF-000012345",
            "product_code_or_udi": "BUDI-DI-1234567890",
            "device_class": "Class IIa",
            "approval_date": "2025-01-20",
            "status": "Registered",
            "raw_category": "Device registration",
            "source_url": "https://ec.europa.eu/tools/eudamed",
            "collected_at": now,
        },
        {
            "source": "EUDAMED",
            "country_or_region": "EU",
            "data_type": "Certificate",
            "product_name": f"{keyword} family",
            "manufacturer": manufacturer,
            "approval_or_submission_no": "MDR-CE-2025-0012",
            "product_code_or_udi": "EMDN Z12030205",
            "device_class": "Class IIa",
            "approval_date": "2025-02-14",
            "status": "Valid",
            "raw_category": "Notified Bodies and Certificates",
            "source_url": "https://ec.europa.eu/tools/eudamed",
            "collected_at": now,
        },
    ]
    return pd.DataFrame(rows, columns=STANDARD_COLUMNS)


def filter_rows(df: pd.DataFrame, filters: SearchFilters) -> pd.DataFrame:
    if df.empty:
        return df

    result = df.copy()
    if filters.keyword:
        mask = result["product_name"].str.contains(filters.keyword, case=False, na=False)
        result = result[mask]
    if filters.manufacturer:
        mask = result["manufacturer"].str.contains(filters.manufacturer, case=False, na=False)
        result = result[mask]
    if filters.approval_no:
        mask = result["approval_or_submission_no"].str.contains(filters.approval_no, case=False, na=False)
        result = result[mask]

    return result.reset_index(drop=True)


def run_collector(name: str, collector: Callable[[], pd.DataFrame]) -> pd.DataFrame:
    with st.spinner(f"{name} 데이터를 조회하는 중입니다..."):
        df = collector()
    st.success(f"{name} 조회 완료: {len(df):,}건")
    return df


def build_excel(
    integrated: pd.DataFrame,
    mfds: pd.DataFrame,
    fda: pd.DataFrame,
    eudamed: pd.DataFrame,
    logs: pd.DataFrame,
    filters: SearchFilters,
) -> bytes:
    sheets = {
        "통합결과": integrated,
        "식약처 원본": mfds,
        "FDA 원본": fda,
        "EUDAMED 원본": eudamed,
        "수집로그": logs,
        "검색조건": pd.DataFrame(
            [
                {"field": "keyword", "value": filters.keyword},
                {"field": "manufacturer", "value": filters.manufacturer},
                {"field": "approval_no", "value": filters.approval_no},
                {"field": "start_date", "value": filters.start_date},
                {"field": "end_date", "value": filters.end_date},
                {"field": "max_pages", "value": filters.max_pages},
                {"field": "request_delay", "value": filters.request_delay},
            ]
        ),
    }

    try:
        return build_excel_with_openpyxl(sheets)
    except ImportError:
        return build_basic_xlsx(sheets)


def build_excel_with_openpyxl(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)

        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 45)

    return output.getvalue()


def build_basic_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    sheet_names = list(sheets.keys())

    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
""" + "".join(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for i in range(1, len(sheet_names) + 1)
            ) + "\n</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
""" + "".join(
                f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                for i in range(1, len(sheet_names) + 1)
            ) + f'<Relationship Id="rId{len(sheet_names) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            + "\n</Relationships>",
        )
        archive.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
""" + "".join(
                f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                for i, name in enumerate(sheet_names, start=1)
            ) + "\n</sheets></workbook>",
        )
        archive.writestr(
            "xl/styles.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>""",
        )

        for index, df in enumerate(sheets.values(), start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", dataframe_to_sheet_xml(df))

    return output.getvalue()


def dataframe_to_sheet_xml(df: pd.DataFrame) -> str:
    rows = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
    xml_rows = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for column_index, value in enumerate(row, start=1):
            ref = f"{column_letter(column_index)}{row_index}"
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>')
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        f'<sheetData>{"".join(xml_rows)}</sheetData>'
        "</worksheet>"
    )


def column_letter(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def ensure_state() -> None:
    defaults = {
        "main_menu": "국가별 인허가 정보",
        "country_menu": "한국",
        "mfds_df": pd.DataFrame(columns=STANDARD_COLUMNS),
        "mfds_raw_rows": [],
        "mfds_active_tab": "item",
        "mfds_condition": {},
        "mfds_page": 1,
        "mfds_total": 0,
        "fda_df": pd.DataFrame(columns=STANDARD_COLUMNS),
        "fda_result_df": pd.DataFrame(columns=FDA_RESULT_COLUMNS),
        "fda_510k_pma_result_df": pd.DataFrame(columns=FDA_RESULT_COLUMNS),
        "fda_tplc_result_df": pd.DataFrame(columns=FDA_RESULT_COLUMNS),
        "fda_maude_df": pd.DataFrame(),
        "fda_510k_pma_counts": {"510(k)": 0, "PMA": 0},
        "fda_tplc_maude_counts": {"TPLC": 0, "MAUDE": 0},
        "fda_source_counts": {"510(k)": 0, "PMA": 0, "TPLC": 0, "MAUDE": 0},
        "eudamed_df": pd.DataFrame(columns=STANDARD_COLUMNS),
        "pms_output_name": "",
        "pms_output_bytes": None,
        "pms_log": "",
        "logs": [],
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def add_log(source: str, status: str, message: str, count: int = 0) -> None:
    st.session_state.logs.append(
        {
            "time": collected_at(),
            "source": source,
            "status": status,
            "count": count,
            "message": message,
        }
    )


def render_metric_row(mfds: pd.DataFrame, fda: pd.DataFrame, eudamed: pd.DataFrame, integrated: pd.DataFrame) -> None:
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("식약처", f"{len(mfds):,}건")
    col2.metric("FDA", f"{len(fda):,}건")
    col3.metric("EUDAMED", f"{len(eudamed):,}건")
    col4.metric("통합 결과", f"{len(integrated):,}건")


def render_table(df: pd.DataFrame, key: str) -> None:
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "source_url": st.column_config.LinkColumn("source_url"),
        },
        key=key,
    )


@st.cache_data
def load_gmp_product_groups() -> pd.DataFrame:
    csv_path = Path(__file__).with_name("gmp_product_groups.csv")
    return pd.read_csv(csv_path, encoding="utf-8-sig")


def render_gmp_grouped_table(df: pd.DataFrame) -> None:
    if df.empty:
        st.info("검색 조건에 해당하는 품목이 없습니다.")
        return

    body_rows: list[str] = []
    for (number, group_name), group_df in df.groupby(
        ["번호", "GMP 품목군"], sort=False, dropna=False
    ):
        rowspan = len(group_df)
        group_text = str(group_name)
        if " (" in group_text:
            korean_name, english_name = group_text.split(" (", 1)
            group_html = f"{escape(korean_name)}<br><span>({escape(english_name)}</span>"
        else:
            group_html = escape(group_text)

        for row_index, (_, row) in enumerate(group_df.iterrows()):
            merged_cells = ""
            if row_index == 0:
                merged_cells = (
                    f'<td class="gmp-number" rowspan="{rowspan}">{int(number)}</td>'
                    f'<td class="gmp-group" rowspan="{rowspan}">{group_html}</td>'
                )
            body_rows.append(
                f'<tr>{merged_cells}<td class="gmp-classification">'
                f'{escape(str(row["구분"]))}</td></tr>'
            )

    table_html = f"""
    <div class="gmp-table-wrap">
      <table class="gmp-grouped-table">
        <thead>
          <tr>
            <th class="gmp-number">번호</th>
            <th class="gmp-group">품목군<br><span>(Product Group)</span></th>
            <th class="gmp-classification">구분</th>
          </tr>
        </thead>
        <tbody>{''.join(body_rows)}</tbody>
      </table>
    </div>
    <style>
      .gmp-table-wrap {{
        max-height: 650px;
        overflow: auto;
        border: 1px solid #222;
      }}
      .gmp-grouped-table {{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        color: #111;
        background: #fff;
      }}
      .gmp-grouped-table th,
      .gmp-grouped-table td {{
        border-right: 1px solid #222;
        border-bottom: 1px solid #222;
        padding: 0.45rem 0.55rem;
        line-height: 1.35;
      }}
      .gmp-grouped-table th {{
        position: sticky;
        top: 0;
        z-index: 1;
        background: #f7f7f7;
        text-align: center;
        font-weight: 600;
      }}
      .gmp-grouped-table .gmp-number {{ width: 7%; text-align: center; }}
      .gmp-grouped-table .gmp-group {{
        width: 28%;
        text-align: center;
        vertical-align: middle;
      }}
      .gmp-grouped-table .gmp-classification {{ width: 65%; }}
      .gmp-grouped-table td.gmp-number {{ vertical-align: middle; }}
      .gmp-grouped-table tr:last-child td {{ border-bottom: 0; }}
      .gmp-grouped-table th:last-child,
      .gmp-grouped-table td:last-child {{ border-right: 0; }}
    </style>
    """
    st.markdown(table_html, unsafe_allow_html=True)

def render_gmp_product_groups() -> None:
    st.subheader("GMP 품목군")
    st.caption("의료기기 제조 및 품질관리 기준 [별표 3]의 GMP 품목군을 표로 정리한 내용입니다.")

    gmp_df = load_gmp_product_groups()
    group_options = (
        gmp_df[["번호", "GMP 품목군"]]
        .drop_duplicates()
        .sort_values("번호")
    )
    group_labels = {
        int(row["번호"]): f'{int(row["번호"])}. {row["GMP 품목군"]}'
        for _, row in group_options.iterrows()
    }

    filter_col, search_col = st.columns([2, 3])
    selected_group = filter_col.selectbox(
        "품목군 선택",
        [0, *group_labels.keys()],
        format_func=lambda value: "전체 품목군" if value == 0 else group_labels[value],
        key="gmp_group_filter",
    )
    search_text = search_col.text_input(
        "표 검색",
        placeholder="품목군, 분류코드 또는 품목명을 입력하세요.",
        key="gmp_table_search",
    ).strip()

    view_df = gmp_df.copy()
    if selected_group:
        view_df = view_df[view_df["번호"] == selected_group]
    if search_text:
        match = view_df.astype(str).apply(
            lambda column: column.str.contains(search_text, case=False, na=False)
        ).any(axis=1)
        view_df = view_df[match]

    st.caption(f"총 {len(view_df):,}개 항목")
    render_gmp_grouped_table(view_df)

def render_mfds_results(tab: str) -> None:
    rows = st.session_state.get("mfds_raw_rows", [])
    active_tab = st.session_state.get("mfds_active_tab", "item")
    if active_tab != tab or not rows:
        st.info("검색 조건을 입력한 뒤 조회를 실행하세요.")
        return

    total = int(st.session_state.get("mfds_total", 0))
    page = int(st.session_state.get("mfds_page", 1))
    total_pages = max((total + 9) // 10, 1)

    st.write(f"총 {total:,}건 | {page:,}/{total_pages:,} 페이지")
    st.dataframe(mfds_rows_to_table(rows, tab), use_container_width=True, hide_index=True)

    prev_col, page_col, go_col, next_col, _ = st.columns([1, 1.4, 1, 1, 4])
    with prev_col:
        if st.button("이전", key=f"mfds_{tab}_prev", disabled=page <= 1):
            run_mfds_search(tab, st.session_state.mfds_condition, page - 1)
            st.rerun()
    with page_col:
        selected = st.number_input("페이지", min_value=1, max_value=total_pages, value=page, step=1, key=f"mfds_{tab}_goto_page_{page}")
    with go_col:
        st.write("")
        if st.button("이동", key=f"mfds_{tab}_goto", disabled=int(selected) == page):
            run_mfds_search(tab, st.session_state.mfds_condition, int(selected))
            st.rerun()
    with next_col:
        if st.button("다음", key=f"mfds_{tab}_next", disabled=page >= total_pages):
            run_mfds_search(tab, st.session_state.mfds_condition, page + 1)
            st.rerun()


def render_mfds_download_controls(tab: str, condition: dict[str, str], ready: bool, delay: float, max_pages: int | None) -> None:
    col1, col2 = st.columns([1, 1.4])
    with col1:
        if st.button("전체 엑셀 만들기", disabled=not ready, key=f"mfds_{tab}_excel"):
            try:
                filename, data, total_count, row_count, rows = make_mfds_download(condition, tab, delay, max_pages)
                st.session_state.mfds_active_tab = tab
                st.session_state.mfds_condition = condition
                st.session_state.mfds_total = total_count
                st.session_state.mfds_raw_rows = rows[:10]
                st.session_state.mfds_df = mfds_rows_to_standard(rows, tab)
                add_log("MFDS", "success", f"eMedi 엑셀 생성: {filename}", row_count)
                st.success(f"완료: {row_count:,}/{total_count:,}건")
                st.download_button(
                    "식약처 엑셀 다운로드",
                    data=data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"mfds_{tab}_download",
                )
            except Exception as exc:
                add_log("MFDS", "error", str(exc), 0)
                st.error(str(exc))
    with col2:
        if tab == "company" and st.button("업종별 업체 수 계산", disabled=not ready, key=f"mfds_{tab}_industry_count"):
            try:
                _, _, total_count, row_count, rows = make_mfds_download(condition, tab, delay, max_pages)
                counts = count_unique_mfds_companies(rows)
                m1, m2, m3 = st.columns(3)
                m1.metric("제조업체 수", f"{counts.get('제조업', 0):,}")
                m2.metric("수입업체 수", f"{counts.get('수입업', 0):,}")
                m3.metric("수집 건수", f"{row_count:,}/{total_count:,}")
                add_log("MFDS", "success", "업종별 업체 수 계산 완료", row_count)
            except Exception as exc:
                add_log("MFDS", "error", str(exc), 0)
                st.error(str(exc))


def render_mfds_item_form_compact() -> dict[str, str]:
    st.markdown("#### 품목검색")
    st.caption("자주 쓰는 조건만 먼저 입력하고, 세부 조건은 고급 검색에서 열어 사용하세요.")

    basic1, basic2, basic3 = st.columns(3)
    query2 = basic1.text_input("명칭", placeholder="제품명, 품목명, 모델명", key="mfds_item_query2_compact")
    entp_name = basic2.text_input("업체명", placeholder="업체명을 입력하세요", key="mfds_item_entp_name_compact")
    item_no = basic3.text_input("품목허가번호", placeholder="예: 제허00-000호", key="mfds_item_no_compact")

    basic4, basic5, basic6 = st.columns(3)
    item_name = basic4.text_input("품목명", key="mfds_item_name_compact")
    brand_name = basic5.text_input("제품명", key="mfds_item_brand_name_compact")
    type_name = basic6.text_input("모델명", key="mfds_item_type_name_compact")

    with st.expander("고급 검색 조건", expanded=False):
        row1 = st.columns(4)
        udidi_code = row1[0].text_input("UDI 코드", key="mfds_item_udidi_compact")
        grade = row1[1].selectbox(
            "품목등급",
            ["0", "1", "2", "3", "4"],
            format_func=lambda value: "전체" if value == "0" else f"{value}등급",
            key="mfds_item_grade_compact",
        )
        item_state = row1[2].selectbox(
            "품목상태",
            ["", "정상", "취소", "취하", "양도", "만료"],
            format_func=lambda value: "전체" if value == "" else value,
            key="mfds_item_state_compact",
        )
        mdentp_prmno = row1[3].text_input("업허가번호", key="mfds_item_mdentp_prmno_compact")

        row2 = st.columns(4)
        with row2[0]:
            indty_cd = mfds_industry_select("mfds_item_industry_compact")
        with row2[1]:
            tcsbiz_state = mfds_business_state_select("mfds_item_business_state_compact")
        mnfacr_nm = row2[2].text_input("제조원", key="mfds_item_mnfacr_nm_compact")
        query = row2[3].text_input("문구검색", placeholder="업체명, 제품명, 품목명, 모델명, 제조원", key="mfds_item_query_compact")

        row3 = st.columns(2)
        rcprslry_cd = row3[0].text_input("요양급여코드", key="mfds_item_rcprslry_cd_compact")
        md_clsf_no = row3[1].text_input("품목분류번호", key="mfds_item_md_clsf_no_compact")

        date_cols = st.columns(4)
        prm_from = date_cols[0].date_input("품목허가일자 시작", value=None, key="mfds_item_prm_from_compact")
        prm_to = date_cols[1].date_input("품목허가일자 종료", value=None, key="mfds_item_prm_to_compact")
        valid_from = date_cols[2].date_input("품목유효기간 시작", value=None, key="mfds_item_valid_from_compact")
        valid_to = date_cols[3].date_input("품목유효기간 종료", value=None, key="mfds_item_valid_to_compact")

        yes_no_cols = st.columns(4)
        with yes_no_cols[0]:
            st.caption("요양급여대상여부")
            rcprslry_trgt = mfds_yes_no("mfds_item_rcprslry_trgt_compact")
        with yes_no_cols[1]:
            st.caption("추적관리대상여부")
            trace_manage = mfds_yes_no("mfds_item_trace_manage_compact")
        with yes_no_cols[2]:
            st.caption("수출용에포함")
            xprtpp_yn = mfds_yes_no("mfds_item_xprtpp_yn_compact")
        with yes_no_cols[3]:
            st.caption("인체이식용여부")
            hmnbd_yn = mfds_yes_no("mfds_item_hmnbd_yn_compact")

    chk_group = st.radio(
        "조회조건",
        ["GROUP_BY_FIELD_01", ""],
        index=0,
        format_func=lambda value: "허가번호 단위" if value == "GROUP_BY_FIELD_01" else "모델명 단위",
        horizontal=True,
        key="mfds_item_chk_group_compact",
    )

    return normalize_condition(
        {
            "query2": query2,
            "udidiCode": locals().get("udidi_code", ""),
            "itemName": item_name,
            "itemNoFullname": item_no,
            "grade": "" if locals().get("grade", "0") == "0" else locals().get("grade", ""),
            "itemState": locals().get("item_state", ""),
            "entpName": entp_name,
            "mdentpPrmno": locals().get("mdentp_prmno", ""),
            "typeName": type_name,
            "brandName": brand_name,
            "indtyCd": "" if locals().get("indty_cd", "1|2|21|22") == "1|2|21|22" else locals().get("indty_cd", ""),
            "tcsbizRsmptSeCd": locals().get("tcsbiz_state", ""),
            "mnfacrNm": locals().get("mnfacr_nm", ""),
            "query": locals().get("query", ""),
            "rcprslryCdInptvl": locals().get("rcprslry_cd", ""),
            "mdClsfNo": locals().get("md_clsf_no", ""),
            "prdlPrmDtFrom": locals().get("prm_from").isoformat() if locals().get("prm_from") else "",
            "prdlPrmDtTo": locals().get("prm_to").isoformat() if locals().get("prm_to") else "",
            "validDateFrom": locals().get("valid_from").isoformat() if locals().get("valid_from") else "",
            "validDateTo": locals().get("valid_to").isoformat() if locals().get("valid_to") else "",
            "rcprslryTrgtYn": locals().get("rcprslry_trgt", ""),
            "traceManageTargetYn": locals().get("trace_manage", ""),
            "xprtppYn": locals().get("xprtpp_yn", ""),
            "hmnbdTspnttyMdYn": locals().get("hmnbd_yn", ""),
            "chkGroup": chk_group,
        },
        preserve_empty_keys={"chkGroup"},
    )

def mfds_form_label(text: str) -> None:
    st.markdown(
        f"""
        <div style="
            border-top: 1px solid #2f5f9f;
            min-height: 42px;
            display: flex;
            align-items: center;
            font-weight: 700;
            color: #123a70;
            font-size: 1.02rem;
            line-height: 1.25;
        ">{escape(text)}</div>
        """,
        unsafe_allow_html=True,
    )


def mfds_pair_row(left_label: str, left_widget: Callable[[], Any], right_label: str, right_widget: Callable[[], Any]) -> tuple[Any, Any]:
    cols = st.columns([0.9, 2.4, 0.9, 2.4], gap="large")
    with cols[0]:
        mfds_form_label(left_label)
    with cols[1]:
        left_value = left_widget()
    with cols[2]:
        mfds_form_label(right_label)
    with cols[3]:
        right_value = right_widget()
    return left_value, right_value


def mfds_two_dates(start_key: str, end_key: str, start_label: str, end_label: str) -> tuple[date | None, date | None]:
    cols = st.columns(2)
    start = cols[0].date_input(start_label, value=None, key=start_key, label_visibility="collapsed")
    end = cols[1].date_input(end_label, value=None, key=end_key, label_visibility="collapsed")
    return start, end


def render_mfds_item_form_aligned() -> dict[str, str]:
    st.markdown("#### 품목검색")

    mfds_pair_row(
        "일상용어",
        lambda: st.button("일상용어 검색 바로가기", use_container_width=True, disabled=True, key="mfds_common_term_link"),
        "",
        lambda: (
            lambda cols: (
                cols[0].button("일상용어 신청", use_container_width=True, disabled=True, key="mfds_common_term_request"),
                cols[1].button("일상용어 신청 결과", use_container_width=True, disabled=True, key="mfds_common_term_result"),
            )
        )(st.columns(2)),
    )
    query2, udidi_code = mfds_pair_row(
        "명칭",
        lambda: st.text_input("명칭", placeholder="제품명/품목명/모델명", key="mfds_item_query2_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
        "UDI코드",
        lambda: st.text_input("UDI코드", placeholder="UDI코드 전체를 입력하세요.", key="mfds_item_udidi_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
    )
    grade, item_state = mfds_pair_row(
        "품목등급",
        lambda: st.selectbox(
            "품목등급",
            ["0", "1", "2", "3", "4"],
            format_func=lambda value: "전체" if value == "0" else f"{value}등급",
            key="mfds_item_grade_aligned",
            label_visibility="collapsed",
        ),
        "품목상태",
        lambda: st.selectbox(
            "품목상태",
            ["", "정상", "취소", "취하", "양도", "만료"],
            format_func=lambda value: "전체" if value == "" else value,
            key="mfds_item_state_aligned",
            label_visibility="collapsed",
        ),
    )
    item_no, entp_name = mfds_pair_row(
        "품목허가번호",
        lambda: st.text_input("품목허가번호", placeholder="예시) 수허00-000호(공백없이)", key="mfds_item_no_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
        "업체명",
        lambda: st.text_input("업체명", placeholder="업체명을 입력하세요.", key="mfds_item_entp_name_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
    )
    indty_cd, tcsbiz_state = mfds_pair_row(
        "업종",
        lambda: mfds_industry_select("mfds_item_industry_aligned"),
        "업상태",
        lambda: mfds_business_state_select("mfds_item_business_state_aligned"),
    )
    mdentp_prmno, mnfacr_nm = mfds_pair_row(
        "업허가번호",
        lambda: st.text_input("업허가번호", placeholder="업허가번호 입력하세요.", key="mfds_item_mdentp_prmno_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
        "제조자",
        lambda: st.text_input("제조자", placeholder="제조자를 입력하세요.", key="mfds_item_mnfacr_nm_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
    )
    type_name, brand_name = mfds_pair_row(
        "모델명",
        lambda: st.text_input("모델명", placeholder="모델명을 입력하세요.", key="mfds_item_type_name_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
        "제품명",
        lambda: st.text_input("제품명", placeholder="제품명을 입력하세요.", key="mfds_item_brand_name_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
    )
    item_name, query = mfds_pair_row(
        "품목명",
        lambda: st.text_input("품목명", placeholder="품목명을 입력하세요.", key="mfds_item_name_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
        "문구검색",
        lambda: st.text_input("문구검색", placeholder="업체명/제품명/품목명/모델명/제조자 등", key="mfds_item_query_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
    )
    rcprslry_cd, md_clsf_no = mfds_pair_row(
        "요양급여코드",
        lambda: st.text_input("요양급여코드", placeholder="'요양급여코드' 전체를 입력하세요.", key="mfds_item_rcprslry_cd_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
        "품목분류번호",
        lambda: st.text_input("품목분류번호", placeholder="'품목분류번호' 전체를 정확하게 입력하세요.", key="mfds_item_md_clsf_no_aligned", label_visibility="collapsed", on_change=submit_mfds_enter_search, args=("item", current_mfds_item_aligned_condition)),
    )
    (prm_from, prm_to), (valid_from, valid_to) = mfds_pair_row(
        "품목허가일자",
        lambda: mfds_two_dates("mfds_item_prm_from_aligned", "mfds_item_prm_to_aligned", "품목허가일자 시작", "품목허가일자 종료"),
        "품목유효기간\n(만료일)",
        lambda: mfds_two_dates("mfds_item_valid_from_aligned", "mfds_item_valid_to_aligned", "품목유효기간 시작", "품목유효기간 종료"),
    )
    rcprslry_trgt, trace_manage = mfds_pair_row(
        "요양급여대상여부",
        lambda: mfds_yes_no("mfds_item_rcprslry_trgt_aligned"),
        "추적관리대상여부",
        lambda: mfds_yes_no("mfds_item_trace_manage_aligned"),
    )
    xprtpp_yn, hmnbd_yn = mfds_pair_row(
        "수출용에한함",
        lambda: mfds_yes_no("mfds_item_xprtpp_yn_aligned"),
        "인체이식용여부",
        lambda: mfds_yes_no("mfds_item_hmnbd_yn_aligned"),
    )

    _, center, _ = st.columns([1.8, 2.4, 1.8])
    with center:
        chk_group = st.radio(
            "조회조건",
            ["GROUP_BY_FIELD_01", ""],
            index=0,
            format_func=lambda value: "허가번호 단위" if value == "GROUP_BY_FIELD_01" else "모델명 단위",
            horizontal=True,
            key="mfds_item_chk_group_aligned",
        )

    return normalize_condition(
        {
            "query2": query2,
            "udidiCode": udidi_code,
            "itemName": item_name,
            "itemNoFullname": item_no,
            "grade": "" if grade == "0" else grade,
            "itemState": item_state,
            "entpName": entp_name,
            "mdentpPrmno": mdentp_prmno,
            "typeName": type_name,
            "brandName": brand_name,
            "indtyCd": "" if indty_cd == "1|2|21|22" else indty_cd,
            "tcsbizRsmptSeCd": tcsbiz_state,
            "mnfacrNm": mnfacr_nm,
            "query": query,
            "rcprslryCdInptvl": rcprslry_cd,
            "mdClsfNo": md_clsf_no,
            "prdlPrmDtFrom": prm_from.isoformat() if prm_from else "",
            "prdlPrmDtTo": prm_to.isoformat() if prm_to else "",
            "validDateFrom": valid_from.isoformat() if valid_from else "",
            "validDateTo": valid_to.isoformat() if valid_to else "",
            "rcprslryTrgtYn": rcprslry_trgt,
            "traceManageTargetYn": trace_manage,
            "xprtppYn": xprtpp_yn,
            "hmnbdTspnttyMdYn": hmnbd_yn,
            "chkGroup": chk_group,
        },
        preserve_empty_keys={"chkGroup"},
    )


def render_mfds_company_form_compact() -> dict[str, str]:
    st.markdown("#### 업체검색")
    st.caption("업체명이나 업허가번호로 먼저 조회하고, 업종/상태/일자는 필요할 때만 추가하세요.")

    basic1, basic2, basic3 = st.columns(3)
    entp_name = basic1.text_input(
        "업체명",
        placeholder="업체명을 입력하세요",
        key="mfds_company_entp_name_compact",
        on_change=submit_mfds_enter_search,
        args=("company", current_mfds_company_compact_condition),
    )
    mdentp_prmno = basic2.text_input(
        "업허가번호",
        key="mfds_company_mdentp_prmno_compact",
        on_change=submit_mfds_enter_search,
        args=("company", current_mfds_company_compact_condition),
    )
    rprsv_nm = basic3.text_input(
        "대표자",
        key="mfds_company_rprsv_nm_compact",
        on_change=submit_mfds_enter_search,
        args=("company", current_mfds_company_compact_condition),
    )

    with st.expander("고급 검색 조건", expanded=False):
        row1 = st.columns(4)
        with row1[0]:
            indty_cd = mfds_industry_select("mfds_company_industry_compact")
        with row1[1]:
            tcsbiz_state = mfds_business_state_select("mfds_company_business_state_compact")
        date_from = row1[2].date_input("업허가일자 시작", value=None, key="mfds_company_prm_from_compact")
        date_to = row1[3].date_input("업허가일자 종료", value=None, key="mfds_company_prm_to_compact")

    return normalize_condition(
        {
            "entpName": entp_name,
            "mdentpPrmno": mdentp_prmno,
            "rprsvNm": rprsv_nm,
            "indtyCd": "" if locals().get("indty_cd", "1|2|21|22") == "1|2|21|22" else locals().get("indty_cd", ""),
            "tcsbizRsmptSeCd": locals().get("tcsbiz_state", ""),
            "entpPrmDtFrom": locals().get("date_from").isoformat() if locals().get("date_from") else "",
            "entpPrmDtTo": locals().get("date_to").isoformat() if locals().get("date_to") else "",
        }
    )


def render_mfds_tab(delay: float, max_pages: int | None) -> None:
    st.subheader("식약처 eMedi 검색")
    st.caption("기존 start_emedi.bat로 실행하던 eMedi 품목검색/업체검색 기능을 이 탭 안에서 수행합니다.")

    item_tab, company_tab = st.tabs(["품목검색", "업체검색"])

    with item_tab:
        condition = render_mfds_item_form_aligned()
        ready = has_real_condition(condition)
        action_col, download_col, _ = st.columns([1, 1.4, 5])
        with action_col:
            if st.button("검색", disabled=not ready, type="primary", key="mfds_item_search"):
                try:
                    run_mfds_search("item", condition, 1)
                    st.rerun()
                except Exception as exc:
                    add_log("MFDS", "error", str(exc), 0)
                    st.error(str(exc))
        with download_col:
            render_mfds_download_controls("item", condition, ready, delay, max_pages)
        if st.session_state.get("mfds_item_enter_error"):
            st.error(st.session_state.mfds_item_enter_error)
        render_mfds_results("item")

    with company_tab:
        condition = render_mfds_company_form_compact()
        ready = has_real_condition(condition)
        action_col, download_col, _ = st.columns([1, 2.6, 4])
        with action_col:
            if st.button("검색", disabled=not ready, type="primary", key="mfds_company_search"):
                try:
                    run_mfds_search("company", condition, 1)
                    st.rerun()
                except Exception as exc:
                    add_log("MFDS", "error", str(exc), 0)
                    st.error(str(exc))
        with download_col:
            render_mfds_download_controls("company", condition, ready, delay, max_pages)
        if st.session_state.get("mfds_company_enter_error"):
            st.error(st.session_state.mfds_company_enter_error)
        render_mfds_results("company")


def update_combined_fda_results() -> None:
    frames = [
        st.session_state.fda_510k_pma_result_df,
        st.session_state.fda_tplc_result_df,
    ]
    nonempty_frames = [frame for frame in frames if not frame.empty]
    combined = (
        pd.concat(nonempty_frames, ignore_index=True).drop_duplicates().reset_index(drop=True)
        if nonempty_frames
        else pd.DataFrame(columns=FDA_RESULT_COLUMNS)
    )
    st.session_state.fda_result_df = combined
    st.session_state.fda_df = fda_results_to_standard(combined)
    st.session_state.fda_source_counts = {
        **st.session_state.fda_510k_pma_counts,
        **st.session_state.fda_tplc_maude_counts,
    }


def render_fda_result_table(df: pd.DataFrame, title: str, key: str, file_prefix: str) -> None:
    st.markdown(f"#### {title}")
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        key=key,
        column_config={
            "원본 사이트 링크": st.column_config.LinkColumn("원본 사이트 링크"),
        },
    )
    if not df.empty:
        st.download_button(
            f"{title} CSV 다운로드",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{file_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True,
            key=f"{key}_download",
        )


def render_fda_510k_pma_tab(default_product_name: str, default_approval_no: str) -> None:
    with st.expander("참고 FDA 사이트", expanded=False):
        st.markdown(
            f"""
            - [510(k) Premarket Notification]({FDA_510K_URL})
            - [PMA Premarket Approval]({FDA_PMA_URL})
            """
        )

    with st.form("fda_510k_pma_search_form"):
        source_cols = st.columns(2)
        include_510k = source_cols[0].checkbox("510(k)", value=True, key="fda_510k_pma_include_510k")
        include_pma = source_cols[1].checkbox("PMA", value=True, key="fda_510k_pma_include_pma")

        query_cols = st.columns(4)
        product_code = query_cols[0].text_input("Product Code", placeholder="예: DXN", key="fda_510k_pma_product_code")
        device_name = query_cols[1].text_input("Device Name", value=default_product_name, placeholder="예: blood pressure", key="fda_510k_pma_device_name")
        k_number = query_cols[2].text_input("510(k) Number", value=default_approval_no if default_approval_no.upper().startswith("K") else "", placeholder="예: K241234", key="fda_510k_pma_k_number")
        pma_number = query_cols[3].text_input("PMA Number", value=default_approval_no if default_approval_no.upper().startswith("P") else "", placeholder="예: P840001", key="fda_510k_pma_pma_number")

        date_cols = st.columns([1, 1, 1, 3])
        start_date = date_cols[0].date_input(
            "Decision Date 시작일",
            value=date(date.today().year, 1, 1),
            key="fda_510k_pma_start_date",
        )
        end_date = date_cols[1].date_input(
            "Decision Date 종료일",
            value=date.today(),
            key="fda_510k_pma_end_date",
        )
        run = date_cols[2].form_submit_button("510(k), PMA 조회", type="primary", use_container_width=True)

    if run:
        if start_date > end_date:
            st.warning("시작일은 종료일보다 늦을 수 없습니다.")
        elif not any([include_510k, include_pma]):
            st.warning("510(k) 또는 PMA 중 하나 이상 선택하세요.")
        else:
            try:
                with st.spinner("FDA 510(k), PMA 공개 DB를 조회하는 중입니다..."):
                    result_df, _, counts = search_fda_sources(
                        include_510k, include_pma, False, False,
                        product_code.strip(), device_name.strip(),
                        k_number.strip(), pma_number.strip(), "", start_date, end_date,
                    )
                st.session_state.fda_510k_pma_result_df = result_df
                st.session_state.fda_510k_pma_counts = {
                    "510(k)": counts.get("510(k)", 0),
                    "PMA": counts.get("PMA", 0),
                }
                update_combined_fda_results()
                add_log("FDA", "success", "510(k), PMA 조회 완료", len(result_df))
                st.success(f"510(k), PMA 조회 완료: {len(result_df):,}건")
            except Exception as exc:
                add_log("FDA", "error", str(exc), 0)
                st.error(str(exc))

    counts = st.session_state.fda_510k_pma_counts
    metric_cols = st.columns(2)
    metric_cols[0].metric("510(k)", f"{counts.get('510(k)', 0):,}")
    metric_cols[1].metric("PMA", f"{counts.get('PMA', 0):,}")
    render_fda_result_table(
        st.session_state.fda_510k_pma_result_df,
        "510(k), PMA 결과",
        "fda_510k_pma_result_table",
        "fda_510k_pma_results",
    )


def render_fda_tplc_maude_tab() -> None:
    with st.expander("참고 FDA 사이트", expanded=False):
        st.markdown(
            f"""
            - [TPLC Total Product Life Cycle]({FDA_TPLC_URL})
            - [MAUDE Text Search]({FDA_MAUDE_URL})
            """
        )

    with st.form("fda_tplc_maude_search_form"):
        source_cols = st.columns(2)
        include_tplc = source_cols[0].checkbox("TPLC", value=True, key="fda_tplc_maude_include_tplc")
        include_maude = source_cols[1].checkbox("MAUDE", value=True, key="fda_tplc_maude_include_maude")

        query_cols = st.columns(3)
        product_code = query_cols[0].text_input("Product Code", placeholder="예: DXN", key="fda_tplc_maude_product_code")
        device_name = query_cols[1].text_input("Device Name", placeholder="예: blood pressure", key="fda_tplc_maude_device_name")
        regulation_number = query_cols[2].text_input("Regulation No.", placeholder="예: 870.1130", key="fda_tplc_maude_regulation")

        st.caption("날짜 범위는 MAUDE의 Date Report Received에 적용되며, TPLC 분류 조회에는 적용되지 않습니다.")
        date_cols = st.columns([1, 1, 1, 3])
        start_date = date_cols[0].date_input(
            "Date Report Received 시작일",
            value=date(date.today().year, 1, 1),
            key="fda_tplc_maude_start_date",
        )
        end_date = date_cols[1].date_input(
            "Date Report Received 종료일",
            value=date.today(),
            key="fda_tplc_maude_end_date",
        )
        run = date_cols[2].form_submit_button("TPLC, MAUDE 조회", type="primary", use_container_width=True)

    if run:
        if start_date > end_date:
            st.warning("시작일은 종료일보다 늦을 수 없습니다.")
        elif not any([include_tplc, include_maude]):
            st.warning("TPLC 또는 MAUDE 중 하나 이상 선택하세요.")
        elif include_tplc and not any([product_code.strip(), regulation_number.strip()]) and not include_maude:
            st.warning("TPLC 조회에는 Product Code 또는 Regulation No.가 필요합니다.")
        else:
            try:
                with st.spinner("FDA TPLC, MAUDE 공개 DB를 조회하는 중입니다..."):
                    tplc_df, maude_df, counts = search_fda_sources(
                        False, False, include_tplc, include_maude,
                        product_code.strip(), device_name.strip(), "", "",
                        regulation_number.strip(), start_date, end_date,
                    )
                st.session_state.fda_tplc_result_df = tplc_df
                st.session_state.fda_maude_df = maude_df
                st.session_state.fda_tplc_maude_counts = {
                    "TPLC": counts.get("TPLC", 0),
                    "MAUDE": counts.get("MAUDE", 0),
                }
                update_combined_fda_results()
                add_log("FDA", "success", "TPLC, MAUDE 조회 완료", len(tplc_df) + len(maude_df))
                st.success(f"TPLC, MAUDE 조회 완료: {len(tplc_df) + len(maude_df):,}건")
            except Exception as exc:
                add_log("FDA", "error", str(exc), 0)
                st.error(str(exc))

    counts = st.session_state.fda_tplc_maude_counts
    metric_cols = st.columns(2)
    metric_cols[0].metric("TPLC/Class", f"{counts.get('TPLC', 0):,}")
    metric_cols[1].metric("MAUDE", f"{counts.get('MAUDE', 0):,}")

    render_fda_result_table(
        st.session_state.fda_tplc_result_df,
        "TPLC 결과",
        "fda_tplc_result_table",
        "fda_tplc_results",
    )
    st.markdown("#### MAUDE 결과")
    st.dataframe(
        st.session_state.fda_maude_df,
        use_container_width=True,
        hide_index=True,
        key="fda_maude_result_table",
    )


def render_fda_tab(default_product_name: str = "", default_approval_no: str = "") -> None:
    st.subheader("FDA 인허가 정보 조회")
    st.caption("FDA 공개 데이터베이스를 인허가 정보와 제품 생애주기/이상사례로 구분해 조회합니다.")

    approval_tab, lifecycle_tab = st.tabs(["510(k), PMA", "TPLC, MAUDE"])
    with approval_tab:
        render_fda_510k_pma_tab(default_product_name, default_approval_no)
    with lifecycle_tab:
        render_fda_tplc_maude_tab()

def pms_dependency_status() -> dict[str, bool]:
    return {
        "openpyxl": importlib.util.find_spec("openpyxl") is not None,
        "xlsxwriter": importlib.util.find_spec("xlsxwriter") is not None,
        "deep_translator": importlib.util.find_spec("deep_translator") is not None,
        "ollama": importlib.util.find_spec("ollama") is not None,
    }


def load_pms_module():
    module_path = Path(__file__).with_name("pms_auto_arrangement.py")
    spec = importlib.util.spec_from_file_location("pms_auto_arrangement", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("PMS 모듈을 불러올 수 없습니다.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def pms_output_path_from_input(input_file: str) -> str:
    p = Path(input_file)
    return str(p.parent / f"{p.stem}_arranged{p.suffix}")


def excel_mime_type(filename: str) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix == ".xls":
        return "application/vnd.ms-excel"
    if suffix == ".xlsm":
        return "application/vnd.ms-excel.sheet.macroEnabled.12"
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def render_pms_tab() -> None:
    st.subheader("PMS Auto Arrangement")
    st.caption("기존 PMS AUTO arr. AI TESTBED Local.py의 Excel 정리, 번역, 원인 분석, 심각도 평가 기능을 실행합니다.")

    deps = pms_dependency_status()
    missing = [name for name, ok in deps.items() if not ok]
    dep_cols = st.columns(4)
    for col, (name, ok) in zip(dep_cols, deps.items()):
        col.metric(name, "OK" if ok else "Missing")
    if missing:
        st.warning(
            "PMS 실행에 필요한 패키지가 아직 없습니다: "
            + ", ".join(missing)
            + "\n\n아래 명령으로 설치한 뒤 앱을 다시 실행하세요: `pip install -r requirements.txt`"
        )

    uploaded = st.file_uploader("입력 Excel 파일", type=["xlsx", "xlsm", "xls"], key="pms_uploaded_file")
    model_name = st.text_input("LLM 모델명", value="gemma3:4b", help="Ollama에 설치된 모델명을 입력하세요.", key="pms_model_name")

    sheet_names: list[str] = []
    if uploaded is not None:
        tmp_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded.name).suffix) as tmp:
                tmp.write(uploaded.getbuffer())
                tmp_path = Path(tmp.name)
            with pd.ExcelFile(tmp_path) as excel_file:
                sheet_names = list(excel_file.sheet_names)
        except Exception as exc:
            st.error(f"워크시트 목록을 읽지 못했습니다: {exc}")
        finally:
            if tmp_path is not None:
                try:
                    tmp_path.unlink(missing_ok=True)
                except PermissionError:
                    pass

    default_sheets = sheet_names[:1] if sheet_names else []
    selected_sheets = st.multiselect(
        "목표 워크시트",
        sheet_names,
        default=default_sheets,
        help="여러 시트를 선택하면 선택한 순서대로 처리합니다.",
        key="pms_target_sheets",
    )

    if uploaded is not None:
        preview_output = f"{Path(uploaded.name).stem}_arranged{Path(uploaded.name).suffix}"
        st.info(f"저장 파일명: `{preview_output}`")

    progress_main = st.progress(0, text="전체 진행률 0.0%")
    progress_llm = st.progress(0, text="LLM 분석 진행률 0.0%")
    log_area = st.empty()

    run = st.button(
        "PMS 작업 시작",
        type="primary",
        disabled=uploaded is None or bool(missing),
        use_container_width=True,
        key="pms_run",
    )

    if run:
        target_sheets = list(selected_sheets)
        if not target_sheets:
            st.error("목표 워크시트를 하나 이상 선택하세요.")
            return
        if not model_name.strip():
            st.error("LLM 모델명을 입력하세요.")
            return

        try:
            pms_module = load_pms_module()
            log_buffer = StringIO()

            def main_cb(value: float) -> None:
                value = max(0.0, min(float(value), 100.0))
                progress_main.progress(int(value), text=f"전체 진행률 {value:.1f}%")

            def llm_cb(value: float) -> None:
                value = max(0.0, min(float(value), 100.0))
                progress_llm.progress(int(value), text=f"LLM 분석 진행률 {value:.1f}%")

            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                input_path = temp_path / uploaded.name
                input_path.write_bytes(uploaded.getbuffer())
                output_path = Path(pms_output_path_from_input(str(input_path)))

                with st.spinner("PMS 작업을 실행하는 중입니다. 번역/LLM 분석이 포함되어 시간이 걸릴 수 있습니다."):
                    with contextlib.redirect_stdout(log_buffer):
                        pms_module.run_arrangement(
                            str(input_path),
                            str(output_path),
                            target_sheets,
                            model_name.strip(),
                            main_progress_cb=main_cb,
                            llm_progress_cb=llm_cb,
                        )

                if not output_path.exists():
                    raise RuntimeError("PMS 결과 파일이 생성되지 않았습니다.")

                st.session_state.pms_output_name = output_path.name
                st.session_state.pms_output_bytes = output_path.read_bytes()
                st.session_state.pms_log = log_buffer.getvalue()
                progress_main.progress(100, text="전체 진행률 100.0%")
                st.success(f"PMS 작업 완료: {output_path.name}")
                add_log("PMS", "success", f"PMS 작업 완료: {output_path.name}", 1)
        except Exception as exc:
            st.session_state.pms_log = log_buffer.getvalue() if "log_buffer" in locals() else ""
            add_log("PMS", "error", str(exc), 0)
            st.error(str(exc))

    if st.session_state.pms_log:
        log_area.text_area("진행 로그", value=st.session_state.pms_log, height=260)

    if st.session_state.pms_output_bytes:
        st.download_button(
            "PMS 결과 Excel 다운로드",
            data=st.session_state.pms_output_bytes,
            file_name=st.session_state.pms_output_name or "pms_arranged.xlsx",
            mime=excel_mime_type(st.session_state.pms_output_name or "pms_arranged.xlsx"),
            use_container_width=True,
        )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")
    ensure_state()

    st.title(APP_TITLE)

    st.markdown(
        """
        <style>
        section[data-testid="stSidebar"] [data-testid="stButton"] button {
            width: auto;
            min-height: 2.25rem;
            padding: 0.35rem 0;
            font-size: 1.5rem;
            font-weight: 700;
            border: 0;
            border-radius: 0;
            background: transparent;
            box-shadow: none;
            color: #262730;
            justify-content: flex-start;
        }
        section[data-testid="stSidebar"] [data-testid="stButton"] button:hover {
            color: #ff4b4b;
            background: transparent;
        }
        section[data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"],
        section[data-testid="stSidebar"] [data-testid="stButton"] button[data-testid="stBaseButton-primary"] {
            color: #ff4b4b;
            border-bottom: 2px solid #ff4b4b;
        }
        section[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] {
            flex-direction: row;
            gap: 1rem;
        }
        section[data-testid="stSidebar"] [data-testid="stRadio"] label {
            width: auto;
            padding: 0.35rem 0;
            font-size: 1.5rem;
            font-weight: 700;
            border-bottom: 2px solid transparent;
        }
        section[data-testid="stSidebar"] [data-testid="stRadio"] label p {
            font-size: inherit;
            font-weight: inherit;
        }
        section[data-testid="stSidebar"] [data-testid="stRadio"] label > div:first-child {
            display: none;
        }
        section[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
            color: #ff4b4b;
            border-bottom-color: #ff4b4b;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.header("메뉴")
        if st.button(
            "국가별 인허가 정보",
            type="primary" if st.session_state.main_menu == "국가별 인허가 정보" else "secondary",
            use_container_width=False,
        ):
            st.session_state.main_menu = "국가별 인허가 정보"
            st.rerun()

        if st.session_state.main_menu == "국가별 인허가 정보":
            with st.container():
                country_menu = st.radio(
                    "국가 선택",
                    ["한국", "미국", "유럽"],
                    key="country_menu",
                    horizontal=True,
                    label_visibility="collapsed",
                )
            active_page = country_menu

        for menu_name in ["PMS", "통합 결과", "설정/로그"]:
            if st.button(
                menu_name,
                type="primary" if st.session_state.main_menu == menu_name else "secondary",
                use_container_width=False,
                key=f"main_menu_{menu_name}",
            ):
                st.session_state.main_menu = menu_name
                st.rerun()

        if st.session_state.main_menu != "국가별 인허가 정보":
            active_page = st.session_state.main_menu

    keyword = ""
    manufacturer = ""
    approval_no = ""
    date_range = ()
    max_pages = 3
    request_delay = 0.5
    clear = False

    start_date = date_range[0] if isinstance(date_range, tuple) and len(date_range) > 0 else None
    end_date = date_range[1] if isinstance(date_range, tuple) and len(date_range) > 1 else None
    filters = SearchFilters(
        keyword=keyword.strip(),
        manufacturer=manufacturer.strip(),
        approval_no=approval_no.strip(),
        start_date=start_date,
        end_date=end_date,
        max_pages=int(max_pages),
        request_delay=float(request_delay),
    )

    if clear:
        st.session_state.mfds_df = pd.DataFrame(columns=STANDARD_COLUMNS)
        st.session_state.mfds_raw_rows = []
        st.session_state.mfds_active_tab = "item"
        st.session_state.mfds_condition = {}
        st.session_state.mfds_page = 1
        st.session_state.mfds_total = 0
        st.session_state.fda_df = pd.DataFrame(columns=STANDARD_COLUMNS)
        st.session_state.fda_result_df = pd.DataFrame(columns=FDA_RESULT_COLUMNS)
        st.session_state.fda_510k_pma_result_df = pd.DataFrame(columns=FDA_RESULT_COLUMNS)
        st.session_state.fda_tplc_result_df = pd.DataFrame(columns=FDA_RESULT_COLUMNS)
        st.session_state.fda_maude_df = pd.DataFrame()
        st.session_state.fda_510k_pma_counts = {"510(k)": 0, "PMA": 0}
        st.session_state.fda_tplc_maude_counts = {"TPLC": 0, "MAUDE": 0}
        st.session_state.fda_source_counts = {"510(k)": 0, "PMA": 0, "TPLC": 0, "MAUDE": 0}
        st.session_state.eudamed_df = pd.DataFrame(columns=STANDARD_COLUMNS)
        st.session_state.pms_output_name = ""
        st.session_state.pms_output_bytes = None
        st.session_state.pms_log = ""
        st.session_state.logs = []
        st.rerun()

    mfds_df = st.session_state.mfds_df
    fda_df = st.session_state.fda_df
    eudamed_df = st.session_state.eudamed_df
    integrated_df = (
        pd.concat([mfds_df, fda_df, eudamed_df], ignore_index=True)
        .drop_duplicates(subset=["source", "approval_or_submission_no", "product_code_or_udi", "product_name"])
        .reset_index(drop=True)
    )
    logs_df = pd.DataFrame(st.session_state.logs, columns=["time", "source", "status", "count", "message"])

    if active_page == "한국":
        mfds_search_tab, mfds_classification_tab = st.tabs(
            ["품목 및 업체 검색", "의료기기 품목 분류표"]
        )

        with mfds_search_tab:
            render_mfds_tab(filters.request_delay, filters.max_pages)

        with mfds_classification_tab:
            gmp_group_tab, middle_class_tab = st.tabs(["GMP 품목군", "중분류별 품목"])

            with gmp_group_tab:
                render_gmp_product_groups()

            with middle_class_tab:
                st.subheader("중분류별 품목")
                st.info("중분류별 품목 기능이 표시될 영역입니다.")

    elif active_page == "미국":
        render_fda_tab(filters.keyword, filters.approval_no)

    elif active_page == "유럽":
        st.subheader("EUDAMED 조회 결과")
        cols = st.columns(4)
        cols[0].text_input("Basic UDI-DI", key="eudamed_basic_udi")
        cols[1].text_input("SRN", key="eudamed_srn")
        cols[2].selectbox("Risk Class", ["전체", "Class I", "Class IIa", "Class IIb", "Class III", "Class A", "Class B", "Class C", "Class D"])
        cols[3].text_input("Certificate No.", key="eudamed_certificate")
        render_table(eudamed_df, "eudamed_table")

    elif active_page == "PMS":
        render_pms_tab()

    elif active_page == "통합 결과":
        st.subheader("통합 결과")
        cols = st.columns([1, 1, 1, 2])
        source_filter = cols[0].multiselect("기관", ["MFDS", "FDA", "EUDAMED"], default=["MFDS", "FDA", "EUDAMED"])
        status_filter = cols[1].text_input("상태 필터", placeholder="예: Cleared")
        class_filter = cols[2].text_input("등급 필터", placeholder="예: Class II")
        sort_by = cols[3].selectbox("정렬 기준", STANDARD_COLUMNS, index=0)

        view_df = integrated_df[integrated_df["source"].isin(source_filter)] if source_filter else integrated_df
        if status_filter:
            view_df = view_df[view_df["status"].str.contains(status_filter, case=False, na=False)]
        if class_filter:
            view_df = view_df[view_df["device_class"].str.contains(class_filter, case=False, na=False)]
        if not view_df.empty:
            view_df = view_df.sort_values(sort_by).reset_index(drop=True)

        render_table(view_df, "integrated_table")

        excel_bytes = build_excel(view_df, mfds_df, fda_df, eudamed_df, logs_df, filters)
        st.download_button(
            "엑셀 다운로드",
            data=excel_bytes,
            file_name=f"medical_device_approvals_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    elif active_page == "설정/로그":
        st.subheader("설정 / 로그")
        cols = st.columns(2)
        cols[0].metric("최대 페이지 수", filters.max_pages)
        cols[1].metric("요청 간격", f"{filters.request_delay:.1f}초")

        st.info(
            "각 기관 탭에서 개별 검색 조건을 입력하고 조회 또는 다운로드를 실행하세요."
        )
        render_table(logs_df, "logs_table")


if __name__ == "__main__":
    main()

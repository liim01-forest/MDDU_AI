#!/usr/bin/env python3
"""
Download Excel files from the MFDS eMedi medical device search page.

This tool reproduces the site's own search and Excel-download requests. It is
intended for small internal batch jobs with polite request intervals.
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import re
import shutil
import sys
import time
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlencode
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_URL = "https://emedi.mfds.go.kr"
SEARCH_URL = f"{BASE_URL}/search/data/list"
EXCEL_URL = f"{BASE_URL}/search/data/excelDownload"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
)


ITEM_FIELDS = {
    "query2": "명칭",
    "udidiCode": "UDI코드",
    "grade": "품목등급: 0, 1, 2, 3, 4",
    "itemState": "품목상태",
    "itemNoFullname": "품목허가번호",
    "entpName": "업체명",
    "indtyCd": "업종: 1|2|21|22, 1, 2, 21, 22",
    "tcsbizRsmptSeCd": "업상태",
    "mdentpPrmno": "업허가번호",
    "mnfacrNm": "제조자",
    "typeName": "모델명",
    "brandName": "제품명",
    "itemName": "품목명",
    "query": "문구검색",
    "rcprslryCdInptvl": "요양급여코드",
    "mdClsfNo": "품목분류번호",
    "prdlPrmDtFrom": "품목허가일자 시작(YYYY-MM-DD)",
    "prdlPrmDtTo": "품목허가일자 종료(YYYY-MM-DD)",
    "validDateFrom": "품목유효기간 시작(YYYY-MM-DD)",
    "validDateTo": "품목유효기간 종료(YYYY-MM-DD)",
    "rcprslryTrgtYn": "요양급여대상여부: Y, N",
    "traceManageTargetYn": "추적관리대상여부: Y, N",
    "xprtppYn": "수출용에한함: Y, N",
    "hmnbdTspnttyMdYn": "인체이식용여부: Y, N",
    "chkGroup": "조회조건: GROUP_BY_FIELD_01(허가번호 단위), 빈 값(모델명 단위)",
}

COMPANY_FIELDS = {
    "entpName": "업체명",
    "mdentpPrmno": "업허가번호",
    "indtyCd": "업종: 1|2|21|22, 1, 2, 21, 22",
    "tcsbizRsmptSeCd": "업상태",
    "rprsvNm": "대표자",
    "entpPrmDtFrom": "업허가일자 시작(YYYY-MM-DD)",
    "entpPrmDtTo": "업허가일자 종료(YYYY-MM-DD)",
}

ITEM_COLUMNS = [
    ("ENTP_NAME", "업체명"),
    ("ITEM_NAME", "품목명"),
    ("ITEM_NO_FULLNAME", "품목허가번호"),
    ("GRADE", "품목등급"),
    ("ITEM_STATE", "품목상태"),
    ("CANCEL_YN", "취소취하여부"),
    ("BRAND_NAME", "제품명"),
    ("INDTY_CD_NM", "업종"),
    ("MD_CLSF_NO", "품목분류번호"),
    ("TYPE_NAME", "모델명"),
    ("MNFTR_NTN_CD_NM", "제조국가"),
    ("RCPRSLRY_CD_INPTVL", "치료재료코드"),
    ("RCPRSLRY_TRGT_YN", "요양급여대상"),
    ("TRACE_MANAGE_TARGET_YN", "추적관리대상"),
    ("XPRTPP_YN", "수출용에한함"),
    ("ENTP_PRM_DT", "업허가일자"),
    ("MNCLT_NTN_CD_NM", "제조의뢰국가"),
    ("MEDDEV_ITEM_SEQ", "품목일련번호"),
    ("CLSFNO_NM_REGVL", "품목분류명"),
    ("PRDL_PRM_DT", "품목허가일자"),
    ("CLSFNO_ENG_NM_REGVL", "품목영문명"),
    ("TCSBIZ_RSMPT_SE_CD", "업상태"),
    ("VALID_DATE", "유효기간"),
    ("RTRCN_WDRW_DT", "취소/취하일자"),
    ("MDENTP_PRMNO", "업허가번호"),
    ("BLCD_ADDR", "업주소"),
    ("UPDT_APLY_TERM_DT", "갱신신청기간"),
    ("UDIDICD", "UDI코드"),
    ("MUMM_PUNIT_QY", "포장수량"),
    ("MNFACR_NM", "제조자"),
    ("RPRSV_NM", "대표자"),
]

COMPANY_COLUMNS = [
    ("ENTP_NAME", "업체명"),
    ("MDENTP_PRMNO", "업허가번호"),
    ("INDTY_CD_NM", "업종"),
    ("ENTP_PRM_DT", "업허가일자"),
    ("TCSBIZ_RSMPT_SE_CD", "업상태"),
    ("RPRSV_NM", "대표자"),
    ("BLCD_ADDR", "업주소(주소재지)"),
    ("MFDS_MNG_CD_CM", "관할청"),
]


class EmediError(RuntimeError):
    pass


@dataclass
class SearchResult:
    tab: str
    total_count: int
    excel_rows: list[dict[str, Any]]
    params: dict[str, str]

    @property
    def excel_tab_gubun(self) -> str:
        return "2" if self.tab == "company" else "1"


def build_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=4,
        connect=4,
        read=4,
        status=4,
        backoff_factor=1.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET", "POST"),
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            "Origin": BASE_URL,
            "Referer": f"{BASE_URL}/search/data/MNU20237",
        }
    )
    request_with_retry(session, "GET", f"{BASE_URL}/search/data/MNU20237", timeout=30)
    return session


def request_with_retry(session: requests.Session, method: str, url: str, attempts: int = 3, **kwargs: Any) -> requests.Response:
    last_exc: requests.RequestException | None = None
    for attempt in range(1, attempts + 1):
        try:
            return session.request(method, url, **kwargs)
        except requests.RequestException as exc:
            last_exc = exc
            if attempt == attempts:
                break
            time.sleep(1.5 * attempt)
    assert last_exc is not None
    raise last_exc


def normalize_count(value: str | None) -> int:
    if not value:
        return 0
    digits = re.sub(r"[^0-9]", "", value)
    return int(digits) if digits else 0


def extract_script_json(page: str, var_name: str) -> list[dict[str, Any]]:
    match = re.search(rf"var\s+{re.escape(var_name)}\s*=\s*(\[.*?\]);", page, re.S)
    if not match:
        return []
    raw = html.unescape(match.group(1))
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise EmediError(f"{var_name} JSON 파싱에 실패했습니다: {exc}") from exc
    if not isinstance(value, list):
        raise EmediError(f"{var_name} 값이 목록이 아닙니다.")
    return value


def extract_total(page: str, tab: str) -> int:
    input_id = "totSearchCnt2" if tab == "company" else "totSearchCnt"
    count_id = "countExcel2" if tab == "company" else "countExcel"

    input_match = re.search(rf'id="{input_id}"\s+value="([^"]*)"', page)
    if input_match:
        return normalize_count(input_match.group(1))

    count_match = re.search(rf'id\s*=\s*"{count_id}"[^>]*>(.*?)</b>', page, re.S)
    if count_match:
        return normalize_count(count_match.group(1))

    return 0


def default_item_params() -> dict[str, str]:
    return {
        "chkList": "1",
        "toggleBtnState": "",
        "nowPageNum": "1",
        "tabGubun": "1",
        "tcsbizRsmptSeCdNm": "",
        "indtyCdNm": "",
        "itemStateNm": "",
        "mnftrNtnCdNm": "",
        "tmpQrBarcode": "",
        "query2": "",
        "udidiCode": "",
        "grade": "0",
        "itemState": "",
        "itemNoFullname": "",
        "entpName": "",
        "indtyCd": "1|2|21|22",
        "tcsbizRsmptSeCd": "",
        "mdentpPrmno": "",
        "mnfacrNm": "",
        "typeName": "",
        "brandName": "",
        "itemName": "",
        "query": "",
        "rcprslryCdInptvl": "",
        "mdClsfNo": "",
        "prdlPrmDtFrom": "",
        "prdlPrmDtTo": "",
        "validDateFrom": "",
        "validDateTo": "",
        "rcprslryTrgtYn": "",
        "traceManageTargetYn": "",
        "xprtppYn": "",
        "hmnbdTspnttyMdYn": "",
        "chkGroup": "GROUP_BY_FIELD_01",
        "pageNum": "1",
        "searchYn": "true",
        "searchAfKey": "",
        "sort": "",
        "sortOrder": "",
        "searchOn": "Y",
    }


def default_company_params() -> dict[str, str]:
    return {
        "chkList": "2",
        "chkGroup": "GROUP_BY_FIELD_02",
        "nowPageNum2": "1",
        "tabGubun": "2",
        "indtyCdNm": "",
        "tcsbizRsmptSeCdNm": "",
        "indtyCd": "1|2|21|22",
        "tcsbizRsmptSeCd": "",
        "entpName": "",
        "mdentpPrmno": "",
        "rprsvNm": "",
        "entpPrmDtFrom": "",
        "entpPrmDtTo": "",
        "pageNum2": "1",
        "searchYn": "true",
        "searchAfKey2": "",
        "sort": "",
        "sortOrder": "",
        "searchOn": "Y",
    }


def has_search_condition(params: dict[str, str], tab: str) -> bool:
    allowed = COMPANY_FIELDS if tab == "company" else ITEM_FIELDS
    non_search_keys = {"chkGroup"}
    for key in allowed:
        if key in non_search_keys:
            continue
        value = params.get(key, "").strip()
        if value and not (key == "indtyCd" and value == "1|2|21|22") and not (key == "grade" and value == "0"):
            return True
    return False


def search(session: requests.Session, tab: str, overrides: dict[str, str]) -> SearchResult:
    params = default_company_params() if tab == "company" else default_item_params()
    params.update({k: v for k, v in overrides.items() if v is not None})

    if not has_search_condition(params, tab):
        raise EmediError("검색 조건이 없습니다. 사이트 정책상 하나 이상의 검색항목이 필요합니다.")

    resp = request_with_retry(session, "POST", SEARCH_URL, data=params, timeout=90)
    resp.raise_for_status()

    var_name = "excelCompList" if tab == "company" else "excelItemList"
    rows = extract_script_json(resp.text, var_name)
    total_count = extract_total(resp.text, tab)

    return SearchResult(tab=tab, total_count=total_count, excel_rows=rows, params=params)


def search_page(session: requests.Session, tab: str, base_condition: dict[str, str], page_num: int) -> SearchResult:
    overrides = dict(base_condition)
    if tab == "company":
        overrides["pageNum2"] = str(page_num)
        overrides["nowPageNum2"] = str((page_num - 1) * 10 if page_num > 1 else 1)
        overrides["searchYn"] = "" if page_num > 1 else "true"
    else:
        overrides["pageNum"] = str(page_num)
        overrides["nowPageNum"] = str((page_num - 1) * 10 if page_num > 1 else 1)
        overrides["searchYn"] = "" if page_num > 1 else "true"
    return search(session, tab, overrides)


def collect_all_rows(
    session: requests.Session,
    tab: str,
    condition: dict[str, str],
    delay: float,
    max_pages: int | None = None,
    progress: Callable[[int, int, int], None] | None = None,
) -> tuple[int, list[dict[str, Any]]]:
    first = search_page(session, tab, condition, 1)
    rows = list(first.excel_rows)
    total_pages = (first.total_count + 9) // 10
    if max_pages is not None:
        total_pages = min(total_pages, max_pages)

    print(f"  총 {first.total_count:,}건, 페이지 {total_pages:,}개 수집 예정")
    if progress:
        progress(1 if total_pages else 0, total_pages, len(rows))
    for page_num in range(2, total_pages + 1):
        if delay > 0:
            time.sleep(delay)
        page = search_page(session, tab, condition, page_num)
        rows.extend(page.excel_rows)
        if progress:
            progress(page_num, total_pages, len(rows))
        if page_num % 10 == 0 or page_num == total_pages:
            print(f"  진행: {page_num:,}/{total_pages:,} 페이지, {len(rows):,}행")
    return first.total_count, rows


def download_excel(session: requests.Session, result: SearchResult, output_path: Path) -> None:
    if result.total_count < 1 or not result.excel_rows:
        raise EmediError("검색 결과가 없습니다.")

    payload = {
        "totalListSize": str(result.total_count),
        "excelTabGubun": result.excel_tab_gubun,
        "excelMapList": result.excel_rows,
    }

    headers = {
        "Accept": "*/*",
        "Content-Type": "application/json;charset=utf-8",
        "Referer": f"{BASE_URL}/search/data/MNU20237",
    }
    params = urlencode(result.params, doseq=False)
    resp = request_with_retry(
        session,
        "POST",
        f"{EXCEL_URL}?{params}",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        timeout=180,
    )
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "")
    if not resp.content or ("html" in content_type.lower() and not resp.content.startswith(b"PK")):
        raise EmediError(f"엑셀 파일 응답이 아닙니다. Content-Type={content_type}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(resp.content)


def excel_col_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def inline_cell(row: int, col: int, value: Any, style: int | None = None) -> str:
    ref = f"{excel_col_name(col)}{row}"
    style_attr = f' s="{style}"' if style is not None else ""
    text = "" if value is None else str(value)
    text = escape(text, {'"': "&quot;"})
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{text}</t></is></c>'


def write_merged_xlsx(rows: list[dict[str, Any]], tab: str, output_path: Path) -> None:
    columns = COMPANY_COLUMNS if tab == "company" else ITEM_COLUMNS
    sheet_rows: list[str] = []
    header = "".join(inline_cell(1, col_idx, label, style=1) for col_idx, (_, label) in enumerate(columns, start=1))
    sheet_rows.append(f'<row r="1">{header}</row>')

    for row_idx, item in enumerate(rows, start=2):
        cells = []
        for col_idx, (key, _) in enumerate(columns, start=1):
            cells.append(inline_cell(row_idx, col_idx, item.get(key, "")))
        sheet_rows.append(f'<row r="{row_idx}">{"".join(cells)}</row>')

    last_col = excel_col_name(len(columns))
    dimension = f"A1:{last_col}{max(len(rows) + 1, 1)}"
    cols_xml = "".join(
        f'<col min="{idx}" max="{idx}" width="{min(max(len(label) + 4, 12), 36)}" customWidth="1"/>'
        for idx, (_, label) in enumerate(columns, start=1)
    )
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>{cols_xml}</cols>
  <sheetData>{"".join(sheet_rows)}</sheetData>
  <autoFilter ref="{dimension}"/>
</worksheet>'''

    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="검색결과" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>eMedi Excel Downloader</dc:creator>
  <cp:lastModifiedBy>eMedi Excel Downloader</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''
    app = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>eMedi Excel Downloader</Application>
</Properties>'''

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", styles)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("docProps/core.xml", core)
        zf.writestr("docProps/app.xml", app)


def company_key(row: dict[str, Any]) -> tuple[str, str]:
    stable_id = str(row.get("MDENTP_SNO") or "").strip()
    if stable_id:
        return ("sno", stable_id)
    permit_no = str(row.get("MDENTP_PRMNO") or "").strip()
    name = str(row.get("ENTP_NAME") or "").strip()
    return ("permit_name", f"{permit_no}|{name}")


def unique_companies(rows: list[dict[str, Any]], industry_name: str) -> list[dict[str, Any]]:
    seen: set[tuple[str, str]] = set()
    companies: list[dict[str, Any]] = []
    for row in rows:
        if str(row.get("INDTY_CD_NM") or "").strip() != industry_name:
            continue
        key = company_key(row)
        if not key[1].strip("|") or key in seen:
            continue
        seen.add(key)
        companies.append(row)
    return sorted(companies, key=lambda item: (str(item.get("ENTP_NAME") or ""), str(item.get("MDENTP_PRMNO") or "")))


def company_sheet_xml(rows: list[dict[str, Any]]) -> str:
    columns = [
        ("ENTP_NAME", "업체명"),
        ("MDENTP_PRMNO", "업허가번호"),
        ("INDTY_CD_NM", "업종"),
        ("TCSBIZ_RSMPT_SE_CD", "업상태"),
        ("RPRSV_NM", "대표자"),
        ("ENTP_PRM_DT", "업허가일자"),
        ("BLCD_ADDR", "주소"),
        ("MDENTP_SNO", "업체고유번호"),
    ]
    sheet_rows = [
        f'<row r="1">{"".join(inline_cell(1, col_idx, label) for col_idx, (_, label) in enumerate(columns, start=1))}</row>'
    ]
    for row_idx, item in enumerate(rows, start=2):
        sheet_rows.append(
            f'<row r="{row_idx}">'
            + "".join(inline_cell(row_idx, col_idx, item.get(key, "")) for col_idx, (key, _) in enumerate(columns, start=1))
            + "</row>"
        )
    last_col = excel_col_name(len(columns))
    dimension = f"A1:{last_col}{max(len(rows) + 1, 1)}"
    widths = [28, 18, 14, 14, 16, 14, 60, 18]
    cols_xml = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(widths, start=1)
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="{dimension}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>{cols_xml}</cols>
  <sheetData>{"".join(sheet_rows)}</sheetData>
  <autoFilter ref="{dimension}"/>
</worksheet>'''


def _next_sheet_index(names: list[str]) -> int:
    used = []
    for name in names:
        match = re.match(r"xl/worksheets/sheet(\d+)\.xml$", name)
        if match:
            used.append(int(match.group(1)))
    return max(used, default=0) + 1


def _next_relationship_id(root: ET.Element) -> str:
    used = []
    for rel in list(root):
        rid = rel.attrib.get("Id", "")
        match = re.match(r"rId(\d+)$", rid)
        if match:
            used.append(int(match.group(1)))
    return f"rId{max(used, default=0) + 1}"


def append_industry_sheets(xlsx_path: Path, rows: list[dict[str, Any]]) -> None:
    manufacturer_rows = unique_companies(rows, "제조업")
    importer_rows = unique_companies(rows, "수입업")
    extra_sheets = [
        ("제조업체", company_sheet_xml(manufacturer_rows)),
        ("수입업체", company_sheet_xml(importer_rows)),
    ]

    spreadsheet_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    content_type_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", spreadsheet_ns)
    ET.register_namespace("r", rel_ns)

    with zipfile.ZipFile(xlsx_path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    sheet_index = _next_sheet_index(list(entries))
    sheet_paths: list[tuple[str, str, str]] = []
    for title, xml in extra_sheets:
        path = f"xl/worksheets/sheet{sheet_index}.xml"
        entries[path] = xml.encode("utf-8")
        sheet_paths.append((title, path, f"worksheets/sheet{sheet_index}.xml"))
        sheet_index += 1

    workbook_path = "xl/workbook.xml"
    rels_path = "xl/_rels/workbook.xml.rels"
    content_types_path = "[Content_Types].xml"

    workbook_root = ET.fromstring(entries[workbook_path])
    sheets_node = workbook_root.find(f"{{{spreadsheet_ns}}}sheets")
    if sheets_node is None:
        raise EmediError("엑셀 workbook.xml에서 sheets 노드를 찾을 수 없습니다.")
    sheet_ids = [int(node.attrib.get("sheetId", "0")) for node in sheets_node.findall(f"{{{spreadsheet_ns}}}sheet")]
    next_sheet_id = max(sheet_ids, default=0) + 1

    rel_root = ET.fromstring(entries[rels_path])
    for title, _, target in sheet_paths:
        rid = _next_relationship_id(rel_root)
        ET.SubElement(
            rel_root,
            f"{{{package_rel_ns}}}Relationship",
            {
                "Id": rid,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "Target": target,
            },
        )
        ET.SubElement(
            sheets_node,
            f"{{{spreadsheet_ns}}}sheet",
            {"name": title, "sheetId": str(next_sheet_id), f"{{{rel_ns}}}id": rid},
        )
        next_sheet_id += 1

    content_root = ET.fromstring(entries[content_types_path])
    existing = {node.attrib.get("PartName") for node in content_root.findall(f"{{{content_type_ns}}}Override")}
    for _, path, _ in sheet_paths:
        part_name = "/" + path
        if part_name not in existing:
            ET.SubElement(
                content_root,
                f"{{{content_type_ns}}}Override",
                {
                    "PartName": part_name,
                    "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                },
            )

    entries[workbook_path] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)
    entries[rels_path] = ET.tostring(rel_root, encoding="utf-8", xml_declaration=True)
    entries[content_types_path] = ET.tostring(content_root, encoding="utf-8", xml_declaration=True)

    temp_path = xlsx_path.with_suffix(".tmp.xlsx")
    if temp_path.exists():
        temp_path.unlink()
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    shutil.move(str(temp_path), str(xlsx_path))


def load_conditions(path: Path) -> list[dict[str, str]]:
    if path.suffix.lower() == ".json":
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return [stringify_values(data)]
        if isinstance(data, list):
            return [stringify_values(item) for item in data]
        raise EmediError("JSON 조건 파일은 객체 또는 객체 배열이어야 합니다.")

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return [stringify_values(row) for row in csv.DictReader(f)]


def stringify_values(row: dict[str, Any]) -> dict[str, str]:
    return {str(k): "" if v is None else str(v) for k, v in row.items() if k}


def safe_filename(text: str) -> str:
    text = re.sub(r'[\\/:*?"<>|]+', "_", text).strip()
    text = re.sub(r"\s+", "_", text)
    return text[:100] or "emedi_result"


def describe_condition(tab: str, condition: dict[str, str]) -> str:
    fields = COMPANY_FIELDS if tab == "company" else ITEM_FIELDS
    parts = []
    for key, label in fields.items():
        value = condition.get(key, "").strip()
        if value and not (key == "indtyCd" and value == "1|2|21|22") and not (key == "grade" and value == "0"):
            parts.append(f"{key}-{value}")
    return safe_filename("_".join(parts))


def parse_key_values(values: list[str]) -> dict[str, str]:
    parsed: dict[str, str] = {}
    for value in values:
        if "=" not in value:
            raise EmediError(f"--set 값은 key=value 형식이어야 합니다: {value}")
        key, raw = value.split("=", 1)
        parsed[key.strip()] = raw.strip()
    return parsed


def print_fields() -> None:
    print("[품목검색 필드]")
    for key, label in ITEM_FIELDS.items():
        print(f"  {key}: {label}")
    print("\n[업체검색 필드]")
    for key, label in COMPANY_FIELDS.items():
        print(f"  {key}: {label}")


def run(args: argparse.Namespace) -> int:
    if args.list_fields:
        print_fields()
        return 0

    conditions = load_conditions(Path(args.conditions)) if args.conditions else [parse_key_values(args.set)]
    output_dir = Path(args.output_dir)
    session = build_session()

    for index, condition in enumerate(conditions, start=1):
        name = describe_condition(args.tab, condition)
        if len(conditions) > 1:
            name = f"{index:03d}_{name}"
        output_path = output_dir / f"{name}.xlsx"

        print(f"[{index}/{len(conditions)}] 검색 중: {condition}")
        if args.site_excel:
            result = search(session, args.tab, condition)
            print(f"  총 {result.total_count:,}건, 사이트 엑셀 요청 데이터 {len(result.excel_rows):,}건")
            if not args.dry_run:
                download_excel(session, result, output_path)
                print(f"  저장 완료: {output_path}")
        else:
            total_count, rows = collect_all_rows(session, args.tab, condition, args.delay, args.max_pages)
            print(f"  수집 완료: {len(rows):,}/{total_count:,}행")
            if not args.dry_run:
                params = default_company_params() if args.tab == "company" else default_item_params()
                params.update(condition)
                result = SearchResult(tab=args.tab, total_count=total_count, excel_rows=rows, params=params)
                download_excel(session, result, output_path)
                append_industry_sheets(output_path, rows)
                print(f"  저장 완료: {output_path}")
        if index < len(conditions):
            time.sleep(args.delay)

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="eMedi 의료기기 검색 결과 엑셀 다운로드 도구")
    parser.add_argument("--tab", choices=["item", "company"], default="item", help="검색 탭")
    parser.add_argument("--set", action="append", default=[], metavar="KEY=VALUE", help="검색 조건. 여러 번 지정 가능")
    parser.add_argument("--conditions", help="CSV 또는 JSON 조건 파일 경로")
    parser.add_argument("--output-dir", default="downloads", help="엑셀 저장 폴더")
    parser.add_argument("--delay", type=float, default=2.0, help="조건 여러 개 처리 시 요청 간격(초)")
    parser.add_argument("--max-pages", type=int, help="테스트용 최대 수집 페이지 수")
    parser.add_argument("--site-excel", action="store_true", help="사이트의 현재 화면 엑셀 다운로드 방식을 그대로 사용")
    parser.add_argument("--dry-run", action="store_true", help="검색만 하고 다운로드하지 않음")
    parser.add_argument("--list-fields", action="store_true", help="사용 가능한 검색 필드 출력")

    args = parser.parse_args()
    if not args.list_fields and not args.conditions and not args.set:
        parser.error("--set 또는 --conditions 중 하나가 필요합니다.")

    try:
        return run(args)
    except requests.RequestException as exc:
        print(f"네트워크 오류: {exc}", file=sys.stderr)
        return 2
    except EmediError as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

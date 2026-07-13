# [필수 패키지] 가상환경에서 한 번 설치: pip install pandas xlsxwriter openpyxl deep-translator ollama
# - xlsxwriter : 결과 xlsx 저장 및 셀 서식(파란색 강조)
# - openpyxl   : xlsx/xlsm/xls 입력 파일 읽기
# - ollama     : 로컬 LLM(원인·심각도 분석). Ollama 앱/서비스도 별도 실행 필요
import pandas as pd
from deep_translator import GoogleTranslator
import time
import re
import sys
import queue
import threading
from pathlib import Path

try:
    import ollama
except Exception:
    ollama = None

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, scrolledtext, messagebox
except Exception:
    tk = None
    ttk = filedialog = scrolledtext = messagebox = None

# UI 관련 코드는 건들지 말기
# 그 쪽은 잘 모름...
# AI 도움 받아서만 건들자

def output_path_from_input(input_file: str) -> str:
    """
    🌟[경로 생성 함수]
    원본 파일 경로를 받아, 원본 파일명 뒤에 '_arranged'를 붙인 새로운 저장 경로를 반환합니다.
    예: C:/data/test.xlsx -> C:/data/test_arranged.xlsx
    """
    p = Path(input_file)
    return str(p.parent / f"{p.stem}_arranged{p.suffix}")


def pms_column_width(column_name: str) -> int:
    widths = {
        "Class / Type": 9,
        "Product Classification": 18,
        "Product Name": 18,
        "Manufacturer": 20,
        "Event Date": 17,
        "Product Code": 8,
        "Device Problem": 18,
        "Patient Problem": 18,
        "Source": 16,
        "Link": 24,
        "Event Description": 48,
        "Event Description (KO)": 58,
        "원인": 14,
        "기기": 17,
        "환자 상태": 18,
        "사용성 관련성": 14,
        "심각도 점수": 14,
        "개선 방안": 18,
    }
    return widths.get(str(column_name), 16)


def normalize_severity_score(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None

    direct_labels = {
        "무시 가능": 1,
        "무시가능": 1,
        "negligible": 1,
        "경미": 2,
        "minor": 2,
        "심각/중대": 3,
        "심각": 3,
        "중대": 3,
        "serious": 3,
        "major": 3,
        "위독": 4,
        "critical": 4,
        "파국적/치명적": 5,
        "파국적": 5,
        "치명적": 5,
        "catastrophic": 5,
        "fatal": 5,
    }
    lowered = text.lower()
    for label, score in direct_labels.items():
        if label.lower() in lowered:
            return score

    match = re.search(r"[1-5]", text)
    if match:
        return int(match.group())
    return None


def estimate_severity_by_criteria(text) -> int | None:
    if pd.isna(text):
        return None
    lowered = str(text).lower()
    if not lowered.strip():
        return None

    no_harm_patterns = [
        "no clinical signs",
        "no symptoms",
        "no patient injury",
        "no patient involvement",
        "no impact",
        "no adverse event",
        "no harm",
        "not injured",
        "정상",
        "증상 없음",
        "상태 없음",
        "환자 위해 없음",
        "부상 없음",
        "영향 없음",
    ]
    fatal_negations = ["no death", "not result in death", "did not result in death", "사망 없음"]
    fatal_terms = ["death", "died", "fatality", "patient expired", "사망"]
    critical_terms = [
        "cardiac arrest",
        "asystole",
        "respiratory failure",
        "life threatening",
        "permanent injury",
        "permanent impairment",
        "irreversible",
        "심정지",
        "호흡부전",
        "생명 위협",
        "영구",
        "돌이킬 수 없는",
    ]
    serious_terms = [
        "medical intervention",
        "surgical intervention",
        "hospitalization",
        "required treatment",
        "treatment required",
        "injury",
        "disability",
        "의학적 개입",
        "외과적 개입",
        "입원",
        "치료 필요",
        "장애",
    ]
    minor_terms = [
        "minor",
        "temporary injury",
        "bruise",
        "abrasion",
        "redness",
        "discomfort",
        "irritation",
        "경미",
        "일시적 부상",
        "타박상",
        "찰과상",
        "불편",
        "자극",
    ]

    if any(term in lowered for term in fatal_terms) and not any(term in lowered for term in fatal_negations):
        return 5
    if any(term in lowered for term in critical_terms):
        return 4
    if any(term in lowered for term in no_harm_patterns):
        return 1
    if any(term in lowered for term in serious_terms):
        return 3
    if any(term in lowered for term in minor_terms):
        return 2
    return None


def apply_pms_excel_style(workbook, worksheet, dataframe: pd.DataFrame) -> dict[str, object]:
    header_format = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#D9D9D9",
            "border": 1,
            "border_color": "#000000",
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "font_name": "Calibri",
            "font_size": 10,
        }
    )
    body_format = workbook.add_format(
        {
            "border": 1,
            "border_color": "#000000",
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "font_name": "Calibri",
            "font_size": 10,
        }
    )
    text_body_format = workbook.add_format(
        {
            "border": 1,
            "border_color": "#000000",
            "align": "left",
            "valign": "top",
            "text_wrap": True,
            "font_name": "Calibri",
            "font_size": 10,
        }
    )
    blue_text_format = workbook.add_format(
        {
            "font_color": "#0070C0",
            "border": 1,
            "border_color": "#000000",
            "align": "left",
            "valign": "top",
            "text_wrap": True,
            "font_name": "Calibri",
            "font_size": 10,
        }
    )

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(dataframe), 1), max(len(dataframe.columns) - 1, 0))
    worksheet.set_row(0, 34, header_format)

    for col_idx, col_name in enumerate(dataframe.columns):
        worksheet.write(0, col_idx, col_name, header_format)
        width = pms_column_width(str(col_name))
        col_format = text_body_format if str(col_name) in {"Link", "Event Description", "Event Description (KO)"} else body_format
        worksheet.set_column(col_idx, col_idx, width, col_format)

    for row_idx in range(1, len(dataframe) + 1):
        worksheet.set_row(row_idx, 118)
        for col_idx, col_name in enumerate(dataframe.columns):
            value = dataframe.iloc[row_idx - 1, col_idx]
            if pd.isna(value):
                value = ""
            cell_format = text_body_format if str(col_name) in {"Link", "Event Description", "Event Description (KO)"} else body_format
            worksheet.write(row_idx, col_idx, value, cell_format)

    return {
        "header": header_format,
        "body": body_format,
        "text_body": text_body_format,
        "blue_text": blue_text_format,
    }


def run_arrangement(
        input_file: str,
        output_file: str,
        target_worksheets: list,
        ollama_model_name: str,
        main_progress_cb=None,
        llm_progress_cb=None
):
    """
    🌟[핵심 데이터 처리 함수]
    콜백 함수(main_progress_cb, llm_progress_cb)를 도입하여,
    로그 창에 퍼센트를 출력하는 대신 백그라운드에서 직접 프로그레스 바로 진행도를 전달합니다.
    """
    

    # 🌟분석 대상이 되는 열 이름(대소문자 및 띄어쓰기 변형 고려)
    target_descriptions = [
        "Event Description",
        "Event description",
        "event description",
        "Event Text"
    ]

    # 구글 번역기 초기화
    translator = GoogleTranslator(source="en", target="ko")
    translation_cache = {}  # 동일한 단어/문장의 중복 번역을 막기 위한 캐시 딕셔너리

    def translate_safe(text):
        """
        🌟[안전한 번역 함수]
        결측치(NaN)나 빈 문자열을 처리하며, 세미콜론(;)으로 구분된 텍스트를 개별 번역 후 다시 합칩니다.
        API 호출 제한을 피하기 위해 0.2초의 대기 시간(time.sleep)과 캐시를 사용합니다.
        """
        if pd.isna(text) or str(text).strip() == "":
            return text

        parts = [p.strip() for p in str(text).split(";")]
        translated_parts = []

        for part in parts:
            if part == "":
                continue

            # 이미 번역한 적 있는 텍스트면 캐시에서 바로 가져옴
            if part in translation_cache:
                translated_parts.append(translation_cache[part])
            else:
                try:
                    translated_text = translator.translate(part)
                    translation_cache[part] = translated_text
                    translated_parts.append(translated_text)
                    time.sleep(0.2)  # 대기 시간
                except Exception as e:
                    print(f"[{part}] 번역 중 오류 발생: {e}")
                    translated_parts.append(part)

        return "; ".join(translated_parts)

    def get_cause_from_llm(event_desc):
        """
        🌟[LLM 원인 분석 함수]
        Ollama(로컬 LLM)를 호출하여 텍스트로 된 이벤트 설명을 분석하고, 문제의 핵심 원인을 1~3단어 명사형으로 추출합니다.
        """
        if pd.isna(event_desc) or str(event_desc).strip() == "":
            return None
        if ollama is None:
            return "알 수 없음"  # pip install ollama 후 Ollama 서비스 실행 필요

        try:
            # AI에게 역할과 응답 규칙을 부여하는 프롬프트 설정
            response = ollama.chat(
                model=ollama_model_name,
                messages=[
                    {
                        "role": "system",
                        "content": """
                        당신은 유능한 의료기기 이상 사례 분석 전문가입니다.
                        제공된 한글 설명을 읽고, 문제의 핵심 원인을 '원인 불명' '하드웨어 오류', '소프트웨어 오류', '사용 오류' 반드시 이 네가지로만 답하세요.
                       문제의 핵심 원인은 이상사례가 발생했을 때 근본적인 원인이 된 것들입니다. 
                       문제의 핵심 원인은 1개만 답변할 수 있습니다.
                       원인 판단의 우선 순위는 다음과 같습니다: 원인 불명 >  하드웨어 오류 = 소프트웨어 오류 > 사용 오류
                       문제의 핵심 원인을 파악하기 어려운 경우엔 '원인 불명'로 답하세요.
                       문제의 핵심 원인이 스피커 결함, 메인보드 고장, 회로기판 고장, 부품 노후화 등 부품과 관련된 것이라고 판단 될 경우엔 '하드웨어 오류'로 답하세요.
                       문제의 핵심 원인이 소프트웨어 관련 오류로 판단 될 경우 답변은 '소프트웨어 오류'로 답하세요. 
                       문제의 핵심 원인이 사용자의 부적절한 사용, 사용자의 부주의, 사용자의 실수, 사용자의 미숙함, 사용자의 오해 등 사용자의 행동이 원인이라고 판단된다면 반드시 '사용 오류 : 문제 원인'로 답하세요.
                       예시) 사용 오류 : 사용자 부주의, 사용 오류: 사용자 교육 부족, 사용 오류 : 사용자 오해
                       
                       앞서 규정한 4개의 답변 방식 이외의 답변은 절대 불가합니다.
                       "원인 불명 하드웨어 오류" 같이 조합된 답변은 절대로 불가합니다.

                       """,
                    },  # 🌟프롬프트
                    {"role": "user", "content": str(event_desc)},
                ],
                options={"temperature": 0.1,
                         "num_predict": 8},  # 일관된 답변을 위해 온도를 0에 가깝게 설정, 생성 글자수를 8 토큰으로 제한
            )
            return response["message"]["content"].strip()
        except Exception as e:
            print(f"LLM(Ollama) 오류 발생: {e}")
            return "알 수 없음"

    def get_severity_from_llm(event_desc, model_name):
        """Ollama API를 이용해 첨부된 기준표 바탕으로 심각도(1~5)를 평가하는 함수"""
        if pd.isna(event_desc) or str(event_desc).strip() == "":
            return None
        if ollama is None:
            return None

        try:
            # 🌟 심각도 기준 표를 바탕으로 작성된 프롬프트
            system_prompt = """
            당신은 유능한 의료기기 이상 사례 심각도(Severity) 평가 전문가입니다.
            다음 [심각도 기준]에 따라 제공된 이벤트 설명의 심각도를 평가하고, 오직 1, 2, 3, 4, 5 중 하나의 '숫자'로만 답변하세요. 부가 설명은 절대 금지입니다.

            [심각도 기준]
            5 (파국적/치명적): 기기 문제로 환자의 사망을 초래한 경우.
            4 (위독): 기기 문제로 영구적 손상이나 돌이킬 수 없는 부상을 초래한 경우.
            3 (심각/중대): 기기 문제로 의학적 또는 외과적 개입이 필요한 부상 또는 장애를 초래한 경우.
            2 (경미): 기기 문제로 의학적 또는 외과적 개입이 필요하지 않은 일시적인 부상 또는 손상을 초래한 경우.
            1 (무시 가능): 기기 문제로 불편 또는 일시적 곤란만 초래했거나, 환자 위해가 없는 경우.
            
            심각도 5는 사망이 명시된 경우에만 부여하세요.
            심정지, 호흡부전, 생명 위협, 영구 손상은 사망이 명시되지 않았다면 4로 평가하세요.
            심각도를 평가할 때는 기기 문제로 인해 발생한 환자의 피해만 고려하세요.
            환자 피해가 없거나 임상 징후/증상이 없다고 명시된 경우 1로 평가하세요.
            예시1: 수술 도중 기기의 고장으로 인해 수술이 중지되었습니다. 환자에게 발생한 부작용은 없습니다. = 1
            예시2: 기기가 사용 도중 넘어져 환자와 충돌했고 환자는 경미한 타박상만 입었습니다. = 2
            예시3: 기기 고장으로 환자에게 치료가 필요한 손상이 발생했습니다. = 3
            예시4: 기기 고장으로 영구적인 신체 손상 또는 돌이킬 수 없는 부상이 발생했습니다. = 4
            예시5: 기기 고장으로 환자가 사망했습니다. = 5
     

            """

            response = ollama.chat(
                model=model_name,
                messages=[
                    {'role': 'system', 'content': system_prompt.strip()},
                    {'role': 'user', 'content': str(event_desc)}
                ],
                options={
                    'temperature': 0.0,  # 일관된 평가를 위해 0
                    'num_predict': 2  # 숫자 하나만 출력하므로 길이 최소화
                }
            )

            result = response['message']['content'].strip()

            # 응답에서 숫자만 확실하게 추출 (예: "정답은 3입니다" -> 3)
            match = re.search(r'\d+', result)
            if match:
                return normalize_severity_score(match.group())
            return normalize_severity_score(result)

        except Exception as e:
            print(f"LLM(Severity) 오류 발생: {e}")
            return None

    def get_improvement_from_llm(event_desc, cause, device_problem, model_name):
        """사용성 관련성이 있는 사례에 대해 한글 개선 방안을 생성합니다."""
        if pd.isna(event_desc) or str(event_desc).strip() == "":
            return "-"
        if ollama is None:
            return "사용자 교육, 사용 절차 및 사용 설명서 개선 검토"

        try:
            response = ollama.chat(
                model=model_name,
                messages=[
                    {
                        "role": "system",
                        "content": """
                        당신은 의료기기 PMS/사용성 리스크 개선 방안 전문가입니다.
                        제공된 이상사례 설명, 원인, 기기 문제를 바탕으로 사용성 관련성이 있는 경우의 개선 방안을 한국어로 작성하세요.
                        답변은 한 문장으로만 작성하고, 교육/표시사항/사용 절차/알람/사용자 인터페이스 개선 중 가장 적절한 조치를 구체적으로 제안하세요.
                        과장하지 말고, 실제 PMS 엑셀에 들어갈 수 있는 짧은 실무 문장으로 답하세요.
                        """,
                    },
                    {
                        "role": "user",
                        "content": f"이상사례 설명: {event_desc}\n원인: {cause}\n기기 문제: {device_problem}",
                    },
                ],
                options={"temperature": 0.1, "num_predict": 80},
            )
            return response["message"]["content"].strip()
        except Exception as e:
            print(f"LLM(개선 방안) 오류 발생: {e}")
            return "사용자 교육, 사용 절차 및 사용 설명서 개선 검토"

    print("\n데이터 파일을 불러오는 중입니다...")
    all_sheets = pd.read_excel(input_file, sheet_name=None)

    # 전체 시트 개수 기반으로 프로그레스 청크 계산 (전체 90% 비중)
    total_sheets = len(target_worksheets)
    chunk = 90.0 / total_sheets if total_sheets > 0 else 90.0
    current_sheet_idx = 0

    # 사용자가 입력한 목표 워크시트들만 순회하며 작업 수행
    for ws_name in target_worksheets:
        base_prog = current_sheet_idx * chunk
        if main_progress_cb: main_progress_cb(base_prog)  # UI 관련 코드

        if ws_name not in all_sheets:
            print(f"\n⚠️ 경고: '{ws_name}' 시트가 파일 내에 존재하지 않아 건너뜁니다.")
            current_sheet_idx += 1
            continue

        print(f"\n=======================================================")
        print(f"▶ ['{ws_name}'] 시트 데이터 처리를 시작합니다...")
        print(f"=======================================================")

        df = all_sheets[ws_name]
        # 열 이름의 불필요한 공백 제거
        df.columns = [" ".join(str(col).split()) for col in df.columns]

        # =========================================================================
        # 🌟 [추가된 부분] 분석 프로세스 전 필수 열(Column) 미리 생성
        # =========================================================================
        df = df.rename(
            columns={
                "Event Text KO": "Event Description (KO)",
                "Event Description(KO)": "Event Description (KO)",
                "Event description(KO)": "Event Description (KO)",
                "Event Description KO": "Event Description (KO)",
                "Event description KO": "Event Description (KO)",
                "환자상태": "환자 상태",
                "개선방안": "개선 방안",
            }
        )
        normalized_columns = []
        for col in df.columns:
            col_text = str(col).strip()
            compact = re.sub(r"[\s_]+", "", col_text).lower()
            if compact in {"eventdescription(ko)", "eventdescriptionko", "eventtextko"}:
                normalized_columns.append("Event Description (KO)")
            else:
                normalized_columns.append(col)
        df.columns = normalized_columns
        df = df.loc[:, ~df.columns.duplicated()]
        required_columns = ["Event Description (KO)", "원인", "기기", "환자 상태", "사용성 관련성", "심각도 점수", "개선 방안"]
        for col in required_columns:
            if col not in df.columns:
                df[col] = None  # 열이 없다면 빈칸(None)으로 새롭게 만들어 둡니다.
            df[col] = df[col].astype("object")

        # 'Event Text' 관련 열 찾기
        description = next(
            (c for c in df.columns if c.lower() in [d.lower() for d in target_descriptions]), None
        )

        if not description:
            print(f"⚠️ 경고: '{ws_name}' 시트에 이벤트 설명 열이 없습니다. 관련 기능을 건너뜁니다.")

        print("\n1. 'Patient Problem' 및 'Device Problem' 번역 중...")
        if main_progress_cb: main_progress_cb(base_prog + 0.1 * chunk)  # UI 관련 코드

        # 🌟 환자 상태 및 기기 문제 열을 찾아 한국어로 번역하여 새 열에 할당
        if description:
            df["Event Description (KO)"] = df[description].apply(translate_safe)

        patient_prb = next((c for c in df.columns if c.lower() == "patient problem"), None)
        if patient_prb:
            df["환자 상태"] = df[patient_prb].apply(translate_safe)

        device_prb = next((c for c in df.columns if c.lower() == "device problem"), None)
        if device_prb:
            df["기기"] = df[device_prb].apply(translate_safe)

        # ---------------------------------------------------------
        # 🌟 [기능 2 & 3] 키워드 기반 원인 기입
        # ---------------------------------------------------------
        print("\n 작업 완료\n")
        print("\n2. 키워드 기반 '원인' 1차 기입 중...")
        if main_progress_cb: main_progress_cb(base_prog + 0.5 * chunk)  # UI 관련 코드
        if llm_progress_cb: llm_progress_cb(100.0)  # UI 관련 코드

        # 🌟설명 열이 존재할 경우, 특정 키워드들을 검색하여 사전 분류 진행
        if description and description in df.columns:
            if "원인" not in df.columns:
                df["원인"] = None

            df["원인"] = df["원인"].astype("object")
            df["사용성 관련성"] = df["사용성 관련성"].astype("object")
            df["원인"] = df["원인"].replace({"분석 실패": "알 수 없음"})
            empty_cause = df["원인"].isna() | (df["원인"].astype(str).str.strip() == "")

            # 2) 원인 불명 관련 키워드 및 텍스트 길이(270자 이하) 기반 분류
            phrases_unknown = [
                "정보 부족", "현재 조사 중 입니다", "원인은 알 수 없", "정상임을 확인",
                "완료되는 대로", "정보는 제공되지 않", "원인을 파악할 수 없",
                "불만 사항은 확인되지 않았", "정상적으로 가동되고", "근본 원인을 알 수 없",
                "근본 원인을 파악할 수 없", "관련 정보를 확보할 수 없어"
            ]
            cond_unknown_keyword = df[description].astype(str).str.contains("|".join(phrases_unknown), na=False,
                                                                            regex=True)
            cond_unknown_length = df[description].astype(str).str.len() <= 270
            df.loc[empty_cause & (cond_unknown_keyword | cond_unknown_length), "원인"] = "원인 불명"

        # ---------------------------------------------------------
        # 🌟 [기능 2 & 3] 키워드 기반 원인 기입
        # ---------------------------------------------------------

        print("\n 작업 완료\n")
        print("\n3. '사용성 관련성' 1차 기입 중...")
        if main_progress_cb: main_progress_cb(base_prog + 0.6 * chunk)  # UI 관련 코드

        # 🌟설명 열이 존재할 경우, 특정 키워드들을 검색하여 사전 분류 진행
        if description and description in df.columns:

            if "사용성 관련성" not in df.columns:
                df["사용성 관련성"] = None

            df["원인"] = df["원인"].astype("object")
            df["사용성 관련성"] = df["사용성 관련성"].astype("object")
            empty_cause = df["원인"].isna() | (df["원인"].astype(str).str.strip() == "")

            # 1) 사용자 오류 관련 키워드 기반 분류
            phrases_user = [
                "사용자 과실", "사용 오류", "사용자 문제", "사용자 오류", "재교육",
                "경보를 놓쳤을", "기능 사용/인지", "사용자 혼란", "교육 문제",
                "사용자의 지식 부족", "사용자 지식 부족", "실수로", "오해",
                "추가 사용자 교육이", "원인은 사용자였습"
            ]
            cond_user = df[description].astype(str).str.contains("|".join(phrases_user), na=False, regex=True)
            df.loc[cond_user, "사용성 관련성"] = "O"
            df.loc[empty_cause & cond_user, "원인"] = "사용 오류"

        # ---------------------------------------------------------
        # 🌟 [기능 4] 조건에 따라 '심각도 점수' 기입
        # ---------------------------------------------------------
        print("\n 작업 완료\n")
        print("\n4. 'Event Type' 열 조건에 따라 '심각도 점수' 기입 중...")

        target_columns = ["Event Type"]

        # 🌟 Event Type의 위해 수준에 따라 기본 심각도 점수 부여
        for col in target_columns:
            if col in df.columns:
                df.loc[df[col] == "Death", "심각도 점수"] = 5
                df.loc[df[col] == "Injury", "심각도 점수"] = 3
        # ---------------------------------------------------------
        # 🌟 [기능 5] 특정 문구 포함 시 '심각도' 변경
        # ---------------------------------------------------------
        print("\n 작업 완료\n")
        print("\n5. 특정 문구 확인 후 '심각도' 변경 중...")
        # 🌟 고위험 키워드가 포함된 경우 기준표에 맞춰 심각도 점수를 보정
        if description and description in df.columns:
            phrases_fatal = ["Death", "Died", "Fatality", "사망", "치명적"]
            phrases_critical = [
                "Cardiac Arrest",
                "Asystole",
                "Respiratory Failure",
                "Life Threatening",
                "Permanent Injury",
                "Permanent Impairment",
                "Irreversible",
                "심정지",
                "호흡부전",
                "생명 위협",
                "영구",
                "돌이킬 수 없는",
                "위독",
            ]
            desc_text = df[description].astype(str)
            cond_fatal = desc_text.apply(
                lambda text: any(term.lower() in text.lower() for term in phrases_fatal)
                and not any(term in text.lower() for term in ["no death", "not result in death", "did not result in death", "사망 없음"])
            )
            cond_critical = desc_text.str.contains("|".join(phrases_critical), na=False, regex=True)
            if "심각도 점수" not in df.columns:
                df["심각도 점수"] = None
            df["심각도 점수"] = df["심각도 점수"].astype("object")
            df.loc[cond_critical, "심각도 점수"] = 4
            df.loc[cond_fatal, "심각도 점수"] = 5


        # ---------------------------------------------------------
        # 🌟 [기능 6] LLM 기반 '심각도 점수' 지능형 분석 및 기입
        # ---------------------------------------------------------
        print("\n 작업 완료\n")
        print("\n6. LLM을 이용해 비어있는 '심각도 점수'를 분석 및 기입 중...")
        if main_progress_cb: main_progress_cb(base_prog + 0.2 * chunk) # UI 관련 코드
        if llm_progress_cb: llm_progress_cb(0.0) # UI 관련 코드

        if description and description in df.columns:
            if '심각도 점수' not in df.columns:
                df['심각도 점수'] = None
            df['심각도 점수'] = df['심각도 점수'].astype('object')
            df['심각도 점수'] = df['심각도 점수'].apply(normalize_severity_score)
            empty_before_llm = df['심각도 점수'].isna() | (df['심각도 점수'].astype(str).str.strip() == "")
            df.loc[empty_before_llm, '심각도 점수'] = df.loc[empty_before_llm, description].apply(estimate_severity_by_criteria)

            # --- [수정된 부분] 비어있는 '심각도 점수' 행만 필터링 ---
            empty_sev_mask = df['심각도 점수'].isna() | (df['심각도 점수'].astype(str).str.strip() == "")
            sev_total_to_analyze = empty_sev_mask.sum()
            current_sev_count = 0

            if sev_total_to_analyze > 0:
                print(f"   - AI 심각도 분석 진행 예정 건수: {sev_total_to_analyze} 건")

                # 비어있는 행만 순회하며 심각도를 LLM에게 물어봅니다.
                for idx in df[empty_sev_mask].index:
                    desc = df.at[idx, description]

                    # LLM 호출
                    extracted_sev = get_severity_from_llm(desc, ollama_model_name)
                    df.at[idx, '심각도 점수'] = normalize_severity_score(extracted_sev) or estimate_severity_by_criteria(desc)

                    current_sev_count += 1
                    if current_sev_count % 25 == 0:
                        print(f"   - AI 분석 진행 상황: {current_sev_count} / {sev_total_to_analyze} 완료")

                    # 직접 콜백을 호출하여 퍼센트 갱신 (더 이상 print 로그 출력하지 않음)
                    sev_completion_rate = 100 / sev_total_to_analyze * current_sev_count if sev_total_to_analyze else 100
                    if llm_progress_cb:
                        llm_progress_cb(sev_completion_rate)
            else:
                print("   - 이미 모든 심각도 점수가 기입되어 LLM 분석을 건너뜁니다.")
                if llm_progress_cb:
                    llm_progress_cb(100.0)

            df['심각도 점수'] = df['심각도 점수'].apply(normalize_severity_score)

        # ---------------------------------------------------------
        # 🌟 [기능 7] LLM(Ollama)을 이용한 지능형 원인 분석
        # ---------------------------------------------------------
        print("\n 작업 완료\n")
        print(f"\n7.LLM(Ollama - {ollama_model_name})을 이용해 비어있는 항목들의 '원인' 분석 중...")
        if main_progress_cb: main_progress_cb(base_prog + 0.7 * chunk) # UI 관련 코드
        if llm_progress_cb: llm_progress_cb(0.0) # UI 관련 코드

        # 🌟 키워드 필터링으로 걸러지지 않아 '원인'이 비어있는 행들을 LLM으로 분석
        if description and description in df.columns:
            empty_cause_mask = df["원인"].isna()
            total_to_analyze = empty_cause_mask.sum()
            current_count = 0

            if total_to_analyze > 0:
                print(f"   - AI 원인 분석 진행 예정 건수: {total_to_analyze} 건")

                for idx in df[empty_cause_mask].index:
                    desc = df.at[idx, description]

                    extracted_cause = get_cause_from_llm(desc)
                    df.at[idx, "원인"] = extracted_cause

                    current_count += 1
                    if current_count % 25 == 0:
                        print(f"   - AI 분석 진행 상황: {current_count} / {total_to_analyze} 완료")

                    # 직접 콜백을 호출하여 퍼센트 갱신
                    completion_rate = 100 / total_to_analyze * current_count if total_to_analyze else 100
                    if llm_progress_cb:
                        llm_progress_cb(completion_rate)

                    # 🌟 [추가] LLM이 '사용 오류'라고 판단하면 '사용성 관련성'에 'O' 자동 기입
                    if extracted_cause and '사용 오류' in extracted_cause:
                        df.at[idx, '사용성 관련성'] = 'O'
            else:
                print("   - 이미 모든 원인이 기입되어 LLM 분석을 건너뜁니다.")
                if llm_progress_cb:
                    llm_progress_cb(100.0)

            # 🌟 [최종 확인] 원인 열에 '사용 오류'가 포함된 모든 행의 사용성 관련성을 'O'로 확정
        if '원인' in df.columns:
            df['원인'] = df['원인'].replace({"분석 실패": "알 수 없음"})

        if '원인' in df.columns and '사용성 관련성' in df.columns:
            df.loc[df['원인'].astype(str).str.contains('사용 오류', na=False), '사용성 관련성'] = 'O'

        if '사용성 관련성' in df.columns:
            df['사용성 관련성'] = df['사용성 관련성'].astype('object')
            empty_usability = df['사용성 관련성'].isna() | (df['사용성 관련성'].astype(str).str.strip() == "")
            df.loc[empty_usability, '사용성 관련성'] = 'X'

        if description and description in df.columns and '개선 방안' in df.columns:
            print("\n8. 사용성 관련성이 있는 사례의 '개선 방안' 작성 중...")
            df['개선 방안'] = df['개선 방안'].astype('object')
            usability_mask = df['사용성 관련성'].astype(str).str.upper().eq('O')
            empty_improvement = df['개선 방안'].isna() | (df['개선 방안'].astype(str).str.strip() == "")
            improvement_targets = df[usability_mask & empty_improvement].index

            for idx in improvement_targets:
                df.at[idx, '개선 방안'] = get_improvement_from_llm(
                    df.at[idx, description],
                    df.at[idx, '원인'] if '원인' in df.columns else "",
                    df.at[idx, '기기'] if '기기' in df.columns else "",
                    ollama_model_name,
                )

            non_usability_empty = (~usability_mask) & empty_improvement
            df.loc[non_usability_empty, '개선 방안'] = "-"

        analysis_columns = ["Event Description (KO)", "원인", "기기", "환자 상태", "사용성 관련성", "심각도 점수", "개선 방안"]
        if description in df.columns:
            leading_columns = list(df.columns[: df.columns.get_loc(description) + 1])
            analysis_existing = [col for col in analysis_columns if col in df.columns]
            trailing_columns = [col for col in df.columns if col not in leading_columns + analysis_existing]
            df = df[leading_columns + analysis_existing + trailing_columns]

        all_sheets[ws_name] = df

        # ---------------------------------------------------------
        # 🌟 [기능 8] 파란색 서식 적용 및 엑셀 최종 저장
        # ---------------------------------------------------------
        print("\n 작업 완료\n")

        current_sheet_idx += 1

    print("\n=======================================================")
    print("8. 모든 시트의 작업을 마치고 최종 파일 저장 및 서식 적용 중...")

    # 마지막 저장 단계 진입 시 진행도 맞춤
    if main_progress_cb: main_progress_cb(90.0) # UI 관련 코드
    if llm_progress_cb: llm_progress_cb(100.0) # UI 관련 코드

    # xlsxwriter 엔진을 사용하여 엑셀 작성 (텍스트 서식 등을 제어하기 위함)
    try:
        import xlsxwriter  # noqa: F401
    except ImportError as e:
        raise ImportError(
            "xlsxwriter 패키지가 없습니다. 터미널에서: pip install xlsxwriter"
        ) from e

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        for sheet_name, sheet_df in all_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

        workbook = writer.book
        sheet_formats = {}
        for sheet_name, sheet_df in all_sheets.items():
            if sheet_name in writer.sheets:
                sheet_formats[sheet_name] = apply_pms_excel_style(workbook, writer.sheets[sheet_name], sheet_df)

        for ws_name in target_worksheets:
            if ws_name in writer.sheets:
                df_target = all_sheets[ws_name]
                formats = sheet_formats.get(ws_name) or apply_pms_excel_style(workbook, writer.sheets[ws_name], df_target)

                description = next(
                    (c for c in df_target.columns if c.lower() in [d.lower() for d in target_descriptions]),
                    None,
                )
                ko_description = "Event Description (KO)" if "Event Description (KO)" in df_target.columns else None

                # 특정 문자열 패턴에 색상을 입히는 로직
                if ko_description:
                    ws = writer.sheets[ws_name]
                    col_idx = df_target.columns.tolist().index(ko_description)

                    # 파란색으로 칠할 타겟 단어 리스트
                    color_target_words = [
                        "사용자 과실",
                        "사용자 문제",
                        "사용자 오류",
                        "재교육",
                        "경보를 놓쳤을",
                        "기능 사용/인지",
                        "사용자 혼란",
                        "교육 문제",
                        "정보 부족",
                        "보고된 문제의 원인은 알 수 없",
                        "정상임을 확인",
                        "조사가 완료되는 대로",
                        "정보는 제공되지 않았",
                        "정확한 원인을 파악할 수 없었습니다",
                        "근본 원인을 파악할 수 없었습니다.",
                        "근본 원인을 파악할 수 없습니다",
                        "근본 원인은 파악할 수 없습니다",
                        "근본 원인은 파악할 수 없었습니다",
                        "근본 원인은 아직",
                        "보고된 문제는 확인되지 않았으며",
                        "불만 사항은 확인되지 않았으며",
                        "문제의 원인을 파악할 수 없었",
                        "문제의 원인은 확인되지",
                        "조사 중입니다",
                        "원인을 확인할 수 없었",
                        "문제의 원인은 아직 확인되지 않았습",
                        "문제의 원인은"
                    ]
                    color_pattern = f"({'|'.join(re.escape(word) for word in color_target_words)})"
                    feature7_marker = "이 코드는 현재 무효인 상황"  # 이 문구 앞은 모두 파랗게

                    # 서식 객체 생성
                    blue_font = workbook.add_format({"font_color": "#0070C0"})
                    blue_wrap_format = formats["blue_text"]
                    wrap_format = formats["text_body"]

                    def build_highlight_pattern(row):
                        dynamic_words = list(color_target_words)
                        for col in ["원인", "기기"]:
                            if col in df_target.columns:
                                value = row.get(col)
                                if not pd.isna(value):
                                    for piece in re.split(r"[,;/\n]", str(value)):
                                        piece = piece.strip()
                                        if len(piece) >= 2 and piece not in {"X", "O", "-", "알 수 없음", "원인 불명"}:
                                            dynamic_words.append(piece)
                        keywords = [
                            "오류", "고장", "결함", "오작동", "문제", "불량", "실패", "부정확",
                            "측정", "스캔", "프로브", "센서", "알람", "배터리", "카테터",
                            "소프트웨어", "하드웨어", "사용자", "교육", "사용 오류",
                        ]
                        dynamic_words.extend(keywords)
                        dynamic_words = sorted(set(dynamic_words), key=len, reverse=True)
                        return f"({'|'.join(re.escape(word) for word in dynamic_words)})"

                    def sentence_chunks(text):
                        return re.split(r"([.!?。！？]\s+|\n+)", text)

                    # 셀 단위로 텍스트를 검사하여 서식을 입힘 (Rich String)
                    for row_idx, val in enumerate(df_target[ko_description], start=1):
                        if pd.isna(val):
                            continue

                        text_str = str(val)
                        rich_parts = []
                        row_data = df_target.iloc[row_idx - 1]
                        row_pattern = build_highlight_pattern(row_data)

                        marker_idx = text_str.find(feature7_marker)

                        if marker_idx != -1:
                            before_text = text_str[:marker_idx]
                            after_text = text_str[marker_idx:]
                        else:
                            before_text = ""
                            after_text = text_str

                        # xlsxwriter write_rich_string 순서: 문자열, 서식, 문자열, 서식, ...
                        if before_text:
                            rich_parts.extend([blue_font, before_text])

                        if after_text:
                            for sentence in sentence_chunks(after_text):
                                if not sentence:
                                    continue
                                if re.search(row_pattern, sentence):
                                    rich_parts.extend([blue_font, sentence])
                                else:
                                    chunks = re.split(color_pattern, sentence)
                                    for c in chunks:
                                        if not c:
                                            continue
                                        if c in color_target_words:
                                            rich_parts.extend([blue_font, c])
                                        else:
                                            rich_parts.append(c)

                        # 하나의 셀 안에 여러 서식이 섞여 있다면 write_rich_string 사용
                        if len(rich_parts) == 2 and not isinstance(rich_parts[0], str) and isinstance(rich_parts[1], str):
                            ws.write_string(row_idx, col_idx, rich_parts[1], blue_wrap_format)
                        elif len(rich_parts) > 1:
                            ws.write_rich_string(row_idx, col_idx, *rich_parts, wrap_format)
                        elif len(rich_parts) == 1:
                            if isinstance(rich_parts[0], str):
                                ws.write_string(row_idx, col_idx, rich_parts[0], wrap_format)

    print(
        f"\n모든 작업이 성공적으로 완료되었습니다. 파일이 정상적으로 열릴 것입니다. \n결과 파일 이름: {output_file}"
    )
    if main_progress_cb: main_progress_cb(100.0) # UI 관련 코드


# 🌟이하는 전부 AI의 도움을 받아 유지 보수하기 난감
class QueueWriter:
    """
    [GUI 로그 연동용 클래스]
    표준 출력(print)을 가로채서 GUI의 큐(queue)로 넘겨줍니다.
    이렇게 하면 백그라운드 스레드에서 발생하는 print문이 메인 GUI의 로그 창에 나타나게 됩니다.
    """

    def __init__(self, q: queue.Queue):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(("log", s))

    def flush(self):
        pass


def launch_gui():
    """
    [UI 구성 및 메인 루프 함수]
    Tkinter를 이용하여 데스크톱 GUI 창을 구성하고, 사용자 입력을 받아 데이터 처리를 실행합니다.
    """
    if tk is None:
        raise RuntimeError("Tkinter GUI를 사용할 수 없는 환경입니다. Streamlit 앱에서는 PMS 탭을 사용하세요.")

    root = tk.Tk()
    root.title("PMS Auto Arrangement (Local Computation Ver.)")
    root.minsize(720, 650)

    msg_q: queue.Queue = queue.Queue()  # 스레드 간 안전한 메시지 전달을 위한 큐
    worker_running = threading.Event()  # 작업이 현재 실행 중인지 상태를 관리

    main = ttk.Frame(root, padding=12)
    main.pack(fill=tk.BOTH, expand=True)

    # 🌟안내 문구 영역
    intro = (
        "이 버전은 Ollama를 로컬 PC에서 사용합니다.\n"
        "Ollama와 LLM 모델이 설치되어 있어야 합니다.\n"
        "컴퓨터의 사양에 따라서 작업 시간의 차이가 큽니다.\n"
        "자세한 사항은 반드시 READ ME 파일을 확인해주세요."
    )
    ttk.Label(main, text=intro, wraplength=680, justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 8))

    # 🌟LLM 모델명 입력 영역
    model_row = ttk.Frame(main)
    model_row.pack(fill=tk.X, pady=4)
    ttk.Label(model_row, text="LLM 모델명:", width=14).pack(side=tk.LEFT)
    model_var = tk.StringVar(value="gemma3:4b") # 모델 변경 시 변경
    ttk.Entry(model_row, textvariable=model_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)

    # 파일 선택 영역
    file_row = ttk.Frame(main)
    file_row.pack(fill=tk.X, pady=4)
    ttk.Label(file_row, text="입력 파일:", width=14).pack(side=tk.LEFT)
    input_var = tk.StringVar()
    ttk.Entry(file_row, textvariable=input_var, width=42).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

    out_label_var = tk.StringVar(value="저장 경로: (입력 파일을 선택하면 자동 설정)")

    def pick_file():
        # 파일 탐색기를 열어 경로를 반환받는 함수
        path = filedialog.askopenfilename(
            title="작업할 엑셀 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")],
        )
        if path:
            input_var.set(path)
            out_label_var.set(f"저장 경로(자동): {output_path_from_input(path)}")

    ttk.Button(file_row, text="찾아보기…", command=pick_file).pack(side=tk.LEFT)

    ttk.Label(main, textvariable=out_label_var, wraplength=680, justify=tk.LEFT).pack(anchor=tk.W, pady=(2, 8))

    # 타겟 워크시트 입력 영역
    ws_row = ttk.Frame(main)
    ws_row.pack(fill=tk.X, pady=4)
    ttk.Label(ws_row, text="목표 워크시트:", width=14).pack(side=tk.LEFT, anchor=tk.N)
    ws_var = tk.StringVar(value="Sheet1")
    ttk.Entry(ws_row, textvariable=ws_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
    ttk.Label(
        main,
        text="쉼표(,)로 구분하여 시트 이름을 입력하세요.",
        foreground="gray",
    ).pack(anchor=tk.W)

    # 진행 상태가 표시되는 로그 출력 영역
    log_frame = ttk.LabelFrame(main, text="진행 로그", padding=6)
    log_frame.pack(fill=tk.BOTH, expand=True, pady=(12, 8))
    log_text = scrolledtext.ScrolledText(log_frame, height=14, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 9))
    log_text.pack(fill=tk.BOTH, expand=True)

    # --- 프로그레스 바 UI 영역 (전체 진행도 누적 & LLM 진행도) ---
    progress_container = ttk.Frame(main)
    progress_container.pack(fill=tk.X, pady=(4, 8))

    # 1) 메인 (전체 누적) 진행도
    main_frame = ttk.Frame(progress_container)
    main_frame.pack(fill=tk.X, pady=(0, 4))
    ttk.Label(main_frame, text="전체 진행도:", width=14).pack(side=tk.LEFT)
    main_progress_var = tk.DoubleVar(value=0.0)
    main_progress_bar = ttk.Progressbar(main_frame, variable=main_progress_var, maximum=100)
    main_progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
    main_percent_label_var = tk.StringVar(value="0.0%")
    ttk.Label(main_frame, textvariable=main_percent_label_var, width=7, anchor="e").pack(side=tk.LEFT, padx=(4, 0))

    # 2) 서브 (LLM) 진행도
    llm_frame = ttk.Frame(progress_container)
    llm_frame.pack(fill=tk.X)
    ttk.Label(llm_frame, text="LLM 분석 진행도:", width=14).pack(side=tk.LEFT)
    llm_progress_var = tk.DoubleVar(value=0.0)
    llm_progress_bar = ttk.Progressbar(llm_frame, variable=llm_progress_var, maximum=100)
    llm_progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
    llm_percent_label_var = tk.StringVar(value="0.0%")
    ttk.Label(llm_frame, textvariable=llm_percent_label_var, width=7, anchor="e").pack(side=tk.LEFT, padx=(4, 0))
    # -------------------------------------------------------------

    # 시작 버튼
    btn_row = ttk.Frame(main)
    btn_row.pack(fill=tk.X)
    run_btn = ttk.Button(btn_row, text="작업 시작")
    run_btn.pack(side=tk.LEFT)

    def append_log(chunk: str):
        # 텍스트 박스에 로그를 추가하는 헬퍼 함수
        log_text.configure(state=tk.NORMAL)
        log_text.insert(tk.END, chunk)
        log_text.see(tk.END)  # 새 로그가 추가될 때마다 가장 아래로 스크롤
        log_text.configure(state=tk.DISABLED)

    def poll_queue():
        """
        [GUI 타이머 루프]
        주기적(120ms)으로 큐를 확인하여 백그라운드 스레드에서 넘어온 로그 메시지,
        진행률 수치, 종료 신호를 각각 목적에 맞게 처리합니다.
        """
        try:
            while True:
                kind, data = msg_q.get_nowait()

                # 로그 메시지 출력
                if kind == "log":
                    append_log(data)

                # 정석적인 방법: 직접 수치를 받아 프로그레스 바 갱신
                elif kind == "prog_main":
                    main_progress_var.set(data)
                    main_percent_label_var.set(f"{data:.1f}%")
                elif kind == "prog_llm":
                    llm_progress_var.set(data)
                    llm_percent_label_var.set(f"{data:.1f}%")

                elif kind == "done_ok":
                    append_log(data + "\n")
                    messagebox.showinfo("완료", data.strip())
                    run_btn.configure(state=tk.NORMAL)
                    worker_running.clear()
                elif kind == "done_err":
                    append_log(data + "\n")
                    messagebox.showerror("오류", data.strip())
                    run_btn.configure(state=tk.NORMAL)
                    worker_running.clear()
        except queue.Empty:
            pass
        root.after(120, poll_queue)

    def do_run():
        """
        [작업 실행 런처]
        사용자 입력값 유효성을 검사한 뒤 별도의 스레드를 생성하여 데이터 처리(run_arrangement)를 시작합니다.
        메인 스레드에서 직접 무거운 작업을 돌리면 창(UI)이 멈추기(프리징) 때문입니다.
        """
        if worker_running.is_set():
            return

        # 1. 입력값 검증 (경로, 모델명, 시트명 등)
        inp = input_var.get().strip()
        if not inp:
            messagebox.showwarning("입력 필요", "작업할 파일을 선택하거나 경로를 입력하세요.")
            return
        if not Path(inp).is_file():
            messagebox.showerror("파일 없음", f"파일을 찾을 수 없습니다:\n{inp}")
            return
        raw_ws = ws_var.get().strip()
        if not raw_ws:
            messagebox.showwarning("입력 필요", "목표 워크시트 이름을 입력하세요.")
            return
        target_worksheets = [ws.strip() for ws in raw_ws.split(",") if ws.strip()]
        if not target_worksheets:
            messagebox.showwarning("입력 필요", "유효한 워크시트 이름이 없습니다.")
            return
        model = model_var.get().strip()
        if not model:
            messagebox.showwarning("입력 필요", "LLM 모델명을 입력하세요.")
            return

        out = output_path_from_input(inp)
        out_label_var.set(f"저장 경로(자동): {out}")

        run_btn.configure(state=tk.DISABLED)  # 작업 중 중복 실행 방지
        worker_running.set()

        # 로그 초기화
        log_text.configure(state=tk.NORMAL)
        log_text.delete("1.0", tk.END)
        log_text.configure(state=tk.DISABLED)

        # --- 시작 시 모든 진행 상황 및 누적 변수 초기화 ---
        main_progress_var.set(0.0)
        main_percent_label_var.set("0.0%")
        llm_progress_var.set(0.0)
        llm_percent_label_var.set("0.0%")

        def worker():
            """
            [백그라운드 작업 스레드]
            콜백 함수(cb_main, cb_llm)를 정의하여 run_arrangement에 넘겨줍니다.
            진행 수치는 이 함수들을 통해 GUI 큐(msg_q)로 직접 전달됩니다.
            """
            old_out = sys.stdout
            sys.stdout = QueueWriter(msg_q)

            # 진행도 전달을 위한 콜백 함수
            def cb_main(val):
                msg_q.put(("prog_main", val))

            def cb_llm(val):
                msg_q.put(("prog_llm", val))

            try:
                # 콜백 인자를 포함하여 메인 처리 함수 실행
                run_arrangement(
                    inp,
                    out,
                    target_worksheets,
                    model,
                    main_progress_cb=cb_main,
                    llm_progress_cb=cb_llm
                )
                msg_q.put(("done_ok", f"저장 완료:\n{out}"))
            except Exception as e:
                msg_q.put(("done_err", f"프로그램 실행 중 오류:\n{e}"))
            finally:
                sys.stdout = old_out  # 작업 완료 후 표준 출력 복구

        threading.Thread(target=worker, daemon=True).start()

    run_btn.configure(command=do_run)
    poll_queue()  # 큐 타이머 시작
    root.mainloop()  # GUI 실행 루프


if __name__ == "__main__":
    launch_gui()

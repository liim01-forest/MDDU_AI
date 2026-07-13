# 의료기기 인허가 정보 수집기

식약처, FDA, EUDAMED 인허가 정보를 한 화면에서 조회하고 엑셀로 정리하기 위한 Streamlit 앱 골격입니다.

현재 버전은 UI와 엑셀 출력 구조를 먼저 검증하기 위한 샘플 데이터 모드입니다. 실제 사이트 연동은 `app.py`의 `sample_mfds`, `sample_fda`, `sample_eudamed` 함수를 수집기 함수로 교체하면 됩니다.

## 실행

```powershell
pip install -r requirements.txt
streamlit run app.py
```

기존 Streamlit 앱과 겹치지 않도록 이 프로젝트는 기본 포트를 `8510`으로 설정했습니다.
실행 후 아래 주소로 접속하면 됩니다.

```text
http://localhost:8510
```

## 구성

- 식약처 탭: 품목정보, 업체정보, 허가정보 조회 화면
- FDA 탭: 510(k), PMA, De Novo, Registration & Listing, Product Classification, AccessGUDID 선택 화면
- EUDAMED 탭: UDI/Device, SRN, Risk Class, Certificate 중심 조회 화면
- 통합 결과 탭: 표준 컬럼 병합, 중복 제거, 필터, 엑셀 다운로드
- 설정 / 로그 탭: 수집 조건과 실행 로그 확인

## 엑셀 시트

- 통합결과
- 식약처 원본
- FDA 원본
- EUDAMED 원본
- 수집로그
- 검색조건

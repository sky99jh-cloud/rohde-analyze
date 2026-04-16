# rohde-analyze — 작업 내역

## 프로젝트 개요

**로데 & 슈바르츠 TX-A 로그 분석기**  
송신기 파라미터 스냅샷 HTML을 파싱해 측정값을 추출하고, 기존 Excel 양식의 **마지막 시트를 복제·갱신**한 뒤 같은 파일로 저장하는 Python 도구입니다.

요구사항 원문은 `info.md`에 정리되어 있습니다.

---

## 구현 범위

| 항목 | 내용 |
|------|------|
| 런타임 | Python 3 |
| GUI | `tkinter` — HTML / Excel 파일 선택, 실행, 진행 표시, 로그 영역 |
| HTML 파싱 | `BeautifulSoup` + `lxml` — caption 기준 테이블 섹션 매칭 |
| Excel | `openpyxl` — 시트 복사, 셀 갱신, 활성 시트·저장 |

---

## 파일 구성

| 파일 | 역할 |
|------|------|
| `main.py` | 앱 진입점. 파일 선택, 백그라운드 스레드에서 `parse_html` → `update_excel` 호출, UI 로그·메시지 박스 |
| `html_parser.py` | HTML에서 `Created on`, Power Limits, Pre-Correction, Amplifier 1/2 섹션 값 추출 |
| `excel_handler.py` | 마지막 시트 복사 → 날짜·F3/I3·AMP·특이사항 반영 → 활성 시트를 새 시트로 설정 후 저장 |
| `requirements.txt` | `beautifulsoup4`, `lxml`, `openpyxl` |
| `debug_excel.py` | Excel 마지막 시트 셀/병합 범위 확인용 보조 스크립트 |
| `info.md` | 기능·매핑 규칙 요구사항 |

---

## 데이터 매핑 (요구사항 반영)

1. **Power (F3 / I3)**  
   `Power Limits » Power and Limits`의 Forward Power → **F3**, Reflected Power → **I3** (숫자만 사용).

2. **AMP 1 / AMP 2 (B열 레이블 → C열 / D열)**  
   - Supply, Transistors, RF Levels: 해당 Amplifier 섹션에서 추출.  
   - **AMP Temp**: `Amplifier N » Status`의 **Amplifier Temp.** 사용.  
   - 시트 B열 레이블을 정규화해 파싱 키와 매칭 (`excel_handler._AMP_LABEL_MAP`).

3. **특이사항 (Shoulder Distance 등)**  
   - Non Linear: Shoulder Distance / Left / Right (값+단위).  
   - Linear: Measured Ripple, Measured Group Delay (값+단위).  
   - 레이블 행에서 오른쪽 첫 값 셀을 동적으로 찾아 값·단위 열에 기입 (`_find_value_col`).

4. **시트 처리**  
   - 기존 **마지막 시트**를 복사해 워크북 **끝**에 두고, `created_on`이 있으면 시트 이름을 `YYYY-MM-DD` 형태(중복 시 `_1` 등)로 변경.  
   - 상단 1~5행에서 인접 셀 텍스트가 `년`/`월`/`일`인 숫자 셀을 측정 일시로 갱신.  
   - 저장 시 **새로 만든 시트를 활성 시트**로 설정 → Excel에서 다시 열면 해당 시트가 보이도록 처리.

5. **UI 동작**  
   - 분석은 `threading.Thread`로 실행해 GUI 멈춤 방지.  
   - `ttk.Progressbar` indeterminate, 로그는 태그별 색상(info/success/error/detail).

---

## 실행 방법

```bash
pip install -r requirements.txt
python main.py
```

---

## 참고

- 샘플 HTML: `ParamSnapshot_*.html` (저장소에 포함된 예시).  
- Excel 템플릿 구조(열·병합)에 맞춰 B열 레이블·특이사항 행이 정의되어 있어야 자동 매핑이 동작합니다.

---

*문서 작성일: 2026-04-17*

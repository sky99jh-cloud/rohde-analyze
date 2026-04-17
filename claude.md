# rohde-analyze — 작업 내역

## 프로젝트 개요

**ROHDE 송신기 로그 분석기**  
송신기 파라미터 스냅샷 HTML을 파싱해 측정값을 추출하고, 기존 Excel 양식의 **마지막 시트를 복제·갱신**한 뒤 같은 파일로 저장하는 Python 도구입니다.

요구사항 원문은 `info.md`에 정리되어 있습니다.

---

## 구현 범위

| 항목 | 내용 |
|------|------|
| 런타임 | Python 3 |
| GUI | `tkinter` — 분석 모드(DTV/UHDTV), HTML·Excel 선택, 실행, 진행 표시, 로그 |
| HTML 파싱 | `BeautifulSoup` + `lxml` — caption 기준 테이블 섹션 매칭; AMP 개수는 모드에 따라 2 또는 6 |
| Excel | `openpyxl` — 시트 복사, 셀 갱신, 활성 시트·저장, 시트 확대 85% |

---

## 파일 구성

| 파일 | 역할 |
|------|------|
| `main.py` | 앱 진입점. 모드·파일 선택, 스레드에서 `parse_html` → `update_excel`, 모드·파일 종류 검증 |
| `html_parser.py` | `parse_html`, `detect_html_tx_kind` — 생성일, 전력, Pre-Correction(있을 때), Rack 1 Amplifier 1…N |
| `excel_handler.py` | `update_excel`, `detect_excel_tx_kind` — 마지막 시트 복제·갱신·저장 |
| `requirements.txt` | `beautifulsoup4`, `lxml`, `openpyxl` |
| `debug_excel.py` | Excel 마지막 시트 셀/병합 범위 확인용 보조 스크립트 |
| `info.md` | 기능·매핑 규칙 요구사항 |
| `sample/` | 예시 HTML·Excel (`ParamSnapshot_*.html`, DTV `D1TV_*`, UHDTV `U1TV_*` 등) |

---

## 데이터 매핑 (요구사항 반영)

1. **Power (F3 / I3)**  
   `Power Limits » Power and Limits`의 Forward Power → **F3**, Reflected Power → **I3** (숫자만 사용).

2. **AMP (B열 레이블 → C열~)**
   - **DTV**: AMP 1·2 → **C·D열**.  
   - **UHDTV**: AMP 1…6 → **C~H열**.  
   - Supply, Transistors, RF Levels, Status의 Amplifier Temp. 등은 `html_parser`와 `_AMP_LABEL_MAP` / `_EXACT_AMP_LABEL_TO_KEY`(UHDTV 짧은 레이블)로 매칭.

3. **특이사항 (Shoulder Distance 등)**  
   - **DTV만**: Non Linear / Linear 항목 기입.  
   - **UHDTV**: 해당 항목 Excel에 기입하지 않음.

4. **시트·표시**  
   - 기존 **마지막 시트**를 복사해 워크북 **끝**에 두고, `created_on`이 있으면 시트 이름을 **`YYYY-MM`**(중복 시 `_1` 등)으로 변경.  
   - **G2 / I2 / J2**: HTML Created on 기준 연·`MM월`·`DD일`.  
   - 상단 1~5행: 인접 셀 `년`/`월`/`일` 패턴으로 날짜 갱신(기존 로직).  
   - 저장 시 **새 시트 활성**, **확대/축소 85%** 설정.  
   - Excel에서 다시 열면 해당 시트가 보이도록 처리.

5. **UI·검증**  
   - 분석은 `threading.Thread`로 실행.  
   - HTML: `Rack 1 Amplifiers » Amplifier 6` 포함 여부로 DTV/UHDTV 판별(`detect_html_tx_kind`).  
   - Excel 마지막 시트 상단에 **`AMP 6`** 헤더가 있으면 UHDTV 양식(`detect_excel_tx_kind`).  
   - 찾아보기로 파일 선택 직후·실행 시 모드와 불일치하면 경고.

---

## 실행 방법

```bash
pip install -r requirements.txt
python main.py
```

---

## 참고

- 샘플 데이터는 **`sample/`** 디렉터리에 있습니다. (`ParamSnapshot_*.html`, 참고용 `*.xlsx` 등)  
- Excel 템플릿 구조(열·병합)에 맞춰 B열 레이블·특이사항 행이 정의되어 있어야 자동 매핑이 동작합니다.

---

## 날짜별 작업 내역

### 2026-04-17

- 초기 커밋: tkinter GUI, HTML 파싱(`BeautifulSoup`/`lxml`)·Excel 갱신(`openpyxl`) 파이프라인.
- 예시 `ParamSnapshot_*.html`, 참고용 `*.xlsx`를 `sample/` 디렉터리로 정리.
- `claude.md`를 `sample/` 레이아웃에 맞게 정리.
- 앱 표시명·창 제목을 **ROHDE 송신기 로그 분석기**로 통일 (`main.py`).
- 새 시트 저장 전 **확대/축소 85%**(`excel_handler`, `sheet_view.zoomScale`).
- **G2 / I2 / J2**에 HTML **Created on** 일자 반영.
- 시트 이름을 **연·월(`YYYY-MM`)** 단위로 설정.
- **UHDTV** 분석: GUI 라디오(DTV AMP 2 / UHDTV AMP 6), `parse_html(..., num_amplifiers=6)`, 엑셀 **C~H**에 AMP1…6, UHDTV는 **특이사항(Non Linear/Linear) 미기입**.
- **HTML·Excel 종류 검증**: `detect_html_tx_kind`, `detect_excel_tx_kind`; 찾아보기 선택 직후 및 실행 시 모드와 불일치 시 경고.
- `claude.md` 본문을 위 기능에 맞게 갱신.

---

*문서 작성일: 2026-04-17 · 마지막 갱신: 2026-04-17*

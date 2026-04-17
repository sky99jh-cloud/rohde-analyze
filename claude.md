# rohde-analyze — 작업 내역

## 프로젝트 개요

**ROHDE 송신기 로그 분석기**  
송신기 파라미터 스냅샷 **HTML** 또는 **DMB VM602 텍스트 로그**를 파싱해 측정값을 추출하고, 기존 Excel 양식의 **마지막 시트를 복제·갱신**한 뒤 같은 파일(들)로 저장하는 Python 도구입니다.

요구사항 원문은 `info.md`에 정리되어 있습니다.

---

## 구현 범위

| 항목 | 내용 |
|------|------|
| 런타임 | Python 3 |
| GUI | `tkinter` — 분석 모드 **DTV / UHDTV / DMB**, 파일 선택, 실행, 진행 표시, 로그, **R&S 스타일 다크 HUD** 테마 |
| 편차 하이라이트 | `excel_deviation.py` — 최근 1년 시트 평균 대비 이탈 시 셀 빨간색·알림 (**기본 20%** / **FWD·REF·AMP Temp·특이사항 50%**) |
| ROHDE HTML | `BeautifulSoup` + `lxml` — caption 기준 테이블; AMP 2 또는 6 |
| DMB 로그 | 텍스트 파싱 — `tx-1`/`tx-2` × `pa1`~`pa5`, **digital** 줄(PWR_A 등) |
| Excel | `openpyxl` — 시트 복사, 셀 갱신, 활성 시트·저장, 시트 확대 85% |

---

## 파일 구성

| 파일 | 역할 |
|------|------|
| `main.py` | 앱 진입점. DTV·UHDTV(`parse_html`→`update_excel`) / DMB(`parse_dmb_log`→`update_dmb_excel`×2) 분기 |
| `html_parser.py` | `parse_html`, `detect_html_tx_kind` |
| `excel_handler.py` | `update_excel`, `detect_excel_tx_kind` — ROHDE 마지막 시트 복제·갱신 |
| `dmb_parser.py` | `parse_dmb_log`, `detect_dmb_excel_kind` — VM602 `.txt`, TX-A/B 엑셀 구분 |
| `dmb_excel.py` | `update_dmb_excel` — DMB TX-A 또는 TX-B 워크북 한 파일씩 갱신 |
| `excel_deviation.py` | 1년 평균 대비 편차 검사·빨간 글씨·알림 문구 수집 (`apply_deviation_highlight`, `collect_*`) |
| `.cursor/skills/frontend-design/SKILL.md` | Cursor Agent용 UI 품질 가이드(선택) |
| `requirements.txt` | `beautifulsoup4`, `lxml`, `openpyxl` |
| `debug_excel.py` | Excel 마지막 시트 셀/병합 범위 확인용 보조 스크립트 |
| `info.md` | 기능·매핑 규칙 요구사항 |
| `sample/` | HTML·ROHDE·DMB 예시 (`ParamSnapshot_*.html`, `D1TV_*`, `U1TV_*`, `20260323.txt`, `DMB_TX-*.xlsx` 등) |

---

## 데이터 매핑 (요구사항 반영)

### ROHDE (DTV / UHDTV)

1. **Power (F3 / I3)**  
   `Power Limits » Power and Limits`의 Forward Power → **F3**, Reflected Power → **I3**.

2. **AMP (B열 레이블 → C열~)**  
   - **DTV**: AMP 1·2 → **C·D열**. **UHDTV**: AMP 1…6 → **C~H열**.  
   - `_AMP_LABEL_MAP` / `_EXACT_AMP_LABEL_TO_KEY`(UHDTV 짧은 레이블).

3. **특이사항** — **DTV만** Non Linear / Linear. **UHDTV**는 미기입.

4. **시트·표시**  
   - 마지막 시트 복사, 시트명 **`YYYY-MM`**, **G2/I2/J2** Created on, 활성 시트·**85%** 확대.

5. **검증** — `detect_html_tx_kind`, `detect_excel_tx_kind`(찾아보기·실행 시).

6. **편차(ROHDE)** — 이전 시트(최근 1년) 동일 셀 평균과 비교. **F3·I3·AMP Temp·특이사항**은 **50%** 이상, 그 외는 **20%** 이상일 때 빨간 글씨·완료 시 경고.

### DMB (TX-A / TX-B)

- **로그 1개**(`*.txt`)에 `tx-1 pa1`…`pa5`, `tx-2 pa1`…`pa5` 블록; 각 블록에 **I_DRV** 줄·**digital** 줄.
- **TX-A 엑셀** ← `tx-1` 데이터, **TX-B 엑셀** ← `tx-2` 데이터. **PA1~PA5** → 열 **D~H**, 행 **4~12**(전류)·**13~20**(digital: PWR_A … I_DC).
- **E2, G2, D21, D22**는 로그에 없음 → 새 시트에서 **비움**(수동 입력).
- **F1/G1/H1** 로그 날짜 반영, 시트명·85% 확대는 ROHDE와 동일 계열.
- **TX-A / TX-B** 엑셀은 `detect_dmb_excel_kind`(A1에 TX-A / TX-B)로 찾아보기 검증.

- **편차(DMB)** — PA·digital 영역 등: **AMP Temp** 행만 **50%**, 나머지 **20%** 임계값.

---

## GUI (`main.py`)

- **테마**: 로데 프로모 그래픽에 가깝게 **딥 네이비 배경**, 헤더 **이중선형 블루 그라데이션**, 기술 **그리드**, `UHF` 워터마크, 우상단 **타깃** 장식, **스펙트럼 바**(흰·시안·오렌지 계열), 패널 **시안 글로우** 테두리, 로그는 다크 터미널 톤.
- **frontend-design** 스킬(`.cursor/skills/frontend-design/SKILL.md`)을 참고해 **산업용·HUD** 느낌과 **악센트**를 맞춤.

---

## 실행 방법

```bash
pip install -r requirements.txt
python main.py
```

---

## 참고

- 샘플 데이터는 **`sample/`** 디렉터리에 있습니다.  
- ROHDE·DMB 각각 템플릿 열·레이블 구조에 맞춰야 자동 매핑이 동작합니다.

---

## 날짜별 작업 내역

### 2026-04-17

- 초기 커밋: tkinter GUI, HTML 파싱·Excel 갱신 파이프라인.
- 예시 파일 `sample/` 정리, 앱명 **ROHDE 송신기 로그 분석기**, 시트 **85%** 확대, **G2/I2/J2**·시트명 **YYYY-MM**, **UHDTV**(AMP 6)·특이사항 UHDTV 미기입, HTML/Excel 종류 검증.

### 2026-04-17 (DMB)

- **DMB 모드**: VM602 **`.txt`** 로그, **`DMB_TX-A.xlsx` / `DMB_TX-B.xlsx`** 동시 지정·한 번 실행으로 두 파일 저장 (`dmb_parser.py`, `dmb_excel.py`).
- `tx-1` → TX-A, `tx-2` → TX-B; **digital** 줄 기준으로 DIGITAL 영역(PWR_A 등) 매핑.
- 복사된 새 시트에서 **E2, G2, D21, D22**는 **비움**(수동 입력용).
- `claude.md` 본문에 DMB·파일 구성 반영.

### 2026-04-17 (편차·GUI·스킬)

- **`excel_deviation.py`**: ROHDE·DMB 엑셀에 대해 이전 시트(최근 1년) 평균 대비 편차 검사, 임계값 초과 셀 **빨간 글씨**, 완료 시 **경고/로그**. **FWD(F3)·REF(I3)·AMP Temp·특이사항**은 **50%**, 그 외 **20%**.
- **`main.py` / `excel_handler.py` / `dmb_excel.py`**: 편차 파이프라인 연동, 사용자 메시지에 항목별 임계값 안내.
- **GUI**: R&S 프로모 스타일 **다크 HUD** — 그라데이션 헤더, 그리드·워터마크·스펙트럼 바·시안 패널 테두리 등.
- **`.cursor/skills/frontend-design/SKILL.md`**: Cursor Agent용 **frontend-design** 스킬을 프로젝트에 추가(선택 적용).

---

*문서 작성일: 2026-04-17 · 마지막 갱신: 2026-04-17*

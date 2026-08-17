# namu-xlsx-converter

나무위키 표 문법 ↔ Excel(.xlsx) 양방향 변환기

---

## 파일 구성

| 파일 | 방향 | 설명 |
|---|---|---|
| `xlxstonamu.py` | xlsx → 나무위키 | Excel 표를 나무위키 문법으로 변환 |
| `namutoxlsx.py` | 나무위키 → xlsx | 나무위키 RAW 표 문법을 Excel로 변환 |

---

## xlxstonamu.py — Excel → 나무위키

```
usage: xlxstonamu.py [-h] [-r RANGE] [-s SHEET]
                     [-a {intlink,extlink,footnote,macro,table,all} [...]]
                     [--fx {input,output}]
                     file
```

> **주의:** 파일을 더블클릭하지 마세요. cmd/PowerShell/Python 터미널에서 실행하세요.  
> Windows에서는 UTF-8 강제를 위해 먼저 `chcp 65001` 입력을 권장합니다.

### 인수

| 인수 | 설명 |
|---|---|
| `file` | 변환할 Excel 파일 경로 (.xlsx) |
| `-r RANGE` | 변환할 셀 범위 (예: `A1:D5`). 생략 시 시트 전체 |
| `-s SHEET` | 대상 시트명. 생략 시 첫 번째 시트 |
| `-a ...` | 이스케이프 없이 원형 유지할 나무위키 문법 지정 |
| `--fx {input,output}` | 수식 처리 방식. `output`: 계산된 결과값(기본값), `input`: 수식 문자열 그대로 |

`-a` 옵션 값 (지정한 문법은 이스케이프하지 않고 원형 유지):

| 값 | 보존되는 토큰 |
|---|---|
| `intlink` | `[[`, `]]` |
| `extlink` | `[http`로 시작하는 외부링크, `]` |
| `footnote` | `[*` |
| `macro` | `{{{#`, `[include`, `[youtube` |
| `table` | `\|\|` |
| `all` | 위 전체 |

> **주의:** 폰트 서식 토큰(`'''`, `''`, `__`, `~~`, `--`)은 `-a` 지정과 무관하게 **항상 이스케이프**됩니다.

### 예시

```bash
xlxstonamu.py table.xlsx
xlxstonamu.py table.xlsx -r A1:D10 -s Sheet2
xlxstonamu.py table.xlsx -a intlink extlink --fx input
```

---

## namutoxlsx.py — 나무위키 → Excel

```
usage: namutoxlsx.py [-h] [--sheet SHEET] input output
```

### 인수

| 인수 | 설명 |
|---|---|
| `input` | 입력 파일 경로 (.txt 등). `-` 이면 stdin |
| `output` | 출력 파일 경로 (.xlsx) |
| `--sheet SHEET` | 시트 이름. 생략 시 `Sheet1` |

### 지원 문법

- `\|\|` 분리 파싱 — `||||||` 방식(빈 토큰 누적)과 `<-N>` 방식 모두 처리
- `<-N>` colspan, `<|N>` rowspan → 엑셀 병합 셀로 변환
- 큰 colspan 자동 clamp (실제 그리드 폭 초과 시 잘라냄)
- `[br]` → 엑셀 셀 내 개행
- 마크업 자동 제거: `{{{#색상}}}`, `[[링크|텍스트]]`, `[*각주]`, `[[파일:]]`, `'''굵게'''` 등

### 예시

```bash
python namutoxlsx.py input.txt output.xlsx
python namutoxlsx.py input.txt output.xlsx --sheet 시간표
cat raw.txt | python namutoxlsx.py - output.xlsx
```

---

## 변환 범위 및 한계

### xlxstonamu.py
- 병합 셀 → `<-N>` / `<|N>` 속성으로 변환
- 셀 배경색 → `<bgcolor=#xxxxxx>` 속성으로 변환
- 폰트 서식(굵게·기울임·밑줄·취소선) → 나무위키 마크업으로 변환
- 날짜·시간 서식 → `YYYY-MM-DD` / `HH:MM:SS` 문자열로 변환 (`--fx output` 기본값)
- 숫자·통화 등 날짜 외 특수 서식이 적용된 셀 → **오류 중단** (엑셀에서 텍스트로 확정 후 재실행 필요)
- 셀 내 줄바꿈(Alt+Enter) → `[br]` 변환

**서식별 동작 정리:**

| 엑셀 서식 | 동작 | 비고 |
|---|---|---|
| 텍스트(`@`) | 입력값 그대로 출력 | 가장 안전 |
| 일반(`General`) + 문자열 값 | 그대로 출력 | 노선번호·이름 등 |
| 일반(`General`) + 숫자/부동소수점 값 | 원시값 그대로 출력 | 의도와 다를 수 있음 |
| 일반(`General`) + 날짜·시간 값 | `YYYY-MM-DD` / `HH:MM:SS` 변환 | `--fx output` 기준 |
| 숫자·통화 등 특수 서식 | **오류 중단** | 엑셀에서 텍스트 서식으로 변환 후 재실행 |

> 나무위키 표는 보이는 텍스트가 곧 내용이므로, 엑셀의 숫자 서식처럼 동일한 값(`1000`)이 `1,000`·`₩1,000`·`1.0E+3` 등으로 달리 표시될 수 있는 경우는 변환 전 셀 서식을 **텍스트(`@`)로 확정**하는 것을 권장합니다.

### namutoxlsx.py
- 굵기나 기울임꼴 등 서식 정보는 변환하지 않음 (텍스트 내용과 병합 구조, 색상 정도만 보존)
- `{{{#!wiki}}}` 중첩 블록 내부 표는 미지원
- `<nopad>` 전용 빈 셀 행은 빈 행으로 출력됨

---

## 요구사항

-Windows 전체, Linux 가상 환경(venv 등)
```cmd
pip install openpyxl
```

-Linux 시스템(pip 시도 시 external... --break-system-packages 뜨는 경우)
```bash
sudo apt install python3-openpyxl
```

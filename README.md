# xlsx-namuwiki
Convert xlsx to namuwiki grammar

usage: xlxstonamu.py [-h] [-r RANGE] [-s SHEET]
                  [-a {intlink,extlink,footnote,macro,table,all} [{intlink,extlink,footnote,macro,table,all} ...]]
                  [--fx {input,output}]
                  file
파일을 바로 클릭하지 마세요. 자기 할 일만 하고 바로 꺼집니다. cmd나 PowerShell이나 Python을 킨 다음 명령줄에 입력하세요.

Excel(.xlsx) 표를 나무위키 표 문법으로 변환합니다.
Windows에서는 UTF-8을 강제하기 위해 우선 터미널(또는 cmd 또는 PowerShell 또는 Python)에서 chcp 65001 을 입력하는 것을 권장합니다.

positional arguments:
  file                  변환할 Excel 파일 경로 (.xlsx)

options:
  -h, --help            show this help message and exit
  
  -r RANGE, --range RANGE
                        변환할 셀 범위 (예: A1:D5). 생략 시 시트 전체
                        
  -s SHEET, --sheet SHEET
                        대상 시트명. 생략 시 첫 번째 시트
                        
  -a {intlink,extlink,footnote,macro,table,all} [{intlink,extlink,footnote,macro,table,all} ...], --apply {intlink,extlink,footnote,macro,table,all} [{intlink,extlink,footnote,macro,table,all} ...]
                        이스케이프 없이 원형 유지할 나무위키 문법 (intlink: [[링크]], extlink: [URL], footnote: [*주석], macro: {{{매크로}}} 등,
                        table: ||표||, all: 전체)
                        
  --fx {input,output}   수식 처리 방식. output: 계산된 결과값(기본값), input: 수식 문자열 그대로

예시:
  xlxstonamu.py table.xlsx
  xlxstonamu.py table.xlsx -r A1:D10 -s Sheet2
  xlxstonamu.py table.xlsx -a intlink extlink --fx input

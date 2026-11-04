from pathlib import Path

ROOT=Path(__file__).resolve().parent
PROJECT=ROOT.parents[1]
path=PROJECT/'VIDEO_SCENARIO.yml'
text=path.read_text(encoding='utf-8')
start=text.index('design:\n')
end=text.index('editing:\n',start)
text=text[:start]+'''design:
  background: '#F3F5F4'
  ink: '#1C2825'
  muted: '#52645C'
  accent: '#087E70'
  labels: plain-text
  font: Pretendard
  skill: design-taste-frontend
  direction: quiet-instrument
  directionSeed: 259388649
  applicationContent: actual-capture
  layout: 'After 38 seconds, fixed full workspace with Excel left and Celly right; local loupe only'
'''+text[end:]
text=text.replace('  chapterCardMs: 2000\n','''  chapterCardMs: 2000
  chapterCardsBeforeSeconds: 38
  fullWorkspaceAfterSeconds: 38
  fullWorkspaceSeconds: 97
  localLoupe:
    maxSimultaneous: 1
    entryMs: 320
    exitMs: 240
    sourceMarker: thin-accent-outline
    placement: blank-excel-area
''')
text=text.replace('  virtualCursor: true','  virtualCursor: false')
text=text.replace('  layout: workspace-and-closeups','  layout: fixed-workspace-with-local-loupe-after-38s')
text=text.replace('    title: 파일 첨부 없이, 말로 이어가는 Excel 작업\n','    title: 파일 첨부 없이, 말로 이어가는 Excel 작업\n    composition: full-workspace\n')
lines=text.splitlines()
for i,line in enumerate(lines):
    if line.startswith('  continuity:'):
        lines[i]='  continuity: 같은 셀리 세션과 통합문서에서 세 요청을 연속 실행하며 후속 두 요청은 같은 결과 시트에 적용한다. 38초까지 기존 내용과 시간을 유지하고, 38초부터 끝까지 두 앱 전체의 위치와 크기를 고정한다. 요청·합계·수식·저장 상태만 한 번에 하나의 국소 돋보기로 보여준다. 전송은 각 4초, 결과 반영은 각 6초의 실제 녹화다. 이후 단계 안내와 엔딩도 전체 화면을 유지한다. 전체 작업 공간은 총 97초다.'
    if line.startswith('    target: 동일 결과 시트의 금액 셀과 합계만'):
        lines[i]="    target: 동일 결과 시트의 금액과 합계만 천원 단위 표시. 연도·회사명·원래 값·수식은 보존하고 표시형식으로 검증. 전체 화면의 실제 변화 6초, 총합 돋보기 2초, F13 수식 돋보기 2초, 최종 전체 화면 5초로 구성. 변경된 금액은 천원 단위를 명시하고 과거 화면으로 되돌아가는 비교 컷은 넣지 않음"
path.write_text('\n'.join(lines)+'\n',encoding='utf-8')

from pathlib import Path
import hashlib
import json
import shutil

ROOT=Path(__file__).resolve().parent
PROJECT=ROOT.parents[1]
report_path=ROOT/'media-verification.json'
report=json.loads(report_path.read_text(encoding='utf-8'))
assert report['passed'] and report['video']['fullDecodePassed']
report['visualReview']='Passed: 22 encoded frames including 38-second boundary, lens entrance, prompts, totals, formula, save and full-workspace ending; contact plus full-resolution total inspected.'
source=ROOT/'셀리_소개_완성본.mp4'
copies=[PROJECT/'셀리_소개_완성본.mp4',PROJECT/'presentation/assets/demo.mp4']
for target in copies:
    shutil.copy2(source,target)
    assert hashlib.sha256(target.read_bytes()).hexdigest()==report['video']['sha256']
report['identicalCopies']=[str(p) for p in copies]
report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')

path=PROJECT/'VIDEO_SCENARIO.md'
lines=path.read_text(encoding='utf-8').splitlines()
replacements={
    '- 본편:':'- 본편: `셀리_소개_완성본.mp4` — **123초, 1920×1080, 30fps**, 한국어 자막·배경음·실제 전송 시점의 클릭 효과음 포함. 전체 작업 공간은 **27개 장면, 총 97초**다. **38초부터 끝까지 전체 화면을 고정**하고 9개 장면에 국소 돋보기를 사용한다. 단계 안내와 엔딩도 전체 화면을 유지한다.',
    '- 채팅용:':'- 채팅용: `셀리_10초_채팅용.gif` — **10초, 1152×648**, 무음. 전체 구간에서 두 앱을 함께 보여주고 중간 요청·합계는 돋보기로 표시한다.',
    '- 영상 검증:':'- 영상 검증: 완성 MP4 전체 디코딩, 38초 경계·돋보기 입장·입력·합계·수식·저장·엔딩을 포함한 인코딩 화면 **22개**, GIF 재생 시간과 크기를 확인했다. 전체 화면 실제 녹화 9개(전송 3개·결과 반영 3개·저장 1개·처리 2개)를 사용한다. `workspace-verification.json`과 `media-verification.json`에 검사 결과를 기록했다.',
    '- 전체 화면 보강:':'- 전체 화면 보강: 검증된 take 2 원본만 재편집했다. 38초 이후 확대 컷과 별도 간지·엔딩 카드를 제거하고, 전체 화면 위 돋보기로 실제 문장과 숫자를 보여준다. `design-taste-frontend`의 Quiet instrument 방향을 단색 배경·서체·여백·절제된 동작에 적용했다. 이번 편집에서 Excel 작업을 다시 실행하지 않았다. 루트 완성본과 발표자료용 동영상은 같은 수정본으로 교체하고 해시 일치를 확인했다.'
}
for i,line in enumerate(lines):
    for prefix,replacement in replacements.items():
        if line.startswith(prefix):
            lines[i]=replacement
path.write_text('\n'.join(lines)+'\n',encoding='utf-8')

path=PROJECT/'VIDEO_SCENARIO.yml'
lines=path.read_text(encoding='utf-8').splitlines()
for i,line in enumerate(lines):
    if line.startswith('captureStatus:'):
        lines[i]='captureStatus: 2026-11-03 전체 화면·돋보기 재편집 및 검증 완료. 38초부터 끝까지 전체 작업 공간 고정, 전체 화면 27개 장면·97초, 국소 돋보기 9개 장면, 실제 전체 화면 녹화 9개 사용. design-taste-frontend의 Quiet instrument 방향 적용. 출력은 123초 1920x1080 30fps MP4와 전체 화면을 유지하는 10초 1152x648 GIF. 전체 디코딩·인코딩 화면 22개·GIF 시간 검증 통과. 기존 계산 검증 60개 값, 총합 52902043315원, 값·수식·원본 500행 보존 및 저장 결과 유지. Excel 작업은 다시 실행하지 않음. 기존 완성본 및 발표자료 사본 교체와 해시 일치 확인. 편집 및 검증 기록은 outputs/demo-video-followup 내 JSON 참조.'
path.write_text('\n'.join(lines)+'\n',encoding='utf-8')
print('Final copies and scenario documentation updated; SHA-256 '+report['video']['sha256'])

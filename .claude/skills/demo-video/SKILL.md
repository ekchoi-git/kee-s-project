---
name: demo-video
description: YAML 시나리오 한 장으로 Playwright 웹 화면 시연 영상을 만든다. 가상 커서, 클릭 효과, 하단 자막, 챕터 카드, 영역 하이라이트를 넣고 1920x1080 MP4로 저장한다. 음성은 없음, 로컬 TTS, OpenAI 중 선택한다. 웹 시연 영상, 데모 영상, 제품 소개 영상을 자동으로 만들 때 사용한다.
---

# YAML 기반 웹 시연 영상

이 스킬은 `scripts/run-demo.mjs` 하나로 시나리오를 실행한다. 매 영상마다 Playwright 코드를 작성하지 않는다. 실행 전 이 폴더에서 `npm install` 및 `npx playwright install chromium`을 한 번 수행한다.

```powershell
node scripts/run-demo.mjs scenario.yml
```

- 출력은 시나리오 파일 기준 `out/<name>/`에 만든다. 최종본은 `final.mp4`(H.264/yuv420p/faststart), 원본 녹화본은 `recording.webm`이다.
- 실패하면 `out/<name>/failure.png`와 `failure.json`을 남기고 실패 지점의 액션을 보고한다.
- 브라우저에서 주입하는 DOM 오버레이가 커서, 클릭 파문, 하단 자막, 챕터 카드, 하이라이트를 담당한다. 실제 Playwright 포인터는 영상에 찍히지 않는다.
- `ffmpeg-static`의 실행 파일로 MP4를 인코딩한다. Playwright 번들 ffmpeg에 의존하지 않는다.

## 시나리오

`templates/scenario.yml`를 복사해 수정한다. 필요한 액션은 `goto`, `move`, `click`, `type`, `wait`, `highlight`, `screenshot`이다. `selector`는 CSS 선택자 또는 Playwright 텍스트 셀렉터(`text=로그인`)를 쓴다.

음성 모드는 `voice.mode`에 `none`, `local`, `openai` 중 하나를 지정한다. 장면의 `narration`은 화면 자막과 나레이션에 공통으로 쓴다. 음성을 쓸 때 실행기는 문장별 음성을 캐시해 먼저 생성하고 길이를 측정한 뒤, 액션 시간과 비교해 장면을 연장하고 여백을 추가한다.

- `none`: 화면과 자막만 녹화한다.
- `local`: Windows 설치 음성 엔진으로 WAV를 만든다. `voice.name`을 비우면 설치된 한국어(ko-KR) 음성을 쓰고, 없으면 실패한다.
- `openai`: `.env`의 `OPENAI_API_KEY`를 사용한다. `voice.voice`를 비우면 `voice.style` 문구로 음색을 고른다(기본 `marin`, 밝은 톤 `coral`, 중후한 톤 `onyx`). 키가 없거나 요청이 실패하면 즉시 실패하며 다른 방식으로 바꾸지 않는다.

`.env`는 Git에 넣지 말고 브라우저 로그나 자막·페이지 데이터에 키를 전달하지 않는다. OpenAI 음성 사용 사실은 배포 영상 설명 등 적절한 곳에 고지한다.

## 검수

완료 뒤 `final.mp4`를 재생하고, `timeline.json`의 장면 길이와 자막 가독성, 하이라이트 위치를 확인한다. 재시도는 의도적으로 시나리오를 고친 뒤 수행한다. 음성을 쓰면 `narration.m4a`, `subtitles.srt`, 문장별 캐시 `voice-cache/`도 같은 폴더에 남는다.

스킬 자체를 고친 뒤에는 `node scripts/smoke.mjs`로 확인한다.

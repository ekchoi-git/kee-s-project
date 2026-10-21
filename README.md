# Excel AI 도우미 (개발 중)

열려 있는 Excel에 한국어로 시키면 수식·피벗·필터·정렬을 바로 적용하고, 직전 작업을 되돌릴 수 있는 Windows 프로그램. 요구사항은 [PRD.md](PRD.md).

## 실행
1. Excel에서 작업할 파일을 연다.
2. `OPENAI_API_KEY`가 사용자 환경 변수에 있어야 한다 (모델 변경: `OPENAI_MODEL`, 기본 `gpt-5.6-terra`).
3. 이 폴더에서:
   ```
   pip install -r requirements.txt
   python -m app
   ```

## 테스트
- `python tests/test_validate.py` — Excel 없이 계획 검사 규칙 확인
- `python tests/e2e_excel.py` — 500행 파일 임시 복사본으로 실행·정답·되돌리기 확인 (`--api`를 붙이면 실제 ChatGPT까지)

## 현재 지원
수식 입력(동적 배열 포함), 피벗(새 시트), 필터, 정렬, 직전 작업 1단계 되돌리기. 의미 분류·데이터 검증은 다음 단계.

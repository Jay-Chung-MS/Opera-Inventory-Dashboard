# Opera Inventory Dashboard

## 개요
스킨이데아(메디필) 재고 대시보드.

## Memory (프로젝트 기억 시스템)
> 세션 시작 시 `.claude/memory/MEMORY.md` 자동 로드 (200줄 제한)
> 상세 토픽 파일은 필요할 때만 lazy load

- 인덱스: `.claude/memory/MEMORY.md`
- 토픽 파일: patterns.md, debugging.md, gotchas.md
- **규칙**: 세션에서 비직관적이거나 반복될 교훈을 발견하면 해당 토픽 파일에 누적
- **금지**: 코드에서 직접 읽을 수 있는 정보, git log로 알 수 있는 정보는 저장하지 않음

## 표기 규칙
- 한국어 (영문 병기), 천단위 쉼표 필수
- 시각화 폰트: 한글 — 맑은 고딕, 영문 — Arial

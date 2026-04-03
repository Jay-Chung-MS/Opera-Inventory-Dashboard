# CHANGELOG — Opera Inventory Dashboard

> 수정 시 이 파일 상단에 기록을 추가해주세요.  
> 형식: `[날짜] 작업자 — 변경 내용`

---

## 2026-04-03 Jay — SCM 대시보드 도입기 1/3 구현

### 추가
- **재고 분석 탭 (신규)**: DIO 게이지바 2종(현재고/입고반영), 부진재고 테이블
- **조달 관리 탭 (신규)**: 쇼티지 그룹(즉시대응 93개/모니터링 31개) + 입고계획 입력 폼(localStorage)
- 전체 현황 KPI 6칸 확장: 평균 DIO(정상 SKU 83일) + 쇼티지 수
- Python 파이프라인: DIO 계산, 부진재고 분석, 조달관리 데이터 추가

### 변경
- 발주 스케줄 탭 삭제 (현업 요청) → 데이터는 조달 관리에서 재활용
- DIO 평균: 과잉/부진 제외 정상 SKU만 평균 + 중앙값 표시
- DIO 게이지 MAX: 고정 180일 → 95th percentile 기반 동적 조정
- 조달관리: S95/S90만 → alert critical/warning 전체 포함 (48→93개)
- fetch → XMLHttpRequest 교체 (file:// 호환)

### 참고
- 현업 요청사항 상세: `SCM_대시보드_도입기_보고서.pdf` 참조
- 2차 도입기 예정: 수요예측 필터, 태스크 댓글, 부진재고 트렌드, 단종관리

---

## 2026-03-24 Jay — 태스크 관리 탭 추가

- 태스크 CRUD, 상태/우선순위 필터, Start/Due Date
- localStorage 저장

## 2026-03-18 Jay — 초기 대시보드 구축

- 전체 현황, SKU 관리, 수요예측, 발주 스케줄 탭
- Python 데이터 파이프라인 (S&OP + 가용재고 + B2B Top SKU → JSON)
- GitHub Pages 배포

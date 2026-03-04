---
name: design-guide
description: "Dashboard UI design guidelines: Slate Indigo palette, spacing rules, typography, component patterns. Reference before any UI changes."
argument_hint: "[component: 'button' | 'card' | 'table' | 'status' | 'all']"
---

# Design Guide — DoK Workflow Dashboard

> 출처: "바이브코딩 결과물, 3초만에 들통나는 이유" (메이커 에반) 6가지 규칙 적용
> 팔레트: Slate Indigo (Coolors 원칙 기반)

## Color Palette

### Brand Colors (2색만 사용)
```
Primary:    #818CF8  (Indigo-400)  — CTA 버튼, 강조, 활성 요소
Accent:     #6366F1  (Indigo-500)  — 그라데이션, hover 상태
```

### Neutral (나머지는 전부 이 계열)
```
Background: #0F172A  (Slate-900)   — 메인 배경
Surface:    #1E293B  (Slate-800)   — 카드, 사이드바, 메트릭
Border:     #334155  (Slate-700)   — 구분선, 테두리
Muted:      #94A3B8  (Slate-400)   — 보조 텍스트, 라벨
Text:       #F1F5F9  (Slate-100)   — 본문
Heading:    #E2E8F0  (Slate-200)   — 제목
```

### Status Colors (기능적, 브랜드 아님)
```
Success:    #4ADE80  (Green-400)   — 완료 ✅
Error:      #F87171  (Red-400)     — 실패 ❌
Warning:    #FBBF24  (Amber-400)   — 실행중 ⏳, 주의
Pending:    #64748B  (Slate-500)   — 대기 ⬜
```

### Conditional (데이터 시각화)
```
Negative:   #F87171  — 마이너스 마진, 에러 수치
Low:        #FBBF24  — 낮은 마진 (10% 미만)
High:       #4ADE80  — 높은 마진 (50% 초과)
Empty Cell: #3B1C1C  — 미입력 셀 배경 (어두운 적색)
```

## Typography

```
Headers:    font-weight 700 (h1), 600 (h2, h3)
Body:       default (400)
Labels:     0.78rem, color: #94A3B8
Monospace:  Consolas, SF Mono, Fira Code — 로그, 코드
```

## Spacing & Radius

```
Card padding:      14px 16px
Card border-radius: 10px
Button radius:     8px
Tab radius:        6px
Alert radius:      8px
Sidebar item gap:  0.35rem
```

## Button Patterns

```css
/* Primary (CTA) */
background: linear-gradient(135deg, #818CF8, #6366F1);
box-shadow: 0 2px 8px rgba(99, 102, 241, 0.25);

/* Primary hover */
background: linear-gradient(135deg, #6366F1, #4F46E5);
box-shadow: 0 4px 12px rgba(99, 102, 241, 0.35);

/* Secondary */
border: 1px solid #334155;
background: transparent;

/* Secondary hover */
border-color: #818CF8;
background: rgba(129, 140, 248, 0.06);
```

## Rules (6가지)

1. **색상을 직접 고르지 마라** — 이 파일의 팔레트에서만 선택
2. **2색만 사용** — Primary (#818CF8) + Accent (#6366F1). 나머지는 Slate 계열
3. **hex 코드 정확히** — 색상명("파란색") 대신 항상 hex 코드 사용
4. **MCP로 디자인하지 마라** — 이 스킬 문서를 참조하여 일관성 유지
5. **새 컴포넌트 추가 시** — 이 가이드의 패턴 따를 것
6. **상태 색상은 기능적** — Success/Error/Warning은 브랜드 색이 아닌 의미 색

## Config Reference

```toml
# dashboard/.streamlit/config.toml
[theme]
base = "dark"
primaryColor = "#818CF8"
backgroundColor = "#0F172A"
secondaryBackgroundColor = "#1E293B"
textColor = "#F1F5F9"
```

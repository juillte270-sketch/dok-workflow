# 일일 업무 현황 카톡 전송 워크플로우

## 개요
당일 워크플로우 실행 결과를 자동 집계하여 카카오톡(나에게 보내기)으로 전송한다.

## 실행 스크립트
`execution/send_daily_summary.py`

## 사용법
```bash
python execution/send_daily_summary.py --date "2026-02-24"
python execution/send_daily_summary.py --date "2026-02-24" --results '{"stage1":true,"stage2":true}'
```

### 파라미터
| 파라미터 | 필수 | 설명 |
|---------|------|------|
| `--date` | X | 대상일 (YYYY-MM-DD, 기본값: 오늘) |
| `--results` | X | 단계별 실행 결과 JSON |

## 입력
- `data/outputs/` 및 `data/inputs/` 디렉토리의 파일 존재 여부
- 카카오톡 토큰 (`kakao_token.json`)

## 출력
- 카카오톡 "나에게 보내기" 메시지
- 토큰 없을 시: `data/outputs/daily_summary_YYYYMMDD.txt` 파일 저장

## 자동 수집 항목
| 항목 | 파일 경로 | 확인 내용 |
|------|----------|----------|
| 배송리스트 | `data/outputs/배송리스트_YYYYMMDD.docx` | 존재 여부 |
| 발주데이터 | `data/inputs/processed_orders_YYYYMMDD.txt` | 존재 여부 |
| 가명세서 | `data/outputs/invoices_YYYYMMDD/` | 매장 수 |
| ZIP | `data/outputs/거래명세서_YYYYMMDD.zip` | 존재 여부 |
| 최종명세서 | `data/outputs/final_invoices_YYYYMMDD/` | 파일 수 |

## 메시지 형식
```
[도크 일일업무 리포트]
2/24(월)

=== 실행 결과 ===
  1. 발주처리/배송리스트 .. OK
  2. 발주시트 입력 .. OK
  ...

=== 생성 파일 ===
  배송리스트 .. OK
  발주데이터 .. OK
  가명세서 .. 12개 매장
  ZIP파일 .. OK

=== 처리 매장 ===
  - 샤브야키_분당점
  - 부엉이산장_강남점
  ...

=== 다음 단계 ===
  단가 입력 후 최종명세서 생성 요청
```

## 카카오톡 전송

### 토큰 갱신
- `kakao_token.json`의 refresh_token으로 자동 갱신
- POST `https://kauth.kakao.com/oauth/token` (grant_type=refresh_token)
- 갱신 실패 시 기존 토큰으로 시도

### 메시지 분할
- 카카오 API 제한: 800자
- 긴 메시지는 자동으로 줄 단위 분할 + "(N/M)" 표시

### 인증 만료 시
- 401 응답 → `kakao_auth.py` 재인증 필요
- 토큰 자체가 없으면 텍스트 파일로 폴백 저장

## 환경 변수
- `KAKAO_REST_API_KEY`: .env에 설정 필요

## 전체 파이프라인에서의 위치
```
[8] 일일 업무 현황 카톡 전송 (새벽, 모든 작업 완료 후) ← 이 워크플로우
```

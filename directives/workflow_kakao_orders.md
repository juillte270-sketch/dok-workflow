# 카카오톡 비즈니스 채널 발주 자동화

## 개요
매장들이 카카오톡 비즈니스 채널(오픈빌더 챗봇)을 통해 발주를 보내면,
웹훅 서버가 자동으로 수신/파싱/큐잉하고, 마감 트리거 시 전체 파이프라인을 실행한다.

## 아키텍처
```
[매장 카카오톡] → 비즈채널 오픈빌더 → POST /kakao/order
                                         ↓
                              Flask Server (port 5050)
                                         ↓
                              kakao_orders_{date}.json (큐)
                                         ↓  운영자 마감
                              orders_today.txt → 전체 파이프라인
```

## 실행 스크립트

| 스크립트 | 용도 |
|---------|------|
| `execution/kakao_order_server.py` | Flask 웹훅 서버 (상시 실행) |
| `execution/kakao_order_manager.py` | CLI 관리 도구 |

## 서버 실행

```bash
# 기본 실행
python execution/kakao_order_server.py

# 디버그 모드
python execution/kakao_order_server.py --debug

# 포트 변경 (기본 5050, kakao_auth.py 5000과 충돌 방지)
python execution/kakao_order_server.py --port 5050
```

## 엔드포인트

### 오픈빌더 스킬 (챗봇용)
| 메서드 | 경로 | 설명 |
|--------|------|------|
| POST | `/kakao/order` | 발주 접수 |
| POST | `/kakao/modify` | 발주 수정 (마지막 발주 교체) |
| POST | `/kakao/cancel` | 발주 취소 |

### 운영자용
| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/status` | 접수 현황 JSON |
| POST | `/trigger` | 마감 + 파이프라인 실행 |
| POST | `/reset` | 오늘 큐 초기화 |
| POST | `/register` | 매장 등록 `{"user_id": "...", "store_name": "..."}` |

## CLI 관리 도구

```bash
# 현황 확인
python execution/kakao_order_manager.py --status

# 특정 날짜 현황
python execution/kakao_order_manager.py --status --date 20260224

# 큐 → orders_today.txt 변환만
python execution/kakao_order_manager.py --export

# 마감 + 전체 파이프라인 실행
python execution/kakao_order_manager.py --trigger

# 큐 초기화
python execution/kakao_order_manager.py --reset

# 매장 등록
python execution/kakao_order_manager.py --register abc123def "샤브야키 도봉점"

# 등록 매장 목록
python execution/kakao_order_manager.py --list-stores
```

## 오픈빌더 설정 가이드

### 1. 비즈니스 채널 생성
- https://business.kakao.com → 채널 생성 → "도크발주"

### 2. 오픈빌더 챗봇 설정
- https://chatbot.kakao.com → 챗봇 생성
- 시나리오 → "발주접수" 블록 생성
- 스킬 연결: `POST https://{ngrok-url}/kakao/order`
- 시나리오 → "발주수정" 블록 생성
- 스킬 연결: `POST https://{ngrok-url}/kakao/modify`
- 시나리오 → "발주취소" 블록 생성
- 스킬 연결: `POST https://{ngrok-url}/kakao/cancel`

### 3. 스킬 서버 설정
- 스킬 탭 → URL 등록 → ngrok 터널 주소

### 4. ngrok 터널
```bash
ngrok http 5050
```
- 생성된 URL을 오픈빌더 스킬 URL에 등록

## 매장 사용자 등록

### 방법 1: 챗봇에서 직접 등록
매장 사용자가 처음 발주 시 안내 메시지 표시:
```
매장등록 샤브야키 도봉점
```

### 방법 2: CLI로 등록
```bash
python execution/kakao_order_manager.py --register USER_ID "매장명"
```

### 방법 3: API로 등록
```bash
curl -X POST http://localhost:5050/register \
  -H "Content-Type: application/json" \
  -d '{"user_id": "abc123", "store_name": "샤브야키 도봉점"}'
```

## 매장 발주 형식

매장에서 카톡으로 보내는 형식 (기존과 동일):
```
숙주 8박스
깻잎 1박스
케일 1박스
```

## 마감 후 파이프라인

`/trigger` 또는 `--trigger` 실행 시:
1. `kakao_orders_{date}.json` → `orders_today.txt` 변환
2. `process_orders.py` → 발주 분류 + 배송리스트
3. `update_order_sheet.py` → 발주시트 입력
4. `create_price_sheet.py` → 단가시트 생성
5. `generate_invoices.py` → 가명세서 생성

## 데이터 파일

| 파일 | 설명 |
|------|------|
| `data/inputs/kakao_orders_{YYYYMMDD}.json` | 일별 발주 큐 |
| `skills/order_processing/resources/kakao_user_store_map.json` | 사용자-매장 매핑 |

## 주의사항
- 서버 포트는 5050 (kakao_auth.py의 5000과 충돌 방지)
- 마감 후 추가 발주는 거부됨 → 운영자가 `--reset`으로 재오픈 가능
- 동일 매장 여러 번 발주 시 모두 병합됨 (수정은 `/kakao/modify` 사용)
- ngrok 무료 플랜은 URL이 매번 변경됨 → 유료 또는 고정 도메인 권장

## 학습 사항
- (이 섹션에 운영 중 발견한 이슈/해결 기록)

# 배송 어플 (DoK Delivery) 워크플로우

## 개요
도크 워크플로우의 배송 업무를 모바일로 확장. 기사 앱 + 관리자 대시보드로 구성.

## 기술 스택
| 레이어 | 기술 |
|--------|------|
| 모바일 | React Native (Expo SDK 55) |
| 백엔드 | Firebase (Firestore + Realtime DB + Storage + Auth) |
| 지도 | Phase 2: 카카오맵 SDK |
| 연동 | Python (execution/upload_deliveries.py) |
| 관리자 뷰 | Streamlit (dashboard/views/p9_delivery.py) |

## 프로젝트 구조
```
delivery-app/
├── app/                          # Expo Router 파일 기반 라우팅
│   ├── (auth)/login.tsx          # 로그인
│   ├── (driver)/                 # 기사 뷰 (탭)
│   │   ├── index.tsx             # 홈: 오늘의 배송
│   │   ├── deliveries.tsx        # 배송 목록 (필터)
│   │   ├── delivery/[id].tsx     # 상세 + 완료 처리
│   │   └── map.tsx               # 지도 (Phase 2)
│   ├── (admin)/                  # 관리자 뷰 (탭)
│   │   ├── index.tsx             # 현황 대시보드
│   │   ├── assign.tsx            # 기사 배정
│   │   └── tracking.tsx          # 실시간 추적
│   └── _layout.tsx               # 인증 분기 (role: driver/admin)
├── components/                   # 공용 컴포넌트
├── lib/                          # Firebase 서비스 레이어
├── hooks/                        # React Hooks
└── constants/theme.ts            # Slate Indigo 팔레트
```

## 데이터 흐름
```
processed_orders_YYYYMMDD.txt
    ↓ (upload_deliveries.py)
Firestore: deliveryDays/{YYYYMMDD}/deliveries/{storeId}
    ↓ (실시간 구독)
기사 앱: 배송 목록 → 상세 → 완료 처리
    ↓ (상태 업데이트)
관리자 앱/대시보드: 현황 모니터링
```

## 실행 방법

### 배송 데이터 업로드 (Python)
```bash
# JSON 미리보기 (Firebase 없이)
python execution/upload_deliveries.py --date 20260304 --json-only

# Dry-run (업로드 없이 출력만)
python execution/upload_deliveries.py --date 20260304 --dry-run

# Firebase 업로드
python execution/upload_deliveries.py --date 20260304
```

### 앱 실행 (React Native)
```bash
cd delivery-app
npm start        # Expo Go
npm run android  # Android
npm run ios      # iOS
```

### 대시보드
Streamlit 대시보드 → "12. 배송 관리" 탭

## Firebase 스키마

### Firestore
- `users/{uid}`: displayName, role(driver/admin), phone, active
- `deliveryDays/{YYYYMMDD}`: status, totalDeliveries, completedCount, createdAt
- `deliveryDays/{YYYYMMDD}/deliveries/{storeId}`: storeName, order, assignedTo, status, items[], completedAt, photoUrl, signatureUrl, notes
- `deliveryDays/{YYYYMMDD}/assignments/{uid}`: driverName, storeIds[]

### Realtime Database
- `driverLocations/{uid}`: lat, lng, heading, speed, updatedAt

### Storage
- `deliveries/{YYYYMMDD}/{storeId}/photo.jpg`: 배송 완료 사진
- `deliveries/{YYYYMMDD}/{storeId}/signature.png`: 서명

## 배송 상태
| 상태 | 설명 | 색상 |
|------|------|------|
| pending | 대기 | amber |
| in_transit | 배송중 | blue |
| delivered | 완료 | green |
| issue | 문제 | red |

## 품목 색상 규칙
- **red**: 재고/창고소분 (processed_orders "재고, 창고소분" 섹션)
- **blue**: 시장구매/시장소분 (processed_orders "시장구매, 시장소분" 섹션)
- **black**: 일반 (공급처 직배송)

## 환경 변수
- `delivery-app/.env`: Firebase 프론트엔드 키, 카카오맵 키
- 프로젝트 `.env`: `FIREBASE_SERVICE_ACCOUNT_KEY` (Python Admin SDK)

## Phase 2 계획
- 오프라인 모드 (AsyncStorage 캐시)
- 푸시 알림 (FCM)
- 배송 통계 (일별/주별 분석)
- 카카오맵 경로 최적화
- HWP → Firestore stores 컬렉션 자동 구축

## 학습사항
- (초기 구현 2026-03-04) Expo SDK 55, Firebase JS SDK 사용
- upload_deliveries.py: simple_docx_converter.py의 parse_sections() 재사용
- 배송 순서는 mappings.json page_config 기반

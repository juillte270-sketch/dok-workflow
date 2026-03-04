---
name: verify-column-refs
description: "Check all scripts use correct column indices for master file sheets (발주/단가/재고). Detects column reference errors that cause data corruption."
---

# Verify Column References

Master file column indices are critical — wrong column = data corruption.

## Files to Read
- `execution/fill_prices.py`
- `execution/create_price_sheet.py`
- `execution/fill_sikbom_targeted.py`
- `execution/fill_auction_from_api.py`
- `execution/update_order_sheet.py`
- `execution/fill_inventory_sheet.py`
- `execution/generate_invoices.py`

## Master File Column Definitions

### 발주시트
| Col | Index | Content |
|-----|-------|---------|
| A | 1 | 매장명 |
| D | 4 | 기준일 |
| E | 5 | 품목 |
| F | 6 | 공급업체 |
| G | 7 | 단위 |
| H | 8 | 입고수량 |
| I | 9 | 매입단가 |
| J | 10 | 총매입금액 (=H*I) |
| K | 11 | 판매단가 |
| L | 12 | 총판매금액 (=K*H) |
| M | 13 | 마진율 |
| N | 14 | 이익률 |

### 단가시트
| Col | Index | Content |
|-----|-------|---------|
| A | 1 | 구분 (직납/재고, 소분/재고) |
| B | 2 | 날짜 |
| C | 3 | 품목 |
| D | 4 | 품위(등급) |
| E | 5 | 단위 |
| F | 6 | 경매최고가 |
| G | 7 | 경매평균가 |
| H | 8 | 식봄 |
| I | 9 | 매입가 |
| J | 10 | 거래 업체 |

## Checks

### 1. fill_prices.py
- [ ] 발주시트: A=1(store), D=4(date), E=5(item), F=6(supplier), G=7(unit), H=8(qty), I=9(buy), K=11(sell), N=14(profit)
- [ ] 단가시트: A=1(cat), B=2(date), C=3(item), E=5(unit), I=9(price), J=10(supplier)
- [ ] Formula refs: J=H*I, L=K*H, M=margin, N=profit range

### 2. fill_sikbom_targeted.py
- [ ] 단가시트: A=1(cat), B=2(date), C=3(item), D=4(grade), E=5(unit), H=8(sikbom price)

### 3. fill_auction_from_api.py
- [ ] 단가시트: A=1(cat), C=3(item), D=4(grade), E=5(unit), F=6(max), G=7(avg)

### 4. create_price_sheet.py
- [ ] 단가시트/단가양식: B=2(date), C=3(item), J=10(supplier)

### 5. update_order_sheet.py
- [ ] 발주시트 column indices match above definition

### 6. Cross-script consistency
- [ ] All scripts agree on the same column mapping
- [ ] No script uses col 8 for something other than 식봄(단가) or 입고수량(발주)

## Report Format
```
[column-refs] fill_prices.py: ✅ All 14 refs correct
[column-refs] fill_sikbom_targeted.py: ✅ All 6 refs correct
[column-refs] Cross-consistency: ✅ No conflicts
```

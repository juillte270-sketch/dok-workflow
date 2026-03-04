"""
DoKH 발주 매핑 학습 시스템

이 스크립트는 실제 발주 데이터와 공급업체 분류 결과를 학습하여
mappings.json 파일을 자동으로 업데이트합니다.

사용법:
    python execution/learn_mappings.py --customer-orders "orders.txt" --supplier-orders "supplier_classified.txt"
    
    또는 대화형 모드:
    python execution/learn_mappings.py --interactive
"""

import json
import os
import re
from collections import defaultdict
from typing import Dict, List, Set, Tuple

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAPPINGS_FILE = os.path.join(PROJECT_ROOT, "skills", "order_processing", "resources", "mappings.json")

class MappingLearner:
    def __init__(self):
        self.mappings = self.load_mappings()
        self.learned_items = defaultdict(set)  # {supplier: set(items)}
        self.learned_stock_rules = []
        
    def load_mappings(self):
        """기존 매핑 로드"""
        with open(MAPPINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def save_mappings(self):
        """업데이트된 매핑 저장"""
        with open(MAPPINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.mappings, f, ensure_ascii=False, indent=4)
        print(f"✅ 매핑 파일 저장 완료: {MAPPINGS_FILE}")
    
    def parse_orders(self, text: str) -> Dict[str, List[Tuple[str, str, float]]]:
        """
        발주 텍스트 파싱
        Returns: {store_or_supplier: [(item, unit, qty), ...]}
        """
        orders = {}
        current_section = None
        
        item_pattern = re.compile(r'^(.+?)\s+([\d.]+)([a-zA-Z가-힣]+).*$')
        
        for line in text.strip().split('\n'):
            line = line.strip()
            if not line:
                continue
                
            if line.startswith('-'):
                current_section = line.lstrip('-').strip()
                if current_section not in orders:
                    orders[current_section] = []
                continue
            
            if current_section and not line.startswith('('):
                match = item_pattern.match(line)
                if match:
                    item = match.group(1).strip()
                    qty = float(match.group(2))
                    unit = match.group(3).strip()
                    orders[current_section].append((item, unit, qty))
        
        return orders
    
    def normalize_item(self, item: str) -> str:
        """품목명 정규화"""
        # 괄호 내용 제거
        item = re.sub(r'\([^)]*\)', '', item)
        
        # 숫자+단위 제거 (예: 양배추45 -> 양배추)
        item = re.sub(r'\d+[KkLl]?$', '', item)
        
        # 특수 매핑
        mappings = {
            '일자콩나물': '콩나물',
            '피양파': '양파',
            '간마늘': '마늘',
            '취청오이': '오이',
            '세척무': '무',
            '알베기배추': '알배기',
            '알배기배추': '알배기'
        }
        
        for old, new in mappings.items():
            if old in item:
                item = new
                break
        
        return item.strip()
    
    def learn_from_examples(self, customer_orders: Dict, supplier_orders: Dict):
        """
        고객 발주와 공급업체 분류 결과를 비교하여 학습
        """
        print("\n🎓 학습 시작...")
        
        # 1. 공급업체별 품목 학습 (mappings.json canonical list)
        supplier_list = list(self.mappings.get('supplier_names',
            ['영운농산', '다모아버섯', '건영농산', '오복상회', '명진농산',
             '풍경농산', '나물향기', '태현상회', '가야웰빙', '이모유통',
             '초원농산', '소리농산', '명화농산', '우신농산', '정복상회',
             '백제유통', '정운']))
        
        for supplier in supplier_list:
            if supplier in supplier_orders:
                for item, unit, qty in supplier_orders[supplier]:
                    normalized = self.normalize_item(item)
                    self.learned_items[supplier].add(normalized)
        
        # 2. 재고/소분 규칙 학습
        if '재고, 창고소분' in supplier_orders:
            for item, unit, qty in supplier_orders['재고, 창고소분']:
                # 매장 정보 추출 (예: "감자 10kg(강남)" -> 감자, 10kg, 강남)
                store_match = re.search(r'\(([^)]+)\)', item)
                if store_match:
                    store_hint = store_match.group(1)
                    item_clean = re.sub(r'\([^)]*\)', '', item).strip()
                    
                    # 고객 발주에서 해당 품목 찾기
                    for customer_store, items in customer_orders.items():
                        if store_hint in customer_store:
                            for cust_item, cust_unit, cust_qty in items:
                                if self.normalize_item(cust_item) == self.normalize_item(item_clean):
                                    # 재고/소분 규칙 발견
                                    rule = {
                                        'item': self.normalize_item(item_clean),
                                        'unit': cust_unit,
                                        'qty': cust_qty,
                                        'store': customer_store
                                    }
                                    self.learned_stock_rules.append(rule)
                                    print(f"  📦 재고 규칙 발견: {item_clean} {cust_qty}{cust_unit} @ {customer_store}")
        
        # 3. 시장구매 규칙 학습
        if '시장구매, 시장소분' in supplier_orders:
            for item, unit, qty in supplier_orders['시장구매, 시장소분']:
                store_match = re.search(r'\(([^)]+)\)', item)
                if store_match:
                    store_hint = store_match.group(1)
                    item_clean = re.sub(r'\([^)]*\)', '', item).strip()
                    
                    for customer_store, items in customer_orders.items():
                        if store_hint in customer_store:
                            for cust_item, cust_unit, cust_qty in items:
                                if self.normalize_item(cust_item) == self.normalize_item(item_clean):
                                    rule = {
                                        'item': self.normalize_item(item_clean),
                                        'unit': cust_unit,
                                        'qty': cust_qty,
                                        'store': customer_store,
                                        'target': 'MARKET'
                                    }
                                    self.learned_stock_rules.append(rule)
                                    print(f"  🛒 시장구매 규칙 발견: {item_clean} {cust_qty}{cust_unit} @ {customer_store}")
    
    def update_mappings(self):
        """학습한 내용을 mappings.json에 반영"""
        print("\n📝 매핑 업데이트 중...")
        
        # 1. 품목-공급업체 매핑 업데이트
        updated_count = 0
        new_count = 0
        
        for supplier, items in self.learned_items.items():
            if supplier not in self.mappings['item_supplier_map']:
                self.mappings['item_supplier_map'][supplier] = []
                print(f"  ➕ 새 공급업체 추가: {supplier}")
            
            existing = set(self.mappings['item_supplier_map'][supplier])
            
            for item in items:
                if item not in existing:
                    self.mappings['item_supplier_map'][supplier].append(item)
                    new_count += 1
                    print(f"  ✨ 새 품목 추가: {supplier} <- {item}")
                else:
                    updated_count += 1
        
        # 2. 재고/시장 규칙 업데이트
        for rule in self.learned_stock_rules:
            # 중복 체크
            exists = False
            for existing_rule in self.mappings['stock_market_rules']:
                if existing_rule.get('item_keyword') == rule['item']:
                    # 기존 규칙에 조건 추가
                    condition = {
                        'unit': rule['unit'],
                        'store': rule['store'],
                        'target': rule.get('target', 'STOCK')
                    }
                    
                    # 중복 조건 체크
                    if condition not in existing_rule['conditions']:
                        existing_rule['conditions'].append(condition)
                        print(f"  🔧 규칙 업데이트: {rule['item']} 조건 추가")
                    exists = True
                    break
            
            if not exists:
                # 새 규칙 추가
                new_rule = {
                    'item_keyword': rule['item'],
                    'conditions': [{
                        'unit': rule['unit'],
                        'store': rule['store'],
                        'target': rule.get('target', 'STOCK')
                    }]
                }
                self.mappings['stock_market_rules'].append(new_rule)
                print(f"  ➕ 새 규칙 추가: {rule['item']}")
        
        print(f"\n📊 업데이트 요약:")
        print(f"  - 기존 품목 확인: {updated_count}개")
        print(f"  - 새 품목 추가: {new_count}개")
        print(f"  - 재고/시장 규칙: {len(self.learned_stock_rules)}개")
    
    def interactive_mode(self):
        """대화형 학습 모드"""
        print("""
╔══════════════════════════════════════════════════════════╗
║          DoKH 발주 매핑 학습 시스템 (대화형 모드)        ║
╚══════════════════════════════════════════════════════════╝

고객 발주 리스트와 공급업체 분류 결과를 입력하면
자동으로 매핑 규칙을 학습합니다.

입력 방법:
1. 고객 발주 리스트를 붙여넣고 Enter
2. 빈 줄을 입력하여 종료
3. 공급업체 분류 결과를 붙여넣고 Enter
4. 빈 줄을 입력하여 종료
""")
        
        print("📋 고객 발주 리스트를 입력하세요 (빈 줄로 종료):")
        customer_lines = []
        while True:
            line = input()
            if not line.strip():
                break
            customer_lines.append(line)
        
        print("\n📦 공급업체 분류 결과를 입력하세요 (빈 줄로 종료):")
        supplier_lines = []
        while True:
            line = input()
            if not line.strip():
                break
            supplier_lines.append(line)
        
        customer_text = '\n'.join(customer_lines)
        supplier_text = '\n'.join(supplier_lines)
        
        customer_orders = self.parse_orders(customer_text)
        supplier_orders = self.parse_orders(supplier_text)
        
        print(f"\n✅ 파싱 완료:")
        print(f"  - 고객 매장: {len(customer_orders)}개")
        print(f"  - 공급업체: {len(supplier_orders)}개")
        
        self.learn_from_examples(customer_orders, supplier_orders)
        self.update_mappings()
        
        # 저장 확인
        response = input("\n💾 변경사항을 저장하시겠습니까? (y/n): ")
        if response.lower() == 'y':
            self.save_mappings()
        else:
            print("❌ 저장 취소됨")

def main():
    import argparse
    parser = argparse.ArgumentParser(description="DoKH 발주 매핑 학습 시스템")
    parser.add_argument("--customer-orders", help="고객 발주 파일 경로")
    parser.add_argument("--supplier-orders", help="공급업체 분류 결과 파일 경로")
    parser.add_argument("--interactive", action="store_true", help="대화형 모드")
    args = parser.parse_args()
    
    learner = MappingLearner()
    
    if args.interactive:
        learner.interactive_mode()
    elif args.customer_orders and args.supplier_orders:
        with open(args.customer_orders, 'r', encoding='utf-8') as f:
            customer_text = f.read()
        with open(args.supplier_orders, 'r', encoding='utf-8') as f:
            supplier_text = f.read()
        
        customer_orders = learner.parse_orders(customer_text)
        supplier_orders = learner.parse_orders(supplier_text)
        
        learner.learn_from_examples(customer_orders, supplier_orders)
        learner.update_mappings()
        learner.save_mappings()
    else:
        print("❌ 오류: --interactive 또는 --customer-orders와 --supplier-orders를 함께 지정해야 합니다.")
        parser.print_help()

if __name__ == "__main__":
    main()

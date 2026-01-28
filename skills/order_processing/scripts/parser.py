import re
from typing import List, Dict, Any

class OrderParser:
    def __init__(self):
        # 기본 패턴: "-거래처명" 또는 "거래처명" 으로 섹션 시작
        self.store_pattern = re.compile(r'^-?\s*([가-힣a-zA-Z0-9\s]+(?:점|1차|무렵|햇살|프루트|선데이|유|브럭시))$')
        # 품목 패턴: "품목명 수량단위" (예: 배추 1망, 숙주 8박스)
        self.item_pattern = re.compile(r'^([가-힣\(\)a-zA-Z0-9,\s]+?)\s+(\d+(?:\.\d+)?)(?:/?([a-zA-Z가-힣]+))?$')

    def parse_text(self, text: str) -> Dict[str, List[Dict[str, Any]]]:
        """
        발주 텍스트를 파싱하여 거래처별 주문 목록을 반환합니다.
        
        Returns:
            {
                "거래처A": [{"item": "품목1", "qty": 1, "unit": "박스"}, ...],
                "거래처B": ...
            }
        """
        orders = {}
        current_store = None
        
        lines = text.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 1. 거래처 확인
            # '-'로 시작하거나 특정 키워드로 끝나는 경우 거래처로 인식 시도
            if line.startswith('-'):
                store_name = line.lstrip('-').strip()
                current_store = store_name
                if current_store not in orders:
                    orders[current_store] = []
                continue
                
            # 2. 품목 확인
            if current_store:
                # 라인 분석
                # 괄호로 시작하거나(메모), 특정 키워드가 있는 경우
                if line.startswith('(') or line.startswith('※'):
                    orders[current_store].append({
                        "type": "memo",
                        "text": line
                    })
                    continue

                # 숫자+단위 패턴 확인 (수량 파악)
                matches = list(re.finditer(r'(\d+(?:\.\d+)?)', line))
                
                if matches:
                    last_match = matches[-1]
                    qty_str = last_match.group(1)
                    qty_start = last_match.start()
                    qty_end = last_match.end()
                    
                    item_part = line[:qty_start].strip()
                    unit_part = line[qty_end:].strip()
                    
                    try:
                        qty = float(qty_str)
                        if qty.is_integer(): qty = int(qty)
                    except ValueError:
                        qty = 0

                    orders[current_store].append({
                        "type": "item",
                        "item": item_part,
                        "qty": qty,
                        "unit": unit_part,
                        "original_line": line
                    })
                else:
                    # 숫자가 없으면 메모나 특수 라인으로 간주
                     orders[current_store].append({
                        "type": "memo",
                        "text": line
                    })
        
        return orders

# 테스트 코드
if __name__ == "__main__":
    sample_text = """
-샤브야키 동탄점
배추 1망
숙주 8박스
깐양파2호 3kg
    """
    parser = OrderParser()
    result = parser.parse_text(sample_text)
    print(result)

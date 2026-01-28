import json
import os
import datetime
from typing import Dict, List, Any
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class DeliveryListGenerator:
    def __init__(self, resource_path: str):
        self.mappings = self._load_mappings(resource_path)
        self.workbook = Workbook()
        
    def _load_mappings(self, path: str) -> Dict:
        mapping_file = os.path.join(path, "mappings.json")
        try:
            with open(mapping_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Warning: Mapping file not found at {mapping_file}")
            return {}

    def _get_supplier_for_item(self, item_name: str) -> str:
        """품목명으로 공급처를 찾습니다 (퍼지 매칭 포함 가능)."""
        supplier_map = self.mappings.get("item_supplier_map", {})
        
        # 1. 완벽 일치 검색
        for supplier, items in supplier_map.items():
            if item_name in items:
                return supplier
                
        # 2. 부분 일치 검색
        for supplier, items in supplier_map.items():
            for map_item in items:
                # "깐양파"가 "깐양파2호"에 포함되는지 등 확인
                if map_item in item_name or item_name in map_item:
                    return supplier
                    
        return "기타/미분류"

    def _get_color_for_item(self, item_name: str) -> str:
        """품목명에 따른 색상 코드(00RRGGBB)를 반환합니다."""
        color_rules = self.mappings.get("color_rules", {})
        
        # Red (재고소분) - FF0000
        for keyword in color_rules.get("red", []):
            if keyword in item_name:
                return "FFFF0000" # Red
                
        # Blue (시장소분) - 0000FF
        for keyword in color_rules.get("blue", []):
            if keyword in item_name:
                return "FF0000FF" # Blue
                
        return None # 기본 검정

    def generate_excel(self, orders: Dict[str, List[Dict]], output_path: str):
        """
        주문 데이터를 기반으로 배송 리스트 엑셀 파일을 생성합니다.
        
        orders: { "거래처명": [{item, qty, unit}, ...], ... }
        """
        # 1. 데이터 재구성: 공급처별 분류
        supplier_orders = {} # { "공급처": { "거래처": [품목들...] } }
        
        for store, items in orders.items():
            for item_data in items:
                item_name = item_data["item"]
                supplier = self._get_supplier_for_item(item_name)
                
                if supplier not in supplier_orders:
                    supplier_orders[supplier] = {}
                if store not in supplier_orders[supplier]:
                    supplier_orders[supplier][store] = []
                    
                supplier_orders[supplier][store].append(item_data)

        # 2. 워크시트 생성 및 스타일링
        ws = self.workbook.active
        ws.title = "배송리스트"
        
        # 현재 요일 확인 (0:월, ... 4:금)
        weekday = datetime.datetime.now().weekday()
        config_key = "friday" if weekday == 4 else "default"
        page_config = self.mappings.get("page_config", {}).get(config_key, {})
        
        current_row = 1
        
        # 3. 페이지별 출력
        # 페이지 구성 순회 (page1, page2 ...)
        sorted_pages = sorted(page_config.keys())
        
        for page_name in sorted_pages:
            targets = page_config[page_name]
            
            # 헤더 출력
            ws.cell(row=current_row, column=1, value=f"=== {page_name} ===")
            ws.cell(row=current_row, column=1).font = Font(bold=True, size=12)
            current_row += 1
            
            if isinstance(targets, list):
                # 특정 거래처 리스트 출력
                # 해당 페이지에 배정된 거래처들의 주문을 출력
                for target_store in targets:
                    # 데이터에 해당 거래처가 있는지 확인
                    # orders 구조: { "거래처": [ {item, qty, unit}, ... ] }
                    if target_store in orders:
                        # 거래처 헤더
                        ws.cell(row=current_row, column=1, value=f"■ {target_store}")
                        ws.cell(row=current_row, column=1).font = Font(bold=True, size=11)
                        current_row += 1
                        
                        # 품목 리스트
                        for item in orders[target_store]:
                            # 텍스트 구성: 품목명 [수량] [단위]
                            parts = [item['item']]
                            if item['qty'] > 0:
                                parts.append(str(item['qty']))
                            if item['unit']:
                                parts.append(item['unit'])
                                
                            item_text = " ".join(parts)
                            cell = ws.cell(row=current_row, column=1, value=f"   - {item_text}")
                            
                            # 색상 적용
                            color = self._get_color_for_item(item['item'])
                            if color:
                                cell.font = Font(color=color)
                            
                            current_row += 1
                        current_row += 1  # 거래처 간 간격 (가독성)
                    
            elif targets == "SUPPLIER_LIST":
                # 공급처별 발주 내역 출력 (발주서 용도)
                for supplier in sorted(supplier_orders.keys()):
                    ws.cell(row=current_row, column=1, value=f"■ {supplier}")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, color="008000") # Green
                    current_row += 1
                    
                    stores_data = supplier_orders[supplier]
                    for store, items in stores_data.items():
                        # 거래처명
                        ws.cell(row=current_row, column=1, value=f"- {store}")
                        ws.cell(row=current_row, column=1).font = Font(bold=True)
                        current_row += 1
                        
                        # 품목들
                        for item in items:
                            parts = []
                            parts.append(item['item'])
                            if item['qty'] > 0:
                                parts.append(f"{item['qty']}")
                            if item['unit']:
                                parts.append(item['unit'])
                                
                            item_str = " ".join(parts)
                            cell = ws.cell(row=current_row, column=2, value=item_str)
                            
                            # 색상 적용
                            color = self._get_color_for_item(item['item'])
                            if color:
                                cell.font = Font(color=color)
                            
                            current_row += 1
                    current_row += 1 # 공급처 간 빈 줄
            
            current_row += 2 # 페이지 구분 여백

        # 파일 저장
        self.workbook.save(output_path)
        print(f"Excel file generated at: {output_path}")

# 테스트
if __name__ == "__main__":
    pass

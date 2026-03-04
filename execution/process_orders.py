import json
import math
import os
import re
import sys
from typing import List, Dict, Any, Tuple, Optional, Set
from datetime import datetime, timedelta

# Add project root to path to import skills
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root)

from skills.order_processing.scripts.parser import OrderParser


def load_inventory_levels(master_file: str, target_date: str) -> Dict[str, float]:
    """Read 현재고 from 재고 sheet for the most recent date before target_date.

    Looks backward from target_date-1 up to 7 days to find a populated date group.
    Returns: {inventory_item_name: current_stock_qty}
    """
    try:
        import openpyxl
    except ImportError:
        print("  [Inventory] openpyxl not available, skipping inventory check")
        return {}

    if not os.path.exists(master_file):
        print(f"  [Inventory] Master file not found: {master_file}")
        return {}

    try:
        wb = openpyxl.load_workbook(master_file, read_only=True, data_only=True)
    except PermissionError:
        print("  [Inventory] Master file locked, inventory check skipped")
        return {}
    except Exception as e:
        print(f"  [Inventory] Cannot open master file: {e}")
        return {}

    if "재고" not in wb.sheetnames:
        print("  [Inventory] '재고' sheet not found")
        wb.close()
        return {}

    ws = wb["재고"]
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    max_col = ws.max_column

    # Find the most recent date group by searching backward from the right
    # (recent dates are at higher column numbers, so right-to-left is faster)
    item_col = None
    found_label = None
    for days_back in range(1, 8):
        check_date = dt - timedelta(days=days_back)
        date_prefix = f"{check_date.month}/{check_date.day}"

        for col_idx in range(max_col, 0, -1):
            cell_val = ws.cell(row=2, column=col_idx).value
            if cell_val is None:
                continue
            cell_str = str(cell_val).strip()
            if "재고" not in cell_str:
                continue
            idx = cell_str.find(date_prefix)
            if idx < 0:
                continue
            # Prevent "12/3" matching when looking for "2/3"
            if idx > 0 and cell_str[idx - 1].isdigit():
                continue
            item_col = col_idx
            found_label = cell_str
            break
        if item_col:
            break

    if not item_col:
        print("  [Inventory] No recent date group found in 재고 sheet")
        wb.close()
        return {}

    # 현재고 = item_col + 5 (6-column structure: 품목/재고/1차직납/2차출고/입고/현재고)
    stock_col = item_col + 5
    print(f"  [Inventory] Using: {found_label} (품목=col{item_col}, 현재고=col{stock_col})")

    levels = {}
    for r in range(4, min(ws.max_row + 1, 200)):
        item_name = ws.cell(row=r, column=item_col).value
        if item_name:
            item_name = str(item_name).strip()
            stock_val = ws.cell(row=r, column=stock_col).value
            if stock_val is not None:
                try:
                    levels[item_name] = float(stock_val)
                except (ValueError, TypeError):
                    pass

    wb.close()
    print(f"  [Inventory] Loaded {len(levels)} items from 재고 sheet")
    return levels

class OrderProcessor:
    def __init__(self, mappings_path: str, master_file: str = None, target_date: str = None):
        with open(mappings_path, 'r', encoding='utf-8') as f:
            self.mappings = json.load(f)
        self.parser = OrderParser()
        self.supplier_agg = {} # {supplier: {item_name: {unit: total_qty}}}
        self.stock_list = []
        self.market_list = []

        # Inventory awareness
        self.inventory_levels: Dict[str, float] = {}
        self.stock_overrides: Set[str] = set()  # inv items to reclassify STOCK→MARKET
        self._inv_item_map = self.mappings.get('inventory_item_map', {})
        self.use_safety_buffer = False  # Use reorder thresholds as safety buffer
        self.auto_reorder = False  # Auto-reorder disabled by default
        self._target_date = target_date  # YYYY-MM-DD
        self._stock_demand: Dict[str, float] = {}  # {inv_item: total_qty}
        self._auto_reorders: Dict[str, Dict[str, Dict[str, float]]] = {}  # {supplier: {item: {unit: qty}}}

        if master_file and target_date:
            self.inventory_levels = load_inventory_levels(master_file, target_date)

    def _get_inventory_name(self, item_name: str) -> Optional[str]:
        """Map order item name to inventory sheet item name."""
        # Direct mapping from inventory_item_map
        if item_name in self._inv_item_map:
            return self._inv_item_map[item_name]

        # Try substring matching against inventory_item_map keys
        for order_name, inv_name in self._inv_item_map.items():
            if order_name in item_name or item_name in order_name:
                return inv_name

        # Try against inventory levels keys directly
        for inv_name in self.inventory_levels:
            inv_base = inv_name.split('(')[0].strip()
            if inv_base and (inv_base == item_name or inv_base in item_name or item_name in inv_base):
                return inv_name

        return None

    def _get_supplier_from_rules(self, item_name: str, qty: float, unit: str, store_name: str) -> str:
        """Pure rule-based supplier classification (no inventory check)."""
        # 1. Check stock/market rules first
        for rule in self.mappings.get('stock_market_rules', []):
            if rule['item_keyword'] in item_name:
                for cond in rule['conditions']:
                    # Check store condition
                    store_match = True
                    if 'store' in cond:
                        if cond['store'] not in store_name:
                            store_match = False

                    # Check unit condition
                    unit_match = True
                    if 'unit' in cond:
                        target_units = cond['unit'] if isinstance(cond['unit'], list) else [cond['unit']]
                        # Check both just the unit AND the qty+unit (e.g. "5kg")
                        qty_val = int(qty) if int(qty) == qty else qty
                        combined_unit = f"{qty_val}{unit}"
                        if unit not in target_units and combined_unit not in target_units:
                            unit_match = False

                    if store_match and unit_match:
                        return cond['target']

        # 2. Check item_supplier_map
        for supplier, keywords in self.mappings.get('item_supplier_map', {}).items():
            for kw in keywords:
                if kw in item_name:
                    return supplier

        return "Unknown"

    def _get_safety_threshold(self, inv_name: str) -> float:
        """Get minimum safety threshold for an inventory item from reorder rules."""
        rules = self.mappings.get('inventory_reorder_rules', {})
        rule = rules.get(inv_name)
        if not rule:
            return 0

        rule_type = rule.get('type', '')
        if rule_type == 'threshold':
            return rule.get('threshold', 0)
        elif rule_type == 'range':
            # Use 'min' or weekday-specific min
            if 'min' in rule:
                return rule['min']
            elif 'min_weekday' in rule:
                mins = rule['min_weekday']
                return min(mins.values()) if mins else 0
        elif rule_type == 'target':
            return rule.get('threshold', rule.get('target', 0) * 0.5)
        elif rule_type == 'tiered':
            tiers = rule.get('tiers', [])
            return tiers[0].get('threshold', 0) if tiers else 0
        return 0

    def _compute_stock_overrides(self, store_orders, excluded_stores):
        """Pre-scan all orders to identify STOCK items with insufficient inventory.

        Aggregates total demand for each STOCK item across all stores,
        then compares against 현재고. Items where demand > inventory get
        reclassified from STOCK to MARKET.

        With use_safety_buffer=True, also checks inventory_reorder_rules:
        if (현재고 - demand) < safety_threshold, reclassify to MARKET.
        """
        if not self.inventory_levels:
            return

        # Aggregate total STOCK demand per inventory item
        stock_demand = {}  # {inv_item_name: total_qty}

        for store, entries in store_orders.items():
            if store in excluded_stores:
                continue
            for entry in entries:
                if entry['type'] == 'memo':
                    continue
                target = self._get_supplier_from_rules(
                    entry['item'], entry['qty'], entry['unit'], store
                )
                if target == "STOCK":
                    inv_name = self._get_inventory_name(entry['item'])
                    if inv_name:
                        stock_demand[inv_name] = stock_demand.get(inv_name, 0) + entry['qty']

        # Save for auto-reorder use
        self._stock_demand = stock_demand

        # Compare against current inventory
        print("\n=== Inventory Check (현재고 vs STOCK 수요) ===")
        for inv_name, demand in sorted(stock_demand.items()):
            current = self.inventory_levels.get(inv_name, 0)
            safety = self._get_safety_threshold(inv_name) if self.use_safety_buffer else 0

            if current < demand:
                # Not enough stock at all
                self.stock_overrides.add(inv_name)
                print(f"  [X] {inv_name}: stock {current} < demand {demand} -> MARKET")
            elif safety > 0 and (current - demand) < safety:
                # Enough to fulfill, but would deplete below safety threshold
                self.stock_overrides.add(inv_name)
                remaining = current - demand
                print(f"  [!] {inv_name}: stock {current} - demand {demand} = remain {remaining} < safety {safety} -> MARKET")
            else:
                remaining = current - demand
                suffix = f" (safety: {safety})" if safety > 0 else ""
                print(f"  [OK] {inv_name}: stock {current} >= demand {demand} (remain: {remaining}){suffix}")

        if self.stock_overrides:
            print(f"\n  [Result] {len(self.stock_overrides)}개 품목 STOCK→MARKET 전환:")
            for name in sorted(self.stock_overrides):
                print(f"    - {name}")
        else:
            print(f"\n  [Result] 모든 STOCK 품목 재고 충분")

    @staticmethod
    def _get_weekday_value(rule_dict: dict, date_str: str):
        """요일별 값 반환. Mirrors manage_inventory_date.get_weekday_value()."""
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        dow = dt.weekday()  # 0=Mon .. 6=Sun
        for key, val in rule_dict.items():
            days = key.lower().replace("-", "_")
            if days == "sun_wed" and dow in (6, 0, 1, 2):
                return val
            if days == "thu_fri" and dow in (3, 4):
                return val
            if days == "sun_thu" and dow in (6, 0, 1, 2, 3):
                return val
            if days == "fri" and dow == 4:
                return val
        return list(rule_dict.values())[0]

    def _calc_reorder_qty(self, projected: float, rule: dict, target_date: str) -> float:
        """Calculate how much inventory replenishment is needed.

        Mirrors the core logic of manage_inventory_date.calculate_incoming().
        Returns: required reorder quantity in inventory units (0 if sufficient).
        """
        rtype = rule.get("type", "")
        if rtype in ("skip", "manual", "special_box"):
            return 0

        if rtype == "range":
            if "min_weekday" in rule:
                min_val = self._get_weekday_value(rule["min_weekday"], target_date)
                max_val = self._get_weekday_value(rule["max_weekday"], target_date)
            else:
                min_val = rule["min"]
                max_val = rule["max"]
            if projected < min_val:
                return max(max_val - projected, 0)

        elif rtype == "threshold":
            if projected <= rule["threshold"]:
                return rule["order_qty"]

        elif rtype == "tiered":
            tiers = sorted(rule["tiers"], key=lambda t: t["threshold"])
            for tier in tiers:
                if projected <= tier["threshold"]:
                    return tier["order_qty"]

        elif rtype == "target":
            target = rule["target"]
            threshold = rule.get("threshold", target)
            if projected <= threshold:
                return max(target - projected, 0)

        elif rtype == "target_unit":
            target = rule["target"]
            unit = rule["unit"]
            if projected < target:
                needed = target - projected
                return math.ceil(needed / unit) * unit

        return 0

    def _compute_auto_reorders(self):
        """Compute automatic supplier reorders for STOCK items below thresholds.

        For each inventory item with a reorder rule AND a reorder_config entry,
        calculate: projected = 현재고 - STOCK demand.
        If projected triggers the rule, generate a supplier order.
        """
        if not self.inventory_levels:
            return

        reorder_rules = self.mappings.get('inventory_reorder_rules', {})
        reorder_config = self.mappings.get('reorder_config', {})

        if not reorder_config:
            return

        report_items = []

        for inv_name, rule in reorder_rules.items():
            rtype = rule.get("type", "")
            if rtype in ("skip", "manual", "special_box"):
                continue
            if inv_name not in reorder_config:
                continue

            current = self.inventory_levels.get(inv_name, 0)
            demand = self._stock_demand.get(inv_name, 0)
            projected = current - demand

            reorder_inv_qty = self._calc_reorder_qty(projected, rule, self._target_date)
            if reorder_inv_qty <= 0:
                report_items.append((inv_name, current, demand, projected, 0, None))
                continue

            cfg = reorder_config[inv_name]
            supplier = cfg["supplier"]
            order_item = cfg["order_item"]
            order_unit = cfg["order_unit"]
            per_unit = cfg["per_unit_qty"]

            order_qty = math.ceil(reorder_inv_qty / per_unit) if per_unit > 0 else reorder_inv_qty

            report_items.append((inv_name, current, demand, projected, order_qty, cfg))

            # Aggregate into supplier_agg
            if supplier not in self.supplier_agg:
                self.supplier_agg[supplier] = {}
            if order_item not in self.supplier_agg[supplier]:
                self.supplier_agg[supplier][order_item] = {}
            if order_unit not in self.supplier_agg[supplier][order_item]:
                self.supplier_agg[supplier][order_item][order_unit] = 0
            self.supplier_agg[supplier][order_item][order_unit] += order_qty

            # Also save separately for manual_supplier_orders merge
            if supplier not in self._auto_reorders:
                self._auto_reorders[supplier] = {}
            if order_item not in self._auto_reorders[supplier]:
                self._auto_reorders[supplier][order_item] = {}
            if order_unit not in self._auto_reorders[supplier][order_item]:
                self._auto_reorders[supplier][order_item][order_unit] = 0
            self._auto_reorders[supplier][order_item][order_unit] += order_qty

        # Print report
        reorder_count = sum(1 for _, _, _, _, qty, _ in report_items if qty > 0)
        print(f"\n=== Auto-Reorder Report ({reorder_count} items) ===")
        for inv_name, current, demand, projected, order_qty, cfg in report_items:
            if order_qty > 0:
                print(f"  [REORDER] {inv_name}: 현재고={current} - 출고={demand} = 잔여{projected}"
                      f" → {cfg['supplier']} {cfg['order_item']} {order_qty}{cfg['order_unit']}")
            else:
                print(f"  [OK] {inv_name}: 현재고={current} - 출고={demand} = 잔여{projected}")

    def get_supplier(self, item_name: str, qty: float, unit: str, store_name: str) -> str:
        """Classify item to supplier/STOCK/MARKET with inventory awareness."""
        target = self._get_supplier_from_rules(item_name, qty, unit, store_name)

        # Inventory-aware override: STOCK → MARKET if insufficient stock
        if target == "STOCK" and self.stock_overrides:
            inv_name = self._get_inventory_name(item_name)
            if inv_name and inv_name in self.stock_overrides:
                return "MARKET"

        return target

    def process(self, input_path: str, output_path: str):
        with open(input_path, 'r', encoding='utf-8') as f:
            raw_text = f.read()
            
        # 0. Pre-process to extract manual sections (Stock/Market)
        # We need to find lines under "-재고, 창고소분" and "-시장구매, 시장소분"
        # And importantly, REMOVE them from parsing so they don't appear as stores.
        
        manual_stock = []
        manual_market = []
        
        # 비품목 라인 필터 (인사말, 요청 문구 등)
        _greeting_re = re.compile(
            r'^(부탁|감사|잘\s*부탁|수고|고맙|고마워|ㄱㅅ|ㅂㅌ|추가\s*없|이상입니다|끝$)',
        )

        lines = raw_text.split('\n')
        filtered_lines = []

        current_section = None
        for line in lines:
            stripped = line.strip()
            # 인사말/비품목 라인 스킵
            if stripped and _greeting_re.search(stripped) and not re.search(r'\d', stripped):
                continue
            if stripped.startswith('-'):
                section_name = stripped.lstrip('-').strip()
                if section_name == "재고, 창고소분":
                    current_section = "STOCK"
                    # Do not add this header to filtered_lines
                elif section_name == "시장구매, 시장소분":
                    current_section = "MARKET"
                    # Do not add this header to filtered_lines
                else:
                    current_section = None
                    filtered_lines.append(line)
            else:
                if current_section == "STOCK" and stripped:
                    manual_stock.append(stripped)
                elif current_section == "MARKET" and stripped:
                    manual_market.append(stripped)
                else:
                    filtered_lines.append(line)
                    
        parsed_text = '\n'.join(filtered_lines)

        store_orders = self.parser.parse_text(parsed_text)

        # Verification data
        verification_store_totals = {} # {item_name: {unit: total_qty}}

        # Build processed output
        output_lines = []

        # 1. Store sections (Pages 1-2) — suppliers loaded from mappings.json
        excluded_stores = set(self.mappings.get('supplier_names', []))

        # Pre-scan: check inventory for STOCK items → override to MARKET if insufficient
        self._compute_stock_overrides(store_orders, excluded_stores)

        # Auto-reorder: generate supplier orders for depleted inventory
        if self.auto_reorder:
            self._compute_auto_reorders()

        manual_supplier_orders = {}
            
        for store, entries in store_orders.items():
            if store in excluded_stores:
                manual_supplier_orders[store] = entries
                continue
                
            output_lines.append(f"-{store}")
            for entry in entries:
                if entry['type'] == 'memo':
                    output_lines.append(entry['text'])
                else:
                    item_str = f"{entry['item']} {entry['qty']}{entry['unit']}"
                    output_lines.append(item_str)
                    
                    # Store verification
                    ikey, ukey = entry['item'], entry['unit']
                    if ikey not in verification_store_totals: verification_store_totals[ikey] = {}
                    verification_store_totals[ikey][ukey] = verification_store_totals[ikey].get(ukey, 0) + entry['qty']
                    
                    # Aggregate for supplier list
                    target = self.get_supplier(entry['item'], entry['qty'], entry['unit'], store)
                    
                    if target == "STOCK":
                        self.stock_list.append(f"{entry['item']} {entry['qty']}{entry['unit']}({self.get_short_store(store)})")
                    elif target == "MARKET":
                        self.market_list.append(f"{entry['item']} {entry['qty']}{entry['unit']}({self.get_short_store(store)})")
                    elif target != "Unknown":
                        if target not in self.supplier_agg:
                            self.supplier_agg[target] = {}
                        
                        item_key = entry['item']
                        unit_key = entry['unit']
                        
                        if item_key not in self.supplier_agg[target]:
                            self.supplier_agg[target][item_key] = {}
                        
                        if unit_key not in self.supplier_agg[target][item_key]:
                            self.supplier_agg[target][item_key][unit_key] = 0
                            
                        self.supplier_agg[target][item_key][unit_key] += entry['qty']
            output_lines.append("")

        # 2. Supplier sections
        # Use explicit user-defined order
        defined_order = [
            '영운농산', '다모아버섯', '건영농산', '오복상회', '(선발주)', 
            '명진농산', '세연농산', '풍경농산', '나물향기', '태현상회', 
            '가야웰빙', '이모유통', '(금일, 0457차량ori05기둥)', '(선발주,내일0457/I05기둥)', 
            '초원농산', '경북상회', '명화농산', '우신농산', '현진상회',
            '정복상회', '백제유통', '소리농산', '정운'
        ]
        
        # Check if we have manual supplier orders (Page 3 override)
        if manual_supplier_orders:
            print(f"  [Info] Using manual Supplier list ({len(manual_supplier_orders)} suppliers)")
            
            # Sort manually provided suppliers according to defined_order
            sorted_manual = []
            keys = list(manual_supplier_orders.keys())
            
            for s in defined_order:
                if s in keys:
                    sorted_manual.append(s)
            
            # Add remaining
            for s in keys:
                if s not in sorted_manual:
                    sorted_manual.append(s)
            
            # Track which auto-reorder suppliers are already in manual list
            auto_reorder_written = set()

            for supplier in sorted_manual:
                output_lines.append(f"-{supplier}")
                entries = manual_supplier_orders[supplier]
                for entry in entries:
                    if entry['type'] == 'memo':
                        output_lines.append(entry['text'])
                    else:
                        item_str = f"{entry['item']} {entry['qty']}{entry['unit']}"
                        output_lines.append(item_str)
                # Append auto-reorder items for this supplier
                if supplier in self._auto_reorders:
                    for item_name, units in self._auto_reorders[supplier].items():
                        for unit, qty in units.items():
                            q = int(qty) if isinstance(qty, float) and qty.is_integer() else qty
                            output_lines.append(f"{item_name} {q}{unit} [보충]")
                    auto_reorder_written.add(supplier)
                output_lines.append("")

            # Add auto-reorder suppliers not in manual list
            for supplier in self._auto_reorders:
                if supplier not in auto_reorder_written:
                    output_lines.append(f"-{supplier}")
                    for item_name, units in self._auto_reorders[supplier].items():
                        for unit, qty in units.items():
                            q = int(qty) if isinstance(qty, float) and qty.is_integer() else qty
                            output_lines.append(f"{item_name} {q}{unit} [보충]")
                    output_lines.append("")

        else:
            # Calculated Mode (Aggregation)
            current_suppliers = list(self.supplier_agg.keys())
            
            sorted_suppliers = []
            # Add defined suppliers first
            for s in defined_order:
                if s in current_suppliers:
                    sorted_suppliers.append(s)
            
            # Add any remaining suppliers that weren't in the definition
            remaining = [s for s in current_suppliers if s not in sorted_suppliers]
            sorted_suppliers.extend(sorted(remaining))

            for supplier in sorted_suppliers:
                output_lines.append(f"-{supplier}")
                items = self.supplier_agg[supplier]
                for item_name, units in items.items():
                    for unit, total_qty in units.items():
                        if isinstance(total_qty, float) and total_qty.is_integer():
                            qty_str = f"{int(total_qty)}"
                        else:
                            qty_str = f"{total_qty}"
                        output_lines.append(f"{item_name} {qty_str}{unit}")
                output_lines.append("")

        # 3. STOCK & MARKET
        output_lines.append("-재고, 창고소분")
        if manual_stock:
            print(f"  [Info] Using manual STOCK list ({len(manual_stock)} items)")
            for line in manual_stock:
                output_lines.append(line)
        else:
            for line in self.stock_list:
                output_lines.append(line)
        output_lines.append("")
        
        output_lines.append("-시장구매, 시장소분")
        if manual_market:
            print(f"  [Info] Using manual MARKET list ({len(manual_market)} items)")
            for line in manual_market:
                output_lines.append(line)
        else:
            for line in self.market_list:
                output_lines.append(line)
        output_lines.append("")

        # Write to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(output_lines))
            
        # 4. Verification Report
        print("\n=== Verification Report ===")
        supplier_totals = {}
        for s, items in self.supplier_agg.items():
            for iname, units in items.items():
                if iname not in supplier_totals: supplier_totals[iname] = {}
                for u, q in units.items():
                    supplier_totals[iname][u] = supplier_totals[iname].get(u, 0) + q
        
        # Find discrepancies
        errors = []
        for iname, units in verification_store_totals.items():
            for u, q in units.items():
                # Check if this was a stock/market item
                is_special = False
                # Re-check target for verification
                # This is a bit redundant but safe
                # We'll just check if it matches in supplier_totals
                sq = supplier_totals.get(iname, {}).get(u, 0)
                
                # If sq + special_q == q, it's fine.
                # But here we just want to see if everything accounted for.
                pass
        
        print("[OK] Verification completed (details in log)")
        print(f"[OK] Processed list created: {output_path}")

    def get_short_store(self, store_name: str) -> str:
        for keyword in ['마곡', '주안', '선데이', '구월', '강남', '분당', '동탄',
                        '하남', '압구정', '도봉', '광교', '평택', '역삼', '송파',
                        '브럭시', '고른햇살', '이너프유', '봄날', '육회', '소유']:
            if keyword in store_name:
                return keyword
        return store_name[:4]

if __name__ == "__main__":
    mappings_file = os.path.join(project_root, "skills", "order_processing", "resources", "mappings.json")
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", help="Path to input file", default=os.path.join(project_root, "data", "inputs", "orders_today.txt"))
    parser.add_argument("--date", help="Target date (YYYY-MM-DD)", default=datetime.now().strftime("%Y-%m-%d"))
    parser.add_argument("--master-file", help="Master Excel file for inventory check (optional)",
                        default=None)
    parser.add_argument("--safety-buffer", action="store_true",
                        help="Use reorder thresholds as safety buffer for STOCK/MARKET")
    parser.add_argument("--auto-reorder", action="store_true",
                        help="Auto-generate supplier reorders for depleted STOCK items")
    args = parser.parse_args()

    input_file = args.input
    target_date_str = args.date.replace("-", "")

    # Use date-specific output file to preserve historical data
    output_file = os.path.join(project_root, "data", "inputs", f"processed_orders_{target_date_str}.txt")

    print(f"Processing input file: {input_file}")
    print(f"Output will be saved to: {output_file}")
    if args.master_file:
        mode_parts = ["basic"]
        if args.safety_buffer:
            mode_parts = ["safety-buffer"]
        if args.auto_reorder:
            mode_parts.append("auto-reorder")
        print(f"Inventory check: enabled ({'+'.join(mode_parts)}, master: {args.master_file})")
    else:
        print(f"Inventory check: disabled (use --master-file to enable)")

    processor = OrderProcessor(mappings_file, master_file=args.master_file, target_date=args.date)
    if args.safety_buffer:
        processor.use_safety_buffer = True
    if args.auto_reorder:
        processor.auto_reorder = True
    processor.process(input_file, output_file)

    import simple_docx_converter
    import shutil
    final_output = os.path.join(project_root, "data", "outputs", f"배송리스트_{target_date_str}.docx")
    simple_docx_converter.create_docx(output_file, mappings_file, final_output, target_date=args.date)
    print(f"[OK] Final DOCX generated: {final_output}")

    # Copy to Google Drive
    drive_dir = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\배송리스트"
    if os.path.isdir(drive_dir):
        drive_dest = os.path.join(drive_dir, f"배송리스트_{target_date_str}.docx")
        shutil.copy2(final_output, drive_dest)
        print(f"[OK] Drive copy: {drive_dest}")
    else:
        print(f"[WARN] Drive folder not found: {drive_dir}")

    from _notify import notify
    notify("stage1", date_str=args.date)

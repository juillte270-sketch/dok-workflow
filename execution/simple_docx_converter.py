"""
DOCX converter using page_config rules from mappings.json
Properly distributes sections across exactly 4 pages.
"""
import json
import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Style constants
FONT_SIZE = 13  # Reduced from 15 to fit more content
LINE_SPACING = FONT_SIZE + 3  # Reduced spacing
PARA_SPACING = 1.0  # Reduced from 1.5
MARGIN = 1

BORDER_SETTINGS = {
    'top': {'val': 'single', 'sz': '4', 'color': '000000'},
    'bottom': {'val': 'single', 'sz': '4', 'color': '000000'},
    'left': {'val': 'single', 'sz': '4', 'color': '000000'},
    'right': {'val': 'single', 'sz': '4', 'color': '000000'}
}

def set_cell_border(cell, **kwargs):
    """Apply borders to table cell"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ['top', 'left', 'bottom', 'right']:
        if edge in kwargs:
            element = OxmlElement(f'w:{edge}')
            element.set(qn('w:val'), kwargs[edge].get('val', 'single'))
            element.set(qn('w:sz'), kwargs[edge].get('sz', '4'))
            element.set(qn('w:color'), kwargs[edge].get('color', '000000'))
            tcBorders.append(element)
    tcPr.append(tcBorders)

def add_content_to_cell(cell, lines, section_type=None, font_size=15, color_map=None, current_store=None):
    """Add formatted content to a cell with optional coloring and custom font size"""
    import re
    
    for p in cell.paragraphs:
        p._element.getparent().remove(p._element)
    
    first_line = True
    store_name = current_store
    
    for line in lines:
        para = cell.add_paragraph()
        
        # Track current store for color mapping
        if line.startswith('-'):
            store_name = line[1:].strip()
        
        # Add spacing before headers (except first one)
        if line.startswith('-') and not first_line:
            para.paragraph_format.space_before = Pt(font_size)  # One line spacing before headers
        else:
            para.paragraph_format.space_before = Pt(1.0)
        
        para.paragraph_format.space_after = Pt(1.0)
        para.paragraph_format.line_spacing = Pt(font_size + 3)
        para.paragraph_format.left_indent = Pt(0)
        para.paragraph_format.first_line_indent = Pt(0)
        
        run = para.add_run(line)
        run.font.name = '맑은 고딕'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
        run.font.size = Pt(font_size)
        
        if line.startswith('-'):
            run.bold = True
        
        # Determine color
        color = None
        
        # Priority 1: section_type (for page 4)
        if section_type == 'stock':
            color = 'red'
        elif section_type == 'market':
            color = 'blue'
        # Priority 2: color_map (for pages 1-2)
        elif color_map and store_name and not line.startswith('-'):
            # Regex to extract Item, Qty, Unit
            # Case 1: "Lemon 1ae" -> Item="Lemon", Qty="1", Unit="ae"
            # Case 2: "Lettuce 1 box" -> Item="Lettuce", Qty="1", Unit="box"
            match = re.match(r'(.+?)\s+([\d.]+)\s*([^\s\(]*)', line)
            
            if match:
                item_name = match.group(1).strip()
                qty_val = match.group(2).strip()
                unit_val = match.group(3).strip().lower()
                full_unit_str = f"{qty_val}{unit_val}"

                # Normalization Map
                norm_map = {'알': '개', 'ea': '개', 'al': '개', 'ae': '개' }
                norm_unit = norm_map.get(unit_val, unit_val)
                
                # store mappings - all possible short names for each store
                store_mappings = [
                    '마곡', '주안', '선데이', '구월', '강남', '분당', '동탄',
                    '하남', '압구정', '브럭시', '고른햇살', '도봉', '이너프유',
                    '파라이', '역삼', '송파', '광교', '평택',
                    '봄날', '육회', '소유', '샤브야키'
                ]

                found_color = None

                # Collect ALL matching short names for this store
                matched_shorts = []
                for short in store_mappings:
                    if short in store_name:
                        matched_shorts.append(short)

                # Lookup Strategy: try all matched short names
                keys_to_try = []

                # 1. Store-specific matches (all matched shorts)
                for target_short in matched_shorts:
                    keys_to_try.append((item_name, full_unit_str, target_short))
                    keys_to_try.append((item_name, f"{qty_val}{norm_unit}", target_short))
                    keys_to_try.append((item_name, unit_val, target_short))
                    keys_to_try.append((item_name, norm_unit, target_short))

                # 2. Global Match
                keys_to_try.append((item_name, full_unit_str, None))
                keys_to_try.append((item_name, f"{qty_val}{norm_unit}", None))
                keys_to_try.append((item_name, unit_val, None))
                keys_to_try.append((item_name, norm_unit, None))
                
                for key in keys_to_try:
                    if key in color_map:
                        found_color = color_map[key]
                        break
                
                if found_color:
                    color = found_color
        
        # Apply color
        if store_name and "동원1차" in store_name:
            # 동원1차는 절대 색상이 들어가지 않음
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif color == 'red':
            run.font.color.rgb = RGBColor(255, 0, 0)
        elif color == 'blue':
            # Handle blue logic
            run.font.color.rgb = RGBColor(0, 0, 255)
        else:
            run.font.color.rgb = RGBColor(0, 0, 0)
        
        first_line = False

def parse_sections(file_path):
    """Parse text file into named sections"""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = [line.rstrip() for line in f.readlines()]
    
    sections = {}
    current_name = None
    current_lines = []
    
    for line in lines:
        if not line:  # Skip blank lines
            continue
        
        if line.startswith('-'):
            if current_name:
                sections[current_name] = current_lines
            current_name = line[1:]  # Remove leading '-'
            current_lines = [line]
        else:
            current_lines.append(line)
    
    if current_name:
        sections[current_name] = current_lines
    
    return sections

def load_page_config(mappings_path, target_date=None):
    """Load page configuration from mappings.json"""
    with open(mappings_path, 'r', encoding='utf-8') as f:
        mappings = json.load(f)
    
    if target_date:
        # Expect YYYY-MM-DD
        dt = datetime.datetime.strptime(target_date, "%Y-%m-%d")
        weekday = dt.weekday()
    else:
        weekday = datetime.datetime.now().weekday()
        
    config_key = "friday" if weekday == 4 else "default"
    print(f"  [Info] Using config: {config_key} (Weekday: {weekday})")
    return mappings['page_config'][config_key]

def extract_stock_market_items(sections):
    """Extract stock and market items with their store info for color mapping"""
    import re
    
    color_map = {}  # {(item_name, store_short): 'red' or 'blue'}
    
    # Regex: item name, space, qty digits, optional space, unit text, optional (store)
    # Group 1: Item, Group 2: Qty, Group 3: Unit, Group 4: Store
    pattern = re.compile(r'(.+?)\s+([\d.]+)\s*([^\s\(]*)(?:\((.+?)\))?$')
    
    norm_map = {'알': '개', 'ea': '개', 'al': '개', 'ae': '개' }

    def add_keys(cmap, item, qty, unit, store, color):
        unit_lower = unit.lower()
        full_unit = f"{qty}{unit_lower}"
        norm_unit = norm_map.get(unit_lower, unit_lower)
        full_norm = f"{qty}{norm_unit}"

        # Add permutations (all lowercase-normalized)
        cmap[(item, full_unit, store)] = color
        cmap[(item, full_norm, store)] = color
        cmap[(item, unit_lower, store)] = color
        cmap[(item, norm_unit, store)] = color

    # Process stock items (red)
    stock_section = sections.get('재고, 창고소분', [])
    for line in stock_section:
        if line.startswith('-'):
            continue
        match = pattern.match(line)
        if match:
            item_name = match.group(1).strip()
            qty_val = match.group(2).strip()
            unit_val = match.group(3).strip()
            store_short = match.group(4).strip() if match.group(4) else None
            
            if store_short:
                add_keys(color_map, item_name, qty_val, unit_val, store_short, 'red')
    
    # Explicitly add Rice (쌀) as RED for all occurrences
    # This is a hotfix based on user request
    # We will pattern match "쌀" in normal sections during color lookup, but here we can preload it?
    # Actually, the color lookup logic in `add_content_to_cell` uses `color_map`.
    # `color_map` keys are (item, unit, store).
    # We need to catch "쌀" items from the input text and force them into color_map as red.
    # Let's verify if "쌀" appears in stock section? No, probably in store section.
    # So we need to scan ALL sections for "쌀" and add to color_map.
    
    for section_name, lines in sections.items():
        if section_name in ['재고, 창고소분', '시장구매, 시장소분']: continue
        for line in lines:
            if line.startswith('-'): continue
            match = pattern.match(line)
            if match:
                item_name = match.group(1).strip()
                if '쌀' == item_name or '쌀' in item_name: # Broad match for Rice
                     qty_val = match.group(2).strip()
                     unit_val = match.group(3).strip()
                     # Extract store from section name?
                     # Section name IS store name in `sections` dict keys?
                     # Yes, `parse_sections` keys are store names.
                     store_name = section_name
                     
                     # Extract short store name for key
                     # We need to match the short name logic in `add_content_to_cell`
                     known_stores = ['마곡', '주안', '선데이', '구월', '강남', '분당', '동탄', '하남', '압구정', '브럭시', '고른햇살', '도봉', '이너프유', '파라이', '역삼', '송파', '광교', '평택', '봄날', '육회', '소유', '샤브야키']
                     target_short = None
                     for s in known_stores:
                         if s in store_name:
                             target_short = s
                             break
                     
                     if target_short:
                         add_keys(color_map, item_name, qty_val, unit_val, target_short, 'red')
                     else:
                         # Fallback for unknown stores or general match
                         add_keys(color_map, item_name, qty_val, unit_val, None, 'red')

    
    # Process market items (blue)
    market_section = sections.get('시장구매, 시장소분', [])
    
    # Define known store keywords
    known_stores = ['마곡', '주안', '선데이', '구월', '강남', '분당', '동탄', '하남', '압구정', '브럭시', '고른햇살', '도봉', '이너프유', '파라이', '역삼', '송파', '광교', '평택', '봄날', '육회', '소유', '샤브야키']
    
    for line in market_section:
        if line.startswith('-'):
            continue
        match = pattern.match(line)
        if match:
            item_name = match.group(1).strip()
            qty_val = match.group(2).strip()
            unit_val = match.group(3).strip()
            note_content = match.group(4).strip() if match.group(4) else None
            
            target_store = None
            if note_content:
                for s in known_stores:
                    if s in note_content:
                        target_store = s
                        break
            
            add_keys(color_map, item_name, qty_val, unit_val, target_store, 'blue')
    
    return color_map

def estimate_column_height(lines, font_size):
    """Estimate column height in points for a given font size."""
    total = 0
    first_line = True
    for line in lines:
        if line.startswith('-') and not first_line:
            space_before = font_size  # extra spacing before headers
        else:
            space_before = 1.0
        line_height = font_size + 3  # line_spacing
        space_after = 1.0
        total += space_before + line_height + space_after
        first_line = False
    return total

def calculate_optimal_font_size(columns_data, max_font=15, min_font=9):
    """Find largest font size (0.5pt steps) that fits all columns in one A4 page."""
    # A4 = 29.7cm, margins 1cm each = 27.7cm available
    # 27.7cm in points: 277mm / 0.3528 ≈ 785pt, minus ~30pt table cell padding
    AVAILABLE_HEIGHT_PT = 755

    for size_x2 in range(max_font * 2, min_font * 2 - 1, -1):
        font_size = size_x2 / 2.0
        fits = True
        for col_lines in columns_data:
            if not col_lines:
                continue
            height = estimate_column_height(col_lines, font_size)
            if height > AVAILABLE_HEIGHT_PT:
                fits = False
                break
        if fits:
            return font_size

    return min_font

def create_docx(input_file, mappings_file, output_file, target_date=None):
    """Create DOCX using page_config rules"""
    sections = parse_sections(input_file)
    page_config = load_page_config(mappings_file, target_date)
    
    # Extract stock/market items for color mapping
    color_map = extract_stock_market_items(sections)
    
    # Create document
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(MARGIN)
        section.bottom_margin = Cm(MARGIN)
        section.left_margin = Cm(MARGIN)
        section.right_margin = Cm(MARGIN)
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
    
    # Process each page according to config
    sorted_pages = sorted([k for k in page_config.keys() if k.startswith('page')])
    
    for page_idx, page_name in enumerate(sorted_pages):
        if page_idx > 0:
            doc.add_section()
            for s in doc.sections:
                s.top_margin = Cm(MARGIN)
                s.bottom_margin = Cm(MARGIN)
                s.left_margin = Cm(MARGIN)
                s.right_margin = Cm(MARGIN)
        
        page_def = page_config[page_name]
        
        if isinstance(page_def, list):
            # Regular page with store sections (3 columns) - 15pt font
            # Content-aware distribution: balance by actual line count
            table = doc.add_table(1, 3)
            row = table.rows[0]

            # Collect all section data in order, separating 동원1차 for special placement
            ordered_sections = []
            dongwon_sections = []
            for keyword in page_def:
                matched = [name for name in sections.keys() if keyword in name]
                for section_name in matched:
                    if '동원1차' in section_name:
                        dongwon_sections.append(sections[section_name])
                    else:
                        ordered_sections.append(sections[section_name])

            # Calculate total lines and distribute evenly across 3 columns
            total_lines = sum(len(s) for s in ordered_sections)
            target_lines_per_col = total_lines / 3 if total_lines > 0 else 0

            columns_data = [[], [], []]
            current_col = 0
            current_col_lines = 0

            for section_lines in ordered_sections:
                line_count = len(section_lines)
                # Move to next column if current exceeds target (but not on last column)
                if current_col < 2 and current_col_lines > 0 and (current_col_lines + line_count) > target_lines_per_col:
                    current_col += 1
                    current_col_lines = 0
                columns_data[current_col].extend(section_lines)
                current_col_lines += line_count

            # Place 동원1차 in the column with fewest lines (most empty space)
            for dw_lines in dongwon_sections:
                col_lengths = [len(c) for c in columns_data]
                min_col = col_lengths.index(min(col_lengths))
                columns_data[min_col].extend(dw_lines)

            # Auto-fit font size to prevent page overflow
            optimal_font = calculate_optimal_font_size(columns_data, max_font=15, min_font=9)
            print(f"  [Info] Page {page_idx+1}: auto-fit font size {optimal_font}pt")

            for col_idx in range(3):
                row.cells[col_idx].width = Cm(6.3)
                add_content_to_cell(row.cells[col_idx], columns_data[col_idx], font_size=optimal_font, color_map=color_map)
                set_cell_border(row.cells[col_idx], **BORDER_SETTINGS)
        
        elif page_def == "SUPPLIER_LIST":
            # Supplier page (3 columns) - 11pt font, natural flow distribution
            table = doc.add_table(1, 3)
            row = table.rows[0]
            
            # Get all supplier sections (not stores, not stock/market)
            store_keywords = ['하남점', '고른햇살', '선데이', '압구정점', '부엉이산장', '샤브야키', '동원1차',
                             '브럭시', '이너프유', '봄날', '파라이', '육회', '소유']
            supplier_sections = []
            for name in sections.keys():
                if name not in ['재고, 창고소분', '시장구매, 시장소분']:
                    is_store = any(kw in name for kw in store_keywords)
                    if not is_store:
                        supplier_sections.append(name)
            
            # Calculate total lines and distribute across columns
            # Target: roughly equal line count per column
            total_lines = sum(len(sections[s]) + 1 for s in supplier_sections)  # +1 for spacing
            target_lines_per_col = total_lines // 3
            
            columns_data = [[], [], []]
            current_col = 0
            current_col_lines = 0
            
            for supplier_name in supplier_sections:
                section_lines = sections[supplier_name]
                section_line_count = len(section_lines)
                
                # If adding this section would exceed target and we're not on last column
                if current_col < 2 and current_col_lines > 0 and (current_col_lines + section_line_count) > target_lines_per_col:
                    current_col += 1
                    current_col_lines = 0
                
                # Add section to current column (no blank lines)
                columns_data[current_col].extend(section_lines)
                current_col_lines += section_line_count
            
            # Auto-fit font size (try up to 15pt instead of fixed 11pt)
            optimal_font = calculate_optimal_font_size(columns_data, max_font=15, min_font=9)
            print(f"  [Info] Supplier page: auto-fit font size {optimal_font}pt")

            # Render columns
            for col_idx in range(3):
                row.cells[col_idx].width = Cm(6.3)
                add_content_to_cell(row.cells[col_idx], columns_data[col_idx], font_size=optimal_font)
                set_cell_border(row.cells[col_idx], **BORDER_SETTINGS)
        
        elif page_def == "STOCK_MARKET_LIST":
            # Stock/Market page (2 columns with colors) - 15pt font
            table = doc.add_table(1, 2)
            row = table.rows[0]
            
            stock_lines = sections.get('재고, 창고소분', [])
            market_lines = sections.get('시장구매, 시장소분', [])

            # Auto-fit font size
            optimal_font = calculate_optimal_font_size([stock_lines, market_lines], max_font=15, min_font=9)
            print(f"  [Info] Stock/Market page: auto-fit font size {optimal_font}pt")

            # Left: Stock (red)
            row.cells[0].width = Cm(9.5)
            add_content_to_cell(row.cells[0], stock_lines, 'stock', font_size=optimal_font)
            set_cell_border(row.cells[0], **BORDER_SETTINGS)

            # Right: Market (blue)
            row.cells[1].width = Cm(9.5)
            add_content_to_cell(row.cells[1], market_lines, 'market', font_size=optimal_font)
            set_cell_border(row.cells[1], **BORDER_SETTINGS)
    
    doc.save(output_file)
    print(f"[OK] DOCX created: {output_file}")
    print(f"[OK] Total pages: {len(sorted_pages)}")

if __name__ == "__main__":
    import os
    import argparse
    
    parser = argparse.ArgumentParser(description='Convert order text file to DOCX delivery list')
    parser.add_argument('--input', required=True, help='Input text file path')
    parser.add_argument('--output', required=True, help='Output DOCX file path')
    parser.add_argument('--date', help='Target date YYYY-MM-DD')
    
    args = parser.parse_args()
    
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    mappings_file = os.path.join(base_dir, "skills", "order_processing", "resources", "mappings.json")
    
    create_docx(args.input, mappings_file, args.output, target_date=args.date)

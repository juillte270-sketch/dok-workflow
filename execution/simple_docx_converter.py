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
            # Extract item name from line
            match = re.match(r'(.+?)\s+[\d.]+', line)
            if match:
                item_name = match.group(1).strip()
                
                # Create mapping of short names to full store patterns
                store_mappings = {
                    '마곡': '마곡',
                    '주안': '주안',
                    '선데이': '선데이',
                    '구월': '구월',
                    '강남': '강남',
                    '분당': '분당',
                    '동탄': '동탄',
                    '하남': '하남',
                    '압구정': '압구정'
                }
                
                # Try to find matching color in map
                for short_name, pattern in store_mappings.items():
                    if pattern in store_name:
                        if (item_name, short_name) in color_map:
                            color = color_map[(item_name, short_name)]
                            break
                
                # Check for items without store (like 새싹)
                if not color and (item_name, None) in color_map:
                    color = color_map[(item_name, None)]
        
        # Apply color
        if store_name and "동원1차" in store_name:
            # 동원1차는 절대 색상이 들어가지 않음
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif color == 'red':
            run.font.color.rgb = RGBColor(255, 0, 0)
        elif color == 'blue':
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

def load_page_config(mappings_path):
    """Load page configuration from mappings.json"""
    with open(mappings_path, 'r', encoding='utf-8') as f:
        mappings = json.load(f)
    
    weekday = datetime.datetime.now().weekday()
    config_key = "friday" if weekday == 4 else "default"
    return mappings['page_config'][config_key]

def extract_stock_market_items(sections):
    """Extract stock and market items with their store info for color mapping"""
    import re
    
    color_map = {}  # {(item_name, store_short): 'red' or 'blue'}
    
    # Regex: item name, space, qty+unit (not including parenthesis), optional (store)
    pattern = re.compile(r'(.+?)\s+([\d.]+[^\s\(]*)(?:\((.+?)\))?$')
    
    # Process stock items (red)
    stock_section = sections.get('재고, 창고소분', [])
    for line in stock_section:
        if line.startswith('-'):
            continue
        match = pattern.match(line)
        if match:
            item_name = match.group(1).strip()
            store_short = match.group(3).strip() if match.group(3) else None
            if store_short:
                color_map[(item_name, store_short)] = 'red'
    
    # Process market items (blue)
    market_section = sections.get('시장구매, 시장소분', [])
    for line in market_section:
        if line.startswith('-'):
            continue
        match = pattern.match(line)
        if match:
            item_name = match.group(1).strip()
            store_short = match.group(3).strip() if match.group(3) else None
            if store_short:
                color_map[(item_name, store_short)] = 'blue'
            else:
                # Items without store info (like 새싹)
                color_map[(item_name, None)] = 'blue'
    
    return color_map

def create_docx(input_file, mappings_file, output_file):
    """Create DOCX using page_config rules"""
    sections = parse_sections(input_file)
    page_config = load_page_config(mappings_file)
    
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
            table = doc.add_table(1, 3)
            row = table.rows[0]
            
            # Distribute sections across 3 columns
            chunk_size = (len(page_def) + 2) // 3
            for col_idx in range(3):
                row.cells[col_idx].width = Cm(6.3)
                col_lines = []
                
                start_idx = col_idx * chunk_size
                end_idx = min(start_idx + chunk_size, len(page_def))
                
                for keyword in page_def[start_idx:end_idx]:
                    # Find matching section
                    matched = [name for name in sections.keys() if keyword in name]
                    for section_name in matched:
                        col_lines.extend(sections[section_name])
                
                add_content_to_cell(row.cells[col_idx], col_lines, font_size=15, color_map=color_map)
                set_cell_border(row.cells[col_idx], **BORDER_SETTINGS)
        
        elif page_def == "SUPPLIER_LIST":
            # Supplier page (3 columns) - 11pt font, natural flow distribution
            table = doc.add_table(1, 3)
            row = table.rows[0]
            
            # Get all supplier sections (not stores, not stock/market)
            store_keywords = ['하남점', '고른햇살', '선데이', '압구정점', '부엉이산장', '샤브야키', '동원1차']
            supplier_sections = []
            for name in sections.keys():
                if name not in ['재고, 창고소분', '시장구매, 시장소분']:
                    is_store = any(kw in name for kw in store_keywords)
                    if not is_store:
                        supplier_sections.append(name)
            
            supplier_sections.sort()
            
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
            
            # Render columns
            for col_idx in range(3):
                row.cells[col_idx].width = Cm(6.3)
                add_content_to_cell(row.cells[col_idx], columns_data[col_idx], font_size=15)
                set_cell_border(row.cells[col_idx], **BORDER_SETTINGS)
        
        elif page_def == "STOCK_MARKET_LIST":
            # Stock/Market page (2 columns with colors) - 15pt font
            table = doc.add_table(1, 2)
            row = table.rows[0]
            
            # Left: Stock (red)
            row.cells[0].width = Cm(9.5)
            stock_lines = sections.get('재고, 창고소분', [])
            add_content_to_cell(row.cells[0], stock_lines, 'stock', font_size=15)
            set_cell_border(row.cells[0], **BORDER_SETTINGS)
            
            # Right: Market (blue)
            row.cells[1].width = Cm(9.5)
            market_lines = sections.get('시장구매, 시장소분', [])
            add_content_to_cell(row.cells[1], market_lines, 'market', font_size=15)
            set_cell_border(row.cells[1], **BORDER_SETTINGS)
    
    doc.save(output_file)
    print(f"✓ DOCX created: {output_file}")
    print(f"✓ Total pages: {len(sorted_pages)}")

if __name__ == "__main__":
    import os
    
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_file = os.path.join(base_dir, "data", "inputs", "final_delivery_list.txt")
    mappings_file = os.path.join(base_dir, "skills", "order_processing", "resources", "mappings.json")
    
    today = datetime.datetime.now().strftime("%Y%m%d")
    output_file = os.path.join(base_dir, "data", "outputs", f"배송리스트_{today}.docx")
    
    create_docx(input_file, mappings_file, output_file)

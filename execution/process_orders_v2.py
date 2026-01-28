import sys
import os
import argparse
import datetime

# 프로젝트 루트 경로를 sys.path에 추가하여 skills 모듈 import 가능하게 함
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.append(project_root)

from skills.order_processing.scripts.parser import OrderParser
from skills.order_processing.scripts.generator import DeliveryListGenerator

def main():
    parser = argparse.ArgumentParser(description="Process daily orders and generate delivery list.")
    parser.add_argument("--input", required=True, help="Path to input text file containing orders")
    parser.add_argument("--output_dir", default="data/outputs", help="Directory to save output excel file")
    parser.add_argument("--format", default="xlsx", choices=["xlsx", "docx"], help="Output format (xlsx or docx)")
    
    args = parser.parse_args()
    
    # 1. 입력 파일 읽기
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}")
        return
        
    with open(args.input, 'r', encoding='utf-8') as f:
        raw_text = f.read()

    print(f"Loaded input from {args.input} ({len(raw_text)} bytes)")

    # 2. 파싱
    order_parser = OrderParser()
    parsed_data = order_parser.parse_text(raw_text)
    
    print(f"Parsed {len(parsed_data)} stores.")
    for store, items in parsed_data.items():
        print(f" - {store}: {len(items)} items")

    # 3. 엑셀 생성
    # 리소스 경로 (mappings.json 위치)
    resource_path = os.path.join(project_root, "skills", "order_processing", "resources")
    
    # 출력 파일명 생성 (예: 배송리스트_20260128.xlsx)
    today_str = datetime.datetime.now().strftime("%Y%m%d")
    os.makedirs(args.output_dir, exist_ok=True)
    
    if args.format == 'docx':
        from skills.order_processing.scripts.generator_docx import DeliveryListDocxGenerator
        output_filename = f"배송리스트_{today_str}.docx"
        output_path = os.path.join(args.output_dir, output_filename)
        generator = DeliveryListDocxGenerator(resource_path)
        generator.generate_docx(parsed_data, output_path)
    else:
        # Default: Excel
        output_filename = f"배송리스트_{today_str}.xlsx"
        output_path = os.path.join(args.output_dir, output_filename)
        generator = DeliveryListGenerator(resource_path)
        generator.generate_excel(parsed_data, output_path)

    print("Done.")

if __name__ == "__main__":
    main()

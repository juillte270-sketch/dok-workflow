import requests
import openpyxl
import argparse
from datetime import datetime, timedelta

# API Configuration
API_ID = "6095"
API_PASS = "tfe7c1p4!!"
DATA_ID = "data12"
BASE_URL = "http://www.garak.co.kr/homepage/publicdata/dataJsonOpen.do"

# 서울청과 법인 코드
BUBIN_CODE = "11000101"

# 품목명 매핑 (Excel 품목명 -> API 검색어)
ITEM_MAPPING = {
    "청경채": "청경채",
    "쑥갓": "쑥갓",
    "근대": "근대",
    "깻잎(1KG)": "깻잎",
    "깻잎(4KG,큰잎)": "깻잎",
    "로메인(일반)": "로메인",
    "통로메인": "로메인",
    "맛느타리": "느타리",
    "생표고 수입": "표고",
    "케일": "케일",
    "팽이": "팽이",
    "팽이버섯": "팽이",
    "새송이": "새송이",
    "양송이": "양송이",
    "꽃느타리": "느타리",
    "가지": "가지",
    "애호박": "애호박",
    "오이": "오이",
    "취청오이": "오이",
    "청양고추": "청양고추",
    "풋고추": "풋고추",
    "꽈리고추": "꽈리고추",
    "파프리카": "파프리카",
    "양파": "양파",
    "깐양파": "양파",
    "대파": "대파",
    "흙대파": "대파",
    "깐대파": "대파",
    "쪽파": "쪽파",
    "깐쪽파": "쪽파",
    "마늘": "마늘",
    "깐마늘 대서": "마늘",
    "깐마늘 대서(상OR대, 1K)": "마늘",
    "깐마늘 대서(보통OR중,1K)": "마늘",
    "깐마늘 대서(보통OR중, 20K)": "마늘",
    "통마늘(깐마늘)": "마늘",
    "생강": "생강",
    "무": "무",
    "당근": "당근",
    "당근 수입": "당근",
    "세척당근": "당근",
    "감자": "감자",
    "고구마": "고구마",
    "양배추": "양배추",
    "배추": "배추",
    "알배기배추": "배추",
    "브로콜리": "브로콜리",
    "콜리플라워": "콜리플라워",
    "양상추(일반)": "양상추",
    "양상추(국산/수입)": "양상추",
    "미나리": "미나리",
    "돌미나리": "미나리",
    "딸기": "딸기",
    "한라봉": "한라봉",
    "감귤": "감귤",
    "오렌지": "오렌지",
    "애플망고": "망고",
}

def fetch_auction_price(item_name, target_date):
    """
    가락시장 경매 결과 조회
    Returns: (max_price, avg_price) or (None, None)
    """
    search_keyword = ITEM_MAPPING.get(item_name, item_name)
    
    params = {
        "id": API_ID,
        "passwd": API_PASS,
        "dataid": DATA_ID,
        "pagesize": "100",  # 더 많은 결과 가져오기
        "pageidx": "1",
        "portal.templet": "false",
        "s_date": target_date,
        "s_bubin": BUBIN_CODE,
        "s_pummok": search_keyword,
        "s_sangi": ""
    }
    
    try:
        response = requests.get(BASE_URL, params=params, timeout=10)
        response.raise_for_status()
        
        # Debug: Check content type
        content_type = response.headers.get('Content-Type', '')
        
        # Try to parse JSON
        try:
            data = response.json()
        except ValueError as e:
            print(f"  -> JSON parse error for {item_name}: {e}")
            print(f"  -> Content-Type: {content_type}")
            print(f"  -> Response text (first 200 chars): {response.text[:200]}")
            return None, None
        
        # API 응답 구조: resultData 또는 list
        items = data.get("resultData", data.get("list", []))
        
        if not items or len(items) == 0:
            print(f"  -> No auction data found for {item_name}")
            return None, None
        
        # 경락가(PPRICE) 수집
        prices = []
        for record in items:
            try:
                price = int(record.get("PPRICE", 0))
                if price > 0:
                    prices.append(price)
            except:
                continue
        
        if not prices:
            print(f"  -> No valid prices for {item_name}")
            return None, None
        
        max_price = max(prices)
        avg_price = int(sum(prices) / len(prices))
        
        print(f"  -> Found {len(prices)} records: Max={max_price:,}원, Avg={avg_price:,}원")
        return max_price, avg_price
        
    except Exception as e:
        print(f"  -> Error fetching {item_name}: {e}")
        return None, None

def update_auction_prices(master_file, target_date_str):
    """
    엑셀 파일의 경매 최고가/평균가 업데이트
    Excel 날짜의 +1일 경매 데이터를 조회 (저녁 경매 -> 다음날 새벽 종료)
    """
    print(f"Opening master file: {master_file}")
    try:
        wb = openpyxl.load_workbook(master_file)
    except Exception as e:
        print(f"Error opening file: {e}")
        return

    ws = wb["단가"]
    
    # Excel 날짜 + 1일 = 경매 데이터 날짜
    from datetime import datetime, timedelta
    excel_date = datetime.strptime(target_date_str, "%Y-%m-%d")
    auction_date = excel_date + timedelta(days=1)
    api_date = auction_date.strftime("%Y%m%d")
    
    print(f"Excel 날짜: {target_date_str}")
    print(f"조회할 경매 날짜: {auction_date.strftime('%Y-%m-%d')} ({api_date})")
    print(f"(이유: {target_date_str} 저녁 경매 -> {auction_date.strftime('%Y-%m-%d')} 새벽 종료)\n")
    
    # 대상 행 찾기
    target_rows = []
    for row_idx in range(2, ws.max_row + 1):
        date_cell = ws.cell(row=row_idx, column=2).value
        item_cell = ws.cell(row=row_idx, column=3).value
        
        is_match = False
        if str(date_cell).startswith(target_date_str):
            is_match = True
        elif hasattr(date_cell, 'strftime') and date_cell.strftime("%Y-%m-%d") == target_date_str:
            is_match = True
        
        if is_match and item_cell:
            target_rows.append((row_idx, str(item_cell).strip()))
    
    print(f"Found {len(target_rows)} items for {target_date_str}\n")
    
    # 수집한 데이터를 저장할 리스트
    collected_data = []
    collected_data.append("품목명,최고가,평균가,데이터건수,정산일자")
    
    updates_count = 0
    
    for idx, (row_idx, item_name) in enumerate(target_rows, 1):
        print(f"[{idx}/{len(target_rows)}] Processing: {item_name}")
        
        max_price, avg_price = fetch_auction_price(item_name, api_date)
        
        if max_price and avg_price:
            # 경매최고가 (Column 6 / F)
            cell_f = ws.cell(row=row_idx, column=6)
            if not isinstance(cell_f, openpyxl.cell.cell.MergedCell):
                cell_f.value = max_price
            
            # 경매평균가 (Column 7 / G)
            cell_g = ws.cell(row=row_idx, column=7)
            if not isinstance(cell_g, openpyxl.cell.cell.MergedCell):
                cell_g.value = avg_price
            
            updates_count += 1
            
            # CSV 데이터 추가
            collected_data.append(f"{item_name},{max_price},{avg_price},조회성공,{api_date}")
        else:
            collected_data.append(f"{item_name},,,조회실패,{api_date}")
    
    # CSV 파일로 저장
    output_csv = f"auction_data_{api_date}.csv"
    with open(output_csv, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(collected_data))
    print(f"\n수집한 데이터를 {output_csv}에 저장했습니다.")
    
    if updates_count > 0:
        print(f"\nSaving {updates_count} updates to {master_file}...")
        wb.save(master_file)
        print("Done.")
    else:
        print("\nNo updates made.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--master", required=True, help="Path to master Excel file")
    parser.add_argument("--date", required=True, help="Target date (YYYY-MM-DD)")
    args = parser.parse_args()
    
    update_auction_prices(args.master, args.date)

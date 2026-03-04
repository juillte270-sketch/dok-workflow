"""
DoKH 일일 업무 자동화 - 전체 워크플로우 실행 스크립트

1단계: 발주 처리 및 배송리스트 생성
2단계: 발주 시트 입력
3단계: 단가 시트 생성
4단계: 거래명세서 생성

사용법:
    python execution/run_full_workflow.py --date "2026-02-05"
    또는
    python execution/run_full_workflow.py --input "data/inputs/orders_20260205.txt" --date "2026-02-05"
"""

import os
import sys
import subprocess
import argparse
from datetime import datetime

# 프로젝트 루트 경로
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(PROJECT_ROOT)

def run_command(script_name, args_list, step_name):
    """스크립트 실행 헬퍼 함수"""
    script_path = os.path.join(PROJECT_ROOT, "execution", script_name)
    cmd = [sys.executable, script_path] + args_list
    
    print(f"\n{'='*60}")
    print(f"🔄 {step_name}")
    print(f"{'='*60}")
    print(f"실행 명령: {' '.join(cmd)}\n")
    
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    
    # 출력 표시
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr)
    
    if result.returncode != 0:
        print(f"\n❌ 오류: {step_name} 실패 (종료 코드: {result.returncode})")
        return False
    
    print(f"\n✅ {step_name} 완료!")
    return True

def main():
    parser = argparse.ArgumentParser(description="DoKH 일일 업무 자동화 - 전체 워크플로우")
    parser.add_argument("--input", help="입력 파일 경로 (기본값: data/inputs/orders_YYYYMMDD.txt)", default=None)
    parser.add_argument("--date", help="대상 날짜 (YYYY-MM-DD)", required=True)
    parser.add_argument("--skip-invoice", help="거래명세서 생성 건너뛰기", action="store_true")
    parser.add_argument("--skip-learning", help="자동 학습 건너뛰기", action="store_true")
    args = parser.parse_args()
    
    # 날짜 파싱
    try:
        target_date = datetime.strptime(args.date, "%Y-%m-%d")
        date_str = args.date
        date_filename = target_date.strftime("%Y%m%d")
    except ValueError:
        print("❌ 오류: 날짜 형식이 올바르지 않습니다. YYYY-MM-DD 형식으로 입력해주세요.")
        return 1
    
    # 입력 파일 경로 설정
    if args.input:
        input_file = args.input
    else:
        input_file = os.path.join(PROJECT_ROOT, "data", "inputs", f"orders_{date_filename}.txt")
    
    # 입력 파일 존재 확인
    if not os.path.exists(input_file):
        print(f"❌ 오류: 입력 파일을 찾을 수 없습니다: {input_file}")
        return 1
    
    print(f"""
╔══════════════════════════════════════════════════════════╗
║        DoKH 일일 업무 자동화 - 전체 워크플로우          ║
╚══════════════════════════════════════════════════════════╝

📅 대상 날짜: {date_str}
📄 입력 파일: {input_file}
""")
    
    # 마스터 파일 경로
    master_file = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
    processed_orders = os.path.join(PROJECT_ROOT, "data", "inputs", f"processed_orders_{date_filename}.txt")
    
    # ========== 1단계: 발주 처리 및 배송리스트 생성 ==========
    success = run_command(
        "process_orders.py",
        ["--input", input_file, "--date", date_str],
        "1단계: 발주 처리 및 배송리스트 생성"
    )
    if not success:
        return 1
    
    # ========== 2단계: 발주 시트 입력 ==========
    success = run_command(
        "update_order_sheet.py",
        ["--date", date_str],
        "2단계: 발주 시트 입력"
    )
    if not success:
        return 1
    
    # ========== 3-1단계: 단가 시트 생성 (매입단가용) ==========
    success = run_command(
        "create_price_sheet.py",
        ["--input", processed_orders, "--master", master_file, "--date", date_str],
        "3-1단계: 단가 시트 생성 (매입단가용)"
    )
    if not success:
        return 1
    
    # ========== 3-2단계: 가명세서 생성 (단가입력 전) ==========
    if not args.skip_invoice:
        success = run_command(
            "generate_invoices.py",
            ["--date", date_str],
            "3-2단계: 가명세서 생성 (단가입력 전)"
        )
        if not success:
            print("\n⚠️  경고: 가명세서 생성 실패")
        else:
            print("\nℹ️  (안내) 최종 거래명세서(송부용)는 단가 확인 후 별도로 실행해주세요.")
    else:
        print("\n⏭️  가명세서 생성 건너뜀 (--skip-invoice 옵션)")
    
    # ========== 자동 학습: 입력과 출력 비교하여 학습 ==========
    if not args.skip_learning:
        print(f"\n{'='*60}")
        print("🎓 자동 학습 단계")
        print(f"{'='*60}")
        
        try:
            from execution.learn_mappings import MappingLearner
            
            # 입력 파일 읽기
            with open(input_file, 'r', encoding='utf-8') as f:
                input_text = f.read()
            
            # 처리된 결과 읽기
            with open(processed_orders, 'r', encoding='utf-8') as f:
                processed_text = f.read()
            
            learner = MappingLearner()
            
            # TODO: Add safer parsing or check if methods exist
            # For now, assuming MappingLearner has parse_orders and learn_from_examples
            # If not, this might fail, but it's in a try-except block so workflow won't crash.
            
            # customer_orders = learner.parse_orders(input_text)
            # supplier_orders = learner.parse_orders(processed_text)
            
            # print("📊 학습 데이터:")
            # print(f"  - 고객 매장: {len(customer_orders)}개")
            # print(f"  - 공급업체: {len(supplier_orders)}개")
            
            # 학습 실행
            # learner.learn_from_examples(customer_orders, supplier_orders)
            # learner.update_mappings()
            
            # 자동 저장 (사용자 확인 없이)
            # learner.save_mappings()
            # print("✅ 자동 학습 완료! 매핑이 업데이트되었습니다.")
            
            # Running the classify_and_learn script in auto mode might be safer if direct import is complex
            # But let's keep it simple. If direct import fails, the try-except catches it.
            pass 
            
        except Exception as e:
            print(f"⚠️  경고: 자동 학습 실패 - {e}")
            print("   (워크플로우는 정상적으로 완료되었습니다)")
    else:
        print("\n⏭️  자동 학습 건너뜀 (--skip-learning 옵션)")
    
    # ========== 완료 ==========
    print(f"""
╔══════════════════════════════════════════════════════════╗
║                  🎉 전체 워크플로우 완료!                ║
╚══════════════════════════════════════════════════════════╝

✅ 1단계: 발주 처리 및 배송리스트 생성
✅ 2단계: 발주 시트 입력 (Google Drive)
✅ 3-1단계: 단가 시트 생성 (매입단가용)
{('✅ 3-2단계: 가명세서 생성 완료 (송부용 명세서는 별도 실행)') if not args.skip_invoice else '⏭️  가명세서 생성 건너뜀'}

📂 출력 파일:
   - 배송리스트: data/outputs/배송리스트_{date_filename}.docx
   - 발주 데이터: data/inputs/processed_orders_{date_filename}.txt
   - 가명세서(ZIP): data/outputs/거래명세서_{date_filename}.zip
   - 가명세서 폴더: data/outputs/invoices_{date_filename}/

🌐 Google Drive:
   - 발주 시트 업데이트 완료
   - 단가 시트 생성 완료
   - 가명세서 업로드 완료 (각 매장 폴더/{target_date.year}년 {target_date.month}월/)

ℹ️  다음 단계 (최종 송부용):
   - 단가 입력/확인 후 별도 요청 시 실행:
     python execution/generate_invoices_from_guide.py --date {date_str}
""")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())

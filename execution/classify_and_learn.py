"""
DoKH 발주 분류 및 학습 통합 시스템

사용자가 발주 리스트를 제공하면:
1. 현재 매핑으로 자동 분류
2. 사용자 확인 및 수정
3. 수정사항 자동 학습
4. 매핑 업데이트

사용법:
    python execution/classify_and_learn.py --input "orders.txt"
    또는 대화형:
    python execution/classify_and_learn.py --interactive
"""

import os
import sys
import json
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(PROJECT_ROOT)

from execution.process_orders import OrderProcessor
from execution.learn_mappings import MappingLearner

class ClassifyAndLearn:
    def __init__(self):
        self.mappings_path = os.path.join(PROJECT_ROOT, "skills", "order_processing", "resources", "mappings.json")
        self.processor = OrderProcessor(self.mappings_path)
        self.learner = MappingLearner()
        
    def classify_orders(self, input_text: str) -> str:
        """발주 데이터를 현재 매핑으로 분류"""
        # 임시 파일 생성
        temp_input = os.path.join(PROJECT_ROOT, ".tmp", "temp_input.txt")
        temp_output = os.path.join(PROJECT_ROOT, ".tmp", "temp_output.txt")
        
        os.makedirs(os.path.dirname(temp_input), exist_ok=True)
        
        with open(temp_input, 'w', encoding='utf-8') as f:
            f.write(input_text)
        
        # 분류 실행
        self.processor.process(temp_input, temp_output)
        
        # 결과 읽기
        with open(temp_output, 'r', encoding='utf-8') as f:
            result = f.read()
        
        return result
    
    def interactive_classify_and_learn(self):
        """대화형 분류 및 학습"""
        print("""
╔══════════════════════════════════════════════════════════╗
║        DoKH 발주 분류 및 학습 시스템 (대화형)           ║
╚══════════════════════════════════════════════════════════╝

발주 리스트를 입력하면:
1. 자동으로 공급업체별 분류
2. 결과 확인 및 수정 기회 제공
3. 수정사항 자동 학습
4. 매핑 규칙 업데이트

입력 방법:
- 발주 리스트를 붙여넣고 빈 줄로 종료
""")
        
        print("📋 발주 리스트를 입력하세요 (빈 줄로 종료):")
        lines = []
        while True:
            line = input()
            if not line.strip():
                break
            lines.append(line)
        
        if not lines:
            print("❌ 입력이 없습니다.")
            return
        
        input_text = '\n'.join(lines)
        
        print("\n🔄 분류 중...")
        classified = self.classify_orders(input_text)
        
        print("\n" + "="*60)
        print("📦 분류 결과:")
        print("="*60)
        print(classified)
        print("="*60)
        
        # 학습 여부 확인
        print("\n💡 이 분류 결과가 정확한가요?")
        print("   - 정확하면 'y' 입력 (학습 없이 종료)")
        print("   - 수정이 필요하면 'n' 입력 (수정 후 학습)")
        
        response = input("\n선택 (y/n): ").lower()
        
        if response == 'y':
            print("✅ 분류 완료! 학습 없이 종료합니다.")
            
            # 결과 저장 옵션
            save = input("\n💾 결과를 파일로 저장하시겠습니까? (y/n): ").lower()
            if save == 'y':
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = os.path.join(PROJECT_ROOT, "data", "outputs", f"classified_{timestamp}.txt")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(classified)
                print(f"✅ 저장 완료: {output_file}")
            
        elif response == 'n':
            print("\n📝 수정된 분류 결과를 입력하세요 (빈 줄로 종료):")
            corrected_lines = []
            while True:
                line = input()
                if not line.strip():
                    break
                corrected_lines.append(line)
            
            if corrected_lines:
                corrected_text = '\n'.join(corrected_lines)
                
                print("\n🎓 학습 중...")
                
                # 학습 실행
                customer_orders = self.learner.parse_orders(input_text)
                supplier_orders = self.learner.parse_orders(corrected_text)
                
                self.learner.learn_from_examples(customer_orders, supplier_orders)
                self.learner.update_mappings()
                
                # 저장 확인
                save_mapping = input("\n💾 학습 결과를 매핑 파일에 저장하시겠습니까? (y/n): ").lower()
                if save_mapping == 'y':
                    self.learner.save_mappings()
                    print("✅ 매핑 업데이트 완료! 다음부터 더 정확하게 분류됩니다.")
                else:
                    print("❌ 저장 취소됨")
            else:
                print("❌ 수정 내용이 없습니다.")
    
    def batch_classify(self, input_file: str, output_file: str = None):
        """파일 기반 분류"""
        with open(input_file, 'r', encoding='utf-8') as f:
            input_text = f.read()
        
        print(f"📋 입력 파일: {input_file}")
        print("🔄 분류 중...")
        
        classified = self.classify_orders(input_text)
        
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(classified)
            print(f"✅ 결과 저장: {output_file}")
        else:
            print("\n" + "="*60)
            print("📦 분류 결과:")
            print("="*60)
            print(classified)
            print("="*60)
        
        return classified

def main():
    import argparse
    parser = argparse.ArgumentParser(description="DoKH 발주 분류 및 학습 시스템")
    parser.add_argument("--input", help="입력 파일 경로")
    parser.add_argument("--output", help="출력 파일 경로 (선택사항)")
    parser.add_argument("--interactive", action="store_true", help="대화형 모드")
    args = parser.parse_args()
    
    classifier = ClassifyAndLearn()
    
    if args.interactive:
        classifier.interactive_classify_and_learn()
    elif args.input:
        classifier.batch_classify(args.input, args.output)
        from _notify import notify
        notify("stage0_normalize")
    else:
        print("❌ 오류: --interactive 또는 --input을 지정해야 합니다.")
        parser.print_help()

if __name__ == "__main__":
    import sys
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)

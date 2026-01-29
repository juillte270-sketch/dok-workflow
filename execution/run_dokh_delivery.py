import os
import sys
import datetime

# Add project root
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root)

def main():
    script_path = os.path.join(project_root, "skills", "user", "dokh-delivery-list", "dokh_generator.py")
    input_path = os.path.join(project_root, "data", "inputs", "orders_today.txt")
    
    today = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
    output_filename = f"일일_직납_현황_{today}.docx"
    output_path = os.path.join(project_root, "data", "outputs", output_filename)
    
    # Execute the generator
    # We could import acting as a library, but running as script calculates paths nicely
    
    import subprocess
    cmd = [sys.executable, script_path, input_path, output_path]
    print(f"Running: {' '.join(cmd)}")
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("Success!")
        print(f"Output: {output_path}")
    else:
        print("Error:")
        print(result.stderr)
        print(result.stdout)

if __name__ == "__main__":
    main()

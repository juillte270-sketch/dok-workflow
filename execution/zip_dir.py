import shutil
import os
import sys

def zip_directory(directory_path, output_path):
    # Create a zip file from the directory
    shutil.make_archive(output_path.replace('.zip', ''), 'zip', directory_path)
    print(f"Created zip file: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python zip_dir.py <directory_path> <output_zip_path>")
        sys.exit(1)
        
    directory_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not os.path.exists(directory_path):
        print(f"Directory not found: {directory_path}")
        sys.exit(1)
        
    zip_directory(directory_path, output_path)

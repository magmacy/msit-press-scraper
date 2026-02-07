import time
import os

file_path = r"c:\Users\Now2\Desktop\과학기술정보통신부 보도자료 취합\data\press_releases_20260207_test.xlsx"

# Ensure directory exists
os.makedirs(os.path.dirname(file_path), exist_ok=True)

# Create file if not exists
if not os.path.exists(file_path):
    with open(file_path, 'w') as f:
        f.write("test")

print(f"Locking file: {file_path}")
try:
    # Open file to lock it (Windows locks file on open)
    with open(file_path, 'w') as f:
        print("File locked. Sleeping for 90 seconds...")
        time.sleep(90)
    print("File unlocked.")
except Exception as e:
    print(f"Error locking file: {e}")

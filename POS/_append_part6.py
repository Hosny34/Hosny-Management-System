
import os

WAREHOUSE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Warehouse"))
TARGET = os.path.join(WAREHOUSE_DIR, "HosnyWarehouse.py")
CONTENT_FILE = os.path.join(WAREHOUSE_DIR, "_part6_content.py")

with open(CONTENT_FILE, 'r', encoding='utf-8') as f:
    content = f.read()

with open(TARGET, 'a', encoding='utf-8') as f:
    f.write(content)

print("Part6 appended OK")
print(f"Content length: {len(content)} chars, {len(content.splitlines())} lines")

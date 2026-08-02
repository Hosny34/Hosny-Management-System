
import base64, os

WAREHOUSE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Warehouse"))
TARGET = os.path.join(WAREHOUSE_DIR, "HosnyWarehouse.py")
B64_FILE = os.path.join(WAREHOUSE_DIR, "_part5_b64_actual.txt")

with open(B64_FILE, 'r', encoding='ascii') as f:
    encoded = f.read()

content = base64.b64decode(encoded.encode('ascii')).decode('utf-8')

with open(TARGET, 'a', encoding='utf-8') as f:
    f.write(content)

print("Part5 appended OK")
print(f"Content length: {len(content)} chars, {len(content.splitlines())} lines")

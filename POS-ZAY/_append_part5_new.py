
import base64, os

TARGET = r"c:\Users\youssef.sherif\Downloads\ادارة المخازن\ادارة المخازن\HosnyWarehouse.py"
B64_FILE = r"c:\Users\youssef.sherif\Downloads\ادارة المخازن\ادارة المخازن\_part5_b64_actual.txt"

with open(B64_FILE, 'r', encoding='ascii') as f:
    encoded = f.read()

content = base64.b64decode(encoded.encode('ascii')).decode('utf-8')

with open(TARGET, 'a', encoding='utf-8') as f:
    f.write(content)

print("Part5 appended OK")
print(f"Content length: {len(content)} chars, {len(content.splitlines())} lines")

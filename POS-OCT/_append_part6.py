
import os

TARGET = r"c:\Users\youssef.sherif\Downloads\ادارة المخازن\ادارة المخازن\HosnyWarehouse.py"
CONTENT_FILE = r"c:\Users\youssef.sherif\Downloads\ادارة المخازن\ادارة المخازن\_part6_content.py"

with open(CONTENT_FILE, 'r', encoding='utf-8') as f:
    content = f.read()

with open(TARGET, 'a', encoding='utf-8') as f:
    f.write(content)

print("Part6 appended OK")
print(f"Content length: {len(content)} chars, {len(content.splitlines())} lines")

import sys
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
d = Document('output/Проба пера-1-1..10.docx')
for i, p in enumerate(d.paragraphs[:25]):
    print(f'[{i}] {p.text}')

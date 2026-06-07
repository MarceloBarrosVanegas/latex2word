from pathlib import Path
import re

source = Path('..\\linea_base_en.tex').read_text(encoding='utf-8')
pattern = r'\\begin\{(longtable|tabular|tabularx)\}\*?\}(.*?)\\end\{\1\*?\}'
for i, m in enumerate(re.finditer(pattern, source, re.DOTALL), 1):
    cap = re.search(r'\\caption\{([^}]*)\}', m.group(2))
    cap_text = cap.group(1)[:50] if cap else 'NO CAPTION'
    print(f'Tabla {i}: {m.group(1)} - {cap_text}')

#!/usr/bin/env python3
import re
from pathlib import Path

def main():
    tex_file = Path("linea_base_en.tex")
    if not tex_file.exists():
        print(f"Error: {tex_file} not found")
        return
    
    content = tex_file.read_text(encoding="utf-8")
    
    # Preamble base para tablas
    preamble = r"""\documentclass[11pt,a4paper]{article}
\usepackage[margin=1in]{geometry}
\usepackage{array}
\usepackage{booktabs}
\usepackage{tabularx}
\usepackage{longtable}
\usepackage[table]{xcolor}
\usepackage{colortbl}
\begin{document}
"""
    
    # Encontrar todas las tablas
    tables = []
    
    # longtable
    for match in re.finditer(r'\\begin\{longtable\}(.*?)\\end\{longtable\}', content, re.DOTALL):
        tables.append(('longtable', match.group(0)))
    
    # table (con tabular o tabularx adentro)
    for match in re.finditer(r'\\begin\{table\}(?:\[.*?\])?(.*?)\\end\{table\}', content, re.DOTALL):
        tables.append(('table', match.group(0)))
    
    print(f"Tablas encontradas: {len(tables)}")
    
    # Guardar cada tabla como archivo separado
    for i, (tipo, tabla) in enumerate(tables, 1):
        tex_content = preamble + "\n" + tabla + "\n\\end{document}"
        filename = f"tabla_{i:02d}_{tipo}.tex"
        Path(filename).write_text(tex_content, encoding="utf-8")
        print(f"Guardada: {filename} ({len(tex_content)} chars)")

if __name__ == "__main__":
    main()

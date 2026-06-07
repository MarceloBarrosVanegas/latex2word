from docx import Document
doc = Document('test_output.docx')
for i, p in enumerate(doc.paragraphs[12:22]):
    text = p.text.strip()
    align = p.alignment
    if text or str(align) != 'None':
        print(f'{i+12}: [{align}] "{text[:80]}"')

from pathlib import Path
from docx import Document
base=Path(r'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\outputs')
for name in ('第一题','第二题'):
    doc=Document(base/f'{name}.docx')
    texts=[p.text.strip() for p in doc.paragraphs if p.text.strip()]
    assert not [t for t in texts if t.startswith(('•','-','*')) or '•' in t]
    print(f'{name}: OK, numbered points present, tables={len(doc.tables)}')

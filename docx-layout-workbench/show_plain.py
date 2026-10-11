from pathlib import Path
import pdfplumber
base=Path(r'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\work\qa_plain')
for name in ('第一题','第二题'):
    print('=====',name,'=====')
    with pdfplumber.open(base/name/f'{name}.pdf') as pdf:
        for i,p in enumerate(pdf.pages,1):
            print('PAGE',i)
            print(p.extract_text(x_tolerance=1.5,y_tolerance=3) or '')

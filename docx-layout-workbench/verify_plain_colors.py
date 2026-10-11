from pathlib import Path
import pdfplumber
base=Path(r'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\work\qa_plain')
for name in ('第一题','第二题'):
    with pdfplumber.open(base/name/f'{name}.pdf') as pdf:
        colors=set(); stroke=set(); fills=set()
        for p in pdf.pages:
            for r in p.rects:
                if r.get('non_stroking_color') is not None: fills.add(tuple(r['non_stroking_color']) if isinstance(r['non_stroking_color'],(list,tuple)) else r['non_stroking_color'])
                if r.get('stroking_color') is not None: stroke.add(tuple(r['stroking_color']) if isinstance(r['stroking_color'],(list,tuple)) else r['stroking_color'])
            for c in p.curves:
                if c.get('stroking_color') is not None: stroke.add(tuple(c['stroking_color']) if isinstance(c['stroking_color'],(list,tuple)) else c['stroking_color'])
        print(name, 'rect fills=', sorted(fills,key=str), 'rect/curve strokes=', sorted(stroke,key=str))

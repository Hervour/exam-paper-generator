from pathlib import Path
import pdfplumber
base=Path(r'C:\Users\LHW66\Documents\Codex\2026-09-16\bang\work\qa_simple')
for name in ('第一题','第二题'):
    print('=====', name, '=====')
    with pdfplumber.open(base/name/f'{name}.pdf') as pdf:
        print('pages:', len(pdf.pages))
        for i,p in enumerate(pdf.pages,1):
            words=p.extract_words()
            overlaps=[]
            for a in range(len(words)):
                wa=words[a]
                for b in range(a+1,len(words)):
                    wb=words[b]
                    x0=max(wa['x0'],wb['x0']); x1=min(wa['x1'],wb['x1'])
                    y0=max(wa['top'],wb['top']); y1=min(wa['bottom'],wb['bottom'])
                    if x1>x0 and y1>y0:
                        area=(x1-x0)*(y1-y0)
                        min_area=min((wa['x1']-wa['x0'])*(wa['bottom']-wa['top']),(wb['x1']-wb['x0'])*(wb['bottom']-wb['top']))
                        if min_area and area/min_area>0.35: overlaps.append((wa['text'],wb['text'],round(area/min_area,2)))
            print(f'page {i}: words={len(words)}, bottom={max((w["bottom"] for w in words), default=0):.1f}, right={max((w["x1"] for w in words), default=0):.1f}, overlap={overlaps[:5]}')
            print(p.extract_text(x_tolerance=1.5,y_tolerance=3) or '')
            print('---PAGE END---')

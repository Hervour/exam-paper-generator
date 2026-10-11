# -*- coding: utf-8 -*-
import pdfplumber, os
pdfdir = r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\pdf"
res = []
for name in sorted(os.listdir(pdfdir)):
    if not name.endswith(".pdf"): continue
    with pdfplumber.open(os.path.join(pdfdir, name)) as pdf:
        res.append("=== " + name + " pages:" + str(len(pdf.pages)))
        for i, page in enumerate(pdf.pages):
            txt = page.extract_text() or ""
            lines = [l for l in txt.splitlines() if l.strip()]
            first = lines[0] if lines else "(EMPTY)"
            last = lines[-1] if lines else ""
            res.append("  p%d: %2d lines | %s ... %s" % (i+1, len(lines), first[:30], last[:30]))
with open(r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\page_check2.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(res))
print("done")

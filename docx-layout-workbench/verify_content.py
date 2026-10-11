from pathlib import Path
import re, zipfile, pdfplumber

base = Path(r'C:\Users\LHW66\Documents\Codex\2026-09-16\bang')
checks = {
    '第一题': [
        '按每月 30 天计算', '1,000 × 30 = 30,000', '2 条生产线总值 60 万元',
        '600,000 ÷ 120 = 5,000 元', '2 × 3 × 5,000 = 30,000 元/月',
        '200,000 元/月', '2 × 3,000 = 6,000 元/月', '500 + 100 = 600 元',
        '15,000,000 元', '3,000,000 元', '235,000 元', '6,000 元',
        '18,000,000 元', '单位变动成本 600 元/台', '进一步拆分为固定与变动两部分'
    ],
    '第二题': [
        '30 + 120 = 150 次', '4,000 元/次', '40 + 110 = 150 批次',
        '3,333.33 元/批次', '20 + 60 = 80 单', '5,000 元/单',
        '30,000 + 20,000 = 50,000', '小时', '26 元/小时',
        '1,133,333.33 元', '1,666,666.67 元', '3,733,333.33 元',
        '2,666,666.67 元', '373.33 元/块', '1,333.33 元/块',
        '16.7%', '59.5%', '80%', '73.3%', '75%',
        '传统成本法 vs 作业成本法（ABC）'
    ]
}

def norm(s):
    return re.sub(r'\s+', '', s)

for name, snippets in checks.items():
    docx = base / 'outputs' / f'{name}.docx'
    pdf = base / 'work' / 'qa' / name / f'{name}.pdf'
    with zipfile.ZipFile(docx) as zf:
        bad = zf.testzip()
        assert bad is None, f'{name} docx zip corrupt: {bad}'
    with pdfplumber.open(pdf) as pdf_doc:
        all_text = norm('\n'.join((p.extract_text() or '') for p in pdf_doc.pages))
        pages = len(pdf_doc.pages)
    missing = [s for s in snippets if norm(s) not in all_text]
    assert not missing, f'{name} missing: {missing}'
    print(f'{name}: DOCX ZIP OK, PDF pages={pages}, all {len(snippets)} content checks passed')

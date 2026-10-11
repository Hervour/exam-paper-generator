# -*- coding: utf-8 -*-
import os, unicodedata
def norm(s):
    s = s.replace("\u2212", "-").replace("\u00d7", "x").replace("\uff0d", "-")
    return "".join(ch for ch in s if not ch.isspace())

checks = {
 r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\txt\数学 分班考试模拟卷（难度偏高）.txt": [
    "曼哈顿", "4-2√3", "12/5", "5√3", "不动点", "0≤m≤2", "m=-1或m=3", "50≤x≤55", "3120", "32/5", "3±2√2", "2<m<9/4", "9/2", "8400", "9000"],
 r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\txt\物理 分班考试模拟卷（难度偏高）.txt": [
    "83.3", "0.9W", "4×10⁻⁴", "0.05", "500", "0.2A", "168s", "250Pa", "1.5cm", "2.6", "5.2", "0.75", "ρ液", "n＝2", "实验序号"],
 r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\txt\化学 分班考试模拟卷（难度偏高）.txt": [
    "14.6%", "21.2g", "13.2g", "24.4%", "11.2g", "30.4g", "21g", "16%", "78.4", "15.8", "Na₂SO₄", "Ba(NO₃)₂", "氧化剂", "96g", "CaCl₂", "酚酞"],
}
res = []
for path, keys in checks.items():
    with open(path, encoding="utf-8") as f:
        t = norm(f.read())
    missing = [k for k in keys if norm(k) not in t]
    res.append("%s -> missing: %s" % (os.path.basename(path), missing if missing else "NONE"))
with open(r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\check_result.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(res))
print("done")

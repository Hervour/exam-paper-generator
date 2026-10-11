# -*- coding: utf-8 -*-
import os
def norm(s):
    s = s.replace("\u2212", "-").replace("\uff0d", "-").replace("\uff1d", "=").replace("\uff1c", "<").replace("\uff1e", ">")
    return "".join(ch for ch in s if not ch.isspace())
with open(r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\txt\数学 分班考试模拟卷（难度偏高）.txt", encoding="utf-8") as f:
    t = norm(f.read())
keys = ["m=-1或m=3", "2<m<9/4", "0≤m≤2"]
res = []
for k in keys:
    res.append("%s -> %s" % (k, "FOUND" if norm(k) in t else "MISSING"))
with open(r"C:\Users\LHW66\Documents\Codex\2026-08-19\wo\work\check_result2.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(res))
print("done")

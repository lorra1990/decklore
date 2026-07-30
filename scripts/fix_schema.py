#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_schema.py — 修复 pptxgenjs 生成件的 OOXML schema 顺序问题

pptxgenjs 会把 <p:notesMasterIdLst> 写在 <p:sldIdLst> 之后,不合 CT_Presentation
的元素顺序(应在 sldMasterIdLst 之后、sldIdLst 之前)。PowerPoint 容忍,
但严格校验/部分 WPS 版本可能报"需要修复"。本脚本原地重排。

用法: python3 scripts/fix_schema.py build/xxx.pptx
配套: python3 ~/.claude/skills/pptx/scripts/office/validate.py build/xxx.pptx
"""
import re
import shutil
import sys
import zipfile


def fix(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("ppt/presentation.xml").decode("utf-8")

    m = re.search(r"<p:notesMasterIdLst>.*?</p:notesMasterIdLst>", xml, re.S)
    if not m:
        print("无 notesMasterIdLst,不需修复")
        return 0
    seg = m.group(0)
    # 已在 sldIdLst 之前则无需动
    if xml.find(seg) < xml.find("<p:sldIdLst"):
        print("顺序已正确,不需修复")
        return 0

    fixed = xml.replace(seg, "")
    anchor = re.search(r"</p:sldMasterIdLst>", fixed)
    if not anchor:
        print("找不到 sldMasterIdLst 锚点,放弃", file=sys.stderr)
        return 1
    fixed = fixed[: anchor.end()] + seg + fixed[anchor.end():]

    tmp = path + ".fixing"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = fixed.encode("utf-8") if item.filename == "ppt/presentation.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    shutil.move(tmp, path)
    print(f"✓ 已重排 notesMasterIdLst: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(fix(sys.argv[1]))

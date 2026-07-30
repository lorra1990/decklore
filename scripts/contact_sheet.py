#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
contact_sheet.py — 把逐页 slide-*.jpg 拼成一张带页码的九宫格总览图

用途：转图 QA 的第一步先看"全局节奏"——版式是否重复、色彩是否失衡、
     蓝色实底面板是否扎堆、留白节奏是否忽紧忽松。逐页细看放到第二步。

用法:
  python3 scripts/contact_sheet.py build/slide-*.jpg              # 输出 build/sheet.jpg
  python3 scripts/contact_sheet.py build/slide-*.jpg -o out.jpg --cols 4
"""
import argparse
import os
import re
import sys
from PIL import Image, ImageDraw, ImageFont

CELL_W = 480          # 每格宽(px), 高按 16:9 推
PAD = 14              # 格间距
LABEL_H = 30          # 页码条高


def natural_key(p):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", os.path.basename(p))]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("images", nargs="+")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--cols", type=int, default=3)
    args = ap.parse_args()

    paths = sorted([p for p in args.images if os.path.isfile(p)], key=natural_key)
    if not paths:
        print("没有找到图片", file=sys.stderr)
        return 1
    out = args.out or os.path.join(os.path.dirname(paths[0]) or ".", "sheet.jpg")

    cols = args.cols
    rows = (len(paths) + cols - 1) // cols
    cell_h = int(CELL_W * 9 / 16)
    W = cols * CELL_W + (cols + 1) * PAD
    H = rows * (cell_h + LABEL_H) + (rows + 1) * PAD
    sheet = Image.new("RGB", (W, H), (246, 249, 253))
    draw = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 16)
    except Exception:
        font = ImageFont.load_default()

    for i, p in enumerate(paths):
        r, c = divmod(i, cols)
        x = PAD + c * (CELL_W + PAD)
        y = PAD + r * (cell_h + LABEL_H + PAD)
        im = Image.open(p).convert("RGB")
        im.thumbnail((CELL_W, cell_h), Image.LANCZOS)
        # 白底格 + 居中贴图
        draw.rectangle([x, y, x + CELL_W, y + cell_h], fill=(255, 255, 255), outline=(212, 224, 236))
        sheet.paste(im, (x + (CELL_W - im.width) // 2, y + (cell_h - im.height) // 2))
        label = f"P{i + 1}  ·  {os.path.basename(p)}"
        draw.text((x + 4, y + cell_h + 6), label, fill=(90, 106, 130), font=font)

    sheet.save(out, quality=88)
    print(f"✓ {out}  ({len(paths)} 页, {cols} 列)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

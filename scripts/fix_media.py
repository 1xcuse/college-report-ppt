#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修复 pptx 里被导出器写坏的图片字节。

现象：pptx 从第 N 页起图片不显示，ppt/media/image*.png 打开是乱码，
但工程目录 media/ 里的源图是好的（大小与坏块一致）。

做法：按字节大小把 ppt/media 里的坏图映射回工程 media 中的源图，
替换后重新打包。只改媒体字节，不动 XML。

用法：
  python fix_media.py --pptx out.pptx --media-dir <工程>/media [--out out-fixed.pptx]
"""
import argparse, glob, os, shutil, sys, zipfile

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"
JPEG_MAGIC = b"\xff\xd8\xff"


def is_image(data, name):
    if data[:8] == PNG_MAGIC:
        return True
    if data[:3] == JPEG_MAGIC:
        return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--media-dir", required=True)
    ap.add_argument("--out", default="")
    a = ap.parse_args()
    out = a.out or a.pptx.replace(".pptx", "-fixed.pptx")

    by_size = {}
    for f in glob.glob(os.path.join(a.media_dir, "*")):
        if f.lower().endswith((".png", ".jpg", ".jpeg")):
            by_size.setdefault(os.path.getsize(f), f)
    if not by_size:
        sys.exit("media-dir 里没有图片")

    zin = zipfile.ZipFile(a.pptx)
    names = zin.namelist()
    repaired, kept = [], 0
    data_map = {}
    for n in names:
        data = zin.read(n)
        if n.startswith("ppt/media/") and not is_image(data, n):
            src = by_size.get(len(data))
            if src:
                data = open(src, "rb").read()
                repaired.append((n, os.path.basename(src)))
        data_map[n] = data

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(zin.getinfo(n), data_map[n])
    zin.close()
    print(f"repaired {len(repaired)} media parts -> {out}")
    for n, s in repaired:
        print("  ", n, "<-", s)
    if not repaired:
        print("所有媒体字节正常，无需修复")


if __name__ == "__main__":
    main()

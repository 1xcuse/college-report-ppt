#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render_formula.py — 用 matplotlib 把 LaTeX 公式渲染成 PNG，供 PPT 公式页使用。

为什么：PPT 里用文本 run 拼上下标（<sub>/<sup>）会出现「‖·‖₂²」错位、编号不能右对齐的问题。
把公式渲染成透明底 PNG（300 dpi），公式页只放图片 + 右侧编号，看起来才像真 LaTeX 排版。

用法：
  # 1) 单条渲染
  python3 render_formula.py --tex '\\max_{\\mathbf{W}_c} \\mathbb{E}[\\sum_t R(t)]' --out media/f1.png

  # 2) 批量：给 content.json 里所有 formulas[].tex 渲染成图，并回写 img 字段
  python3 render_formula.py --content content.json --media-dir media --update

依赖：matplotlib（DSH 自带 Python 没装时，用系统 /opt/anaconda3/bin/python3）。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

DEFAULT_COLOR = "#1A1A1A"
DEFAULT_FONTSIZE = 20
DEFAULT_DPI = 288  # 4x，显示时按 72dpi 缩放（1pt=1px）


def _mpl():
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        return matplotlib, plt
    except ImportError:
        print("[错误] 没有 matplotlib；请用装了 matplotlib 的解释器，例如 /opt/anaconda3/bin/python3",
              file=sys.stderr)
        return None, None


def render_one(tex: str, out: Path, fontsize: int = DEFAULT_FONTSIZE,
               color: str = DEFAULT_COLOR, dpi: int = DEFAULT_DPI) -> bool:
    matplotlib, plt = _mpl()
    if matplotlib is None:
        return False
    matplotlib.rcParams.update({
        "mathtext.fontset": "stix",
        "font.family": "STIXGeneral",
        "mathtext.default": "regular",
    })
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, "$" + tex.strip().strip("$") + "$", fontsize=fontsize, color=color)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        fig.savefig(out, dpi=dpi, transparent=True, bbox_inches="tight", pad_inches=0.02)
    except Exception as exc:  # noqa: BLE001
        print(f"[错误] 渲染失败：{tex[:60]}（{exc}）", file=sys.stderr)
        return False
    finally:
        plt.close(fig)
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description="LaTeX 公式 → 透明 PNG")
    ap.add_argument("--tex", default="", help="单条公式的 LaTeX 源码（不含 $）")
    ap.add_argument("--out", default="", help="单条模式的输出 PNG")
    ap.add_argument("--content", default="", help="content.json（批量模式）")
    ap.add_argument("--media-dir", default="", help="批量模式的图片输出目录")
    ap.add_argument("--update", action="store_true", help="把 img 字段回写 content.json")
    ap.add_argument("--out-content", default="", help="回写到新文件（缺省原地更新）")
    ap.add_argument("--fontsize", type=int, default=DEFAULT_FONTSIZE)
    ap.add_argument("--color", default=DEFAULT_COLOR)
    ap.add_argument("--dpi", type=int, default=DEFAULT_DPI)
    args = ap.parse_args()

    if args.tex:
        if not args.out:
            ap.error("单条模式需要 --out")
        ok = render_one(args.tex, Path(args.out).expanduser(), args.fontsize, args.color, args.dpi)
        print(("[完成] " if ok else "[失败] ") + str(Path(args.out).expanduser()))
        return 0 if ok else 1

    if not args.content:
        ap.error("需要 --tex/--out 或 --content/--media-dir")
    if not args.media_dir:
        ap.error("批量模式需要 --media-dir")
    content_path = Path(args.content).expanduser()
    media = Path(args.media_dir).expanduser()
    data = json.loads(content_path.read_text(encoding="utf-8"))
    rendered = failed = 0
    for pi, page in enumerate(data.get("pages") or [], 1):
        if not isinstance(page, dict) or page.get("type") != "formula":
            continue
        for fi, f in enumerate(page.get("formulas") or []):
            if not isinstance(f, dict) or not f.get("tex"):
                continue
            name = f.get("img") or f"formula_p{pi:02d}_{fi}.png"
            name = os.path.basename(str(name))
            path = media / name
            if render_one(str(f["tex"]), path, args.fontsize, args.color, args.dpi):
                f["img"] = name
                f.pop("t", None)  # 有图就不再放文本公式
                rendered += 1
            else:
                failed += 1
    if args.update:
        out = Path(args.out_content).expanduser() if args.out_content else content_path
        out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"[完成] 已回写 {out}")
    print(f"[汇总] 渲染成功 {rendered} 条，失败 {failed} 条，输出目录 {media}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())

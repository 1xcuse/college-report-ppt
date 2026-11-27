#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一键独立构建学院汇报 PPT（不依赖 DSH 插件）。

步骤：
  1. build_deck.py      内容 JSON + 档案 → PPTD 工程
  2. pptd_to_pptx.py    PPTD → .pptx（python-pptx 本地轨道）
  3. add_motion.py      给 pptx 补页面转场（本地轨道不写元素动画）

用法：
  python build_report.py --content content.json --profile profile.json \
      --out my-report --pptx 汇报.pptx --media-dir ./figures
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser(description="学院汇报 PPT 一键独立构建（本地轨道）")
    ap.add_argument("--content", required=True, help="content.json")
    ap.add_argument("--profile", default="", help="profile.json（向导输出）")
    ap.add_argument("--out", required=True, help="PPTD 工程目录（不存在会创建）")
    ap.add_argument("--pptx", default="", help="输出 pptx；缺省为 <out>/汇报.pptx")
    ap.add_argument("--crest", default="", help="校徽图片；缺省用模板库/旧版素材")
    ap.add_argument("--logo", default="", help="院徽 / 实验室 logo")
    ap.add_argument("--media-dir", default="", help="论文图目录；图片会拷进工程 media/")
    ap.add_argument("--school", default="", help="学校名（覆盖 profile）")
    ap.add_argument("--college", default="", help="学院名（覆盖 profile）")
    ap.add_argument("--reporter", default="", help="汇报人（覆盖 profile）")
    ap.add_argument("--date", default="", help="封面日期")
    ap.add_argument("--skin", default="", choices=["", "auto", "academic", "campus", "modern", "bold"])
    ap.add_argument("--motion", default="", choices=["", "none", "gentle", "lively"])
    ap.add_argument("--cn-font", default="Microsoft YaHei")
    ap.add_argument("--no-motion", action="store_true", help="不补页面转场")
    args = ap.parse_args()

    skill_dir = Path(__file__).resolve().parent.parent
    scripts = skill_dir / "scripts"
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    py = sys.executable

    cmd = [py, str(scripts / "build_deck.py"), "--content", str(Path(args.content).resolve()),
           "--out", str(out), "--cn-font", args.cn_font]
    for flag, val in (("--profile", args.profile), ("--crest", args.crest), ("--logo", args.logo),
                      ("--media-dir", args.media_dir), ("--school", args.school),
                      ("--college", args.college), ("--reporter", args.reporter),
                      ("--date", args.date), ("--skin", args.skin), ("--motion", args.motion)):
        if val:
            if flag in ("--profile", "--crest", "--logo", "--media-dir"):
                val = str(Path(val).resolve())
            cmd += [flag, str(val)]
    print("[build_report] 生成 PPTD 工程：", " ".join(cmd))
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print("[build_report] build_deck.py 失败；若报缺依赖，请先 pip install -r requirements.txt",
              file=sys.stderr)
        return r.returncode

    deck = out / "deck.pptd"
    if not deck.is_file():
        print(f"[build_report] 没生成 {deck}", file=sys.stderr)
        return 3
    pptx = Path(args.pptx).resolve() if args.pptx else out / "汇报.pptx"
    pptx.parent.mkdir(parents=True, exist_ok=True)
    cmd2 = [py, str(scripts / "pptd_to_pptx.py"), str(deck), "-o", str(pptx), "--force"]
    print("[build_report] 导出 PPTX：", " ".join(cmd2))
    r = subprocess.run(cmd2)
    if r.returncode != 0:
        print("[build_report] 本地轨道导出失败。若文件结构没问题，改用 DSH 的 ppt_export。", file=sys.stderr)
        return r.returncode

    motion = scripts / "add_motion.py"
    if motion.is_file() and not args.no_motion and (args.motion or "") != "none":
        mode = "lively" if (args.motion or "") == "lively" else "gentle"
        tmp = pptx.with_name(pptx.stem + ".motion" + pptx.suffix)
        cmd3 = [py, str(motion), "--pptx", str(pptx), "--out", str(tmp), "--mode",
                "mix" if mode == "lively" else "fade"]
        print("[build_report] 补页面转场：", " ".join(cmd3))
        r = subprocess.run(cmd3)
        if r.returncode != 0:
            print("[build_report] 转场注入失败，保留无转场版本继续。", file=sys.stderr)
        else:
            try:
                tmp.replace(pptx)
            except OSError:
                import shutil
                shutil.move(str(tmp), str(pptx))

    print(f"[build_report] 完成：{pptx}")
    print("[build_report] 提示：本地轨道会丢演讲者备注和元素动画（stderr 有 dropped 清单）；"
          "要完整效果（备注 + 元素动画 + 字体嵌入）请用 DSH 的 ppt_export。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

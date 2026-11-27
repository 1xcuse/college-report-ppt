#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_deck.py — 一键出片：内容 JSON + 档案 → PPTD 工程 / pptx / 讲稿 / 检查报告。

把 build_deck、pptd_to_pptx、add_motion、build_speech、verify_claims、humanize_lint
串成一条命令，适合"内容已经写好，想直接拿到成品"的场景。

用法：
  # 只生成 PPTD 工程（DSH 工具链模式：之后用 ppt_export 导出，效果最完整）
  python3 make_deck.py --content content.json --profile profile.json \
      --out work/deck --media-dir work/figures

  # 同时出 pptx（本地 python-pptx 轨道，会丢备注/元素动画，自动补转场）
  python3 make_deck.py --content content.json --profile profile.json \
      --out work/deck --pptx work/汇报.pptx --media-dir work/figures

  # 连讲稿一起出，并跑回查 + 说人话检查
  python3 make_deck.py --content content.json --profile profile.json \
      --out work/deck --pptx work/汇报.pptx --speech work/讲稿.docx \
      --paper work/paper/paper.md --figures-json work/figures/figures.json \
      --media-dir work/figures --verify --humanize

退出码：0 成功（检查可能有告警），1 有阻断错误（构建/回查失败）。
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import json
from pathlib import Path


def _run(cmd, title):
    print(f"\n=== {title} ===")
    print("$ " + " ".join(str(c) for c in cmd))
    r = subprocess.run([str(c) for c in cmd])
    return r.returncode


def main() -> int:
    ap = argparse.ArgumentParser(description="学院汇报 PPT 一键出片")
    ap.add_argument("--content", required=True, help="content.json")
    ap.add_argument("--profile", default="", help="profile.json（向导输出）")
    ap.add_argument("--out", required=True, help="PPTD 工程目录")
    ap.add_argument("--media-dir", default="", help="论文图目录")
    ap.add_argument("--pptx", default="", help="同时导出 pptx（本地轨道）")
    ap.add_argument("--speech", default="", help="同时生成讲稿 docx")
    ap.add_argument("--speech-md", default="", help="同时生成讲稿 md")
    ap.add_argument("--paper", default="", help="paper_extract 的 paper.md（--verify 用）")
    ap.add_argument("--figures-json", default="", help="extract_figures 的 figures.json（--verify 用）")
    ap.add_argument("--verify", action="store_true", help="跑 verify_claims 回查")
    ap.add_argument("--humanize", action="store_true", help="跑 humanize_lint 说人话检查")
    ap.add_argument("--strict", action="store_true", help="回查/检查有 high 时失败")
    ap.add_argument("--skin", default="", choices=["", "auto", "academic", "campus", "modern", "bold"])
    ap.add_argument("--motion", default="", choices=["", "none", "gentle", "lively"])
    ap.add_argument("--cn-font", default="Microsoft YaHei")
    ap.add_argument("--json", default="", help="把整条流水线结果写一份 json")
    args = ap.parse_args()

    scripts = Path(__file__).resolve().parent
    report = {"steps": [], "ok": True}

    def step(name, cmd):
        rc = _run(cmd, name)
        report["steps"].append({"step": name, "cmd": [str(c) for c in cmd], "rc": rc})
        if rc != 0:
            report["ok"] = False
        return rc

    build = [sys.executable, "-B", scripts / "build_deck.py",
             "--content", Path(args.content).resolve(), "--out", Path(args.out).resolve(),
             "--cn-font", args.cn_font]
    if args.profile:
        build += ["--profile", Path(args.profile).resolve()]
    if args.media_dir:
        build += ["--media-dir", Path(args.media_dir).resolve()]
    if args.skin:
        build += ["--skin", args.skin]
    if args.motion:
        build += ["--motion", args.motion]
    if step("1/5 生成 PPTD 工程", build) != 0:
        report["summary"] = "构建失败"
        _write(args.json, report)
        return 1

    out = Path(args.out).resolve()
    pptx_path = Path(args.pptx).resolve() if args.pptx else None
    if pptx_path:
        conv = [sys.executable, "-B", scripts / "pptd_to_pptx.py",
                str(out / "deck.pptd"), "-o", str(pptx_path), "--force"]
        step("2/5 导出 pptx（本地轨道）", conv)
        motion = scripts / "add_motion.py"
        if motion.is_file() and args.motion != "none":
            tmp = pptx_path.with_name(pptx_path.stem + ".motion" + pptx_path.suffix)
            step("2b/5 补页面转场",
                 [sys.executable, "-B", motion, "--pptx", str(pptx_path), "--out", str(tmp),
                  "--mode", "mix" if args.motion == "lively" else "fade"])
            if tmp.exists():
                tmp.replace(pptx_path)

    if args.speech:
        speech = [sys.executable, "-B", scripts / "build_speech.py",
                  "--content", Path(args.content).resolve(),
                  "--out", Path(args.speech).resolve()]
        if args.profile:
            speech += ["--profile", Path(args.profile).resolve()]
        if pptx_path and pptx_path.is_file():
            speech += ["--pptx", str(pptx_path)]
        if args.speech_md:
            speech += ["--md", Path(args.speech_md).resolve()]
        step("3/5 生成演讲稿 docx", speech)

    if args.verify:
        if not args.paper:
            print("[make_deck][警告] 没给 --paper，跳过 verify_claims")
        else:
            v = [sys.executable, "-B", scripts / "verify_claims.py",
                 "--content", Path(args.content).resolve(), "--paper", Path(args.paper).resolve()]
            if args.figures_json:
                v += ["--figures", Path(args.figures_json).resolve()]
            if args.strict:
                v += ["--strict"]
            step("4/5 内容回查", v)

    if args.humanize:
        h = [sys.executable, "-B", scripts / "humanize_lint.py",
             "--content", Path(args.content).resolve(), "--hints"]
        if args.strict:
            h += ["--strict"]
        step("5/5 说人话检查", h)

    report["summary"] = "完成" if report["ok"] else "有失败步骤"
    report["outputs"] = {
        "deck": str(out),
        "pptx": str(pptx_path) if pptx_path else "",
        "speech": str(Path(args.speech).resolve()) if args.speech else "",
    }
    _write(args.json, report)
    print("\n=== 结果 ===")
    print(report["summary"])
    if report["ok"] and not pptx_path:
        print("提示：这是 DSH 工具链模式。回 DSH 里用 ppt_export({dir: \"%s\", skipCheck: true}) 导出，"
              "可保留备注、字体嵌入和元素动画。" % out)
    return 0 if report["ok"] else 1


def _write(path, obj):
    if not path:
        return
    try:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)
        print(f"流水线报告：{path}")
    except OSError as e:
        print(f"[make_deck][警告] 写报告失败：{e}", file=sys.stderr)


if __name__ == "__main__":
    raise SystemExit(main())

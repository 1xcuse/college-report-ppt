#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""add_motion.py — 给 .pptx 补「页面转场」（轻量兜底，不写元素级动画）。

用途：本地 python-pptx 兜底导出链路（模式 B）没有转场，放映时页与页之间是硬切。
本脚本用 zipfile + lxml 直接改 `ppt/slides/slideN.xml`，在 `</p:sld>` 前插入
`<p:transition spd="…"><p:fade/></p:transition>`（或 push / wipe），其余部件原样保留。

用法：

    python3 add_motion.py --pptx in.pptx --out out.pptx --mode fade|push|wipe|mix [--speed fast|med|slow]
    python3 add_motion.py --pptx in.pptx --out out.pptx --mode mix --speed med

说明：

- 只写「页面转场」（slide transition），不写元素级动画；元素动画走 DSH PPT 工具链
  （ppt_edit 的 motion 字段）。转场是兜底方案里成本最低、观感提升最明显的一层。
- 已存在的 `<p:transition>` 会被替换，不会叠加。
- 按 OOXML 顺序（CT_Slide：cSld → clrMapOvr → transition → timing → extLst）插入；
  有 `<p:timing>` / `<p:extLst>` 时插在它们之前，没有时就是 `</p:sld>` 前。
- 写回前先写临时文件，校验 zip 完整性 + python-pptx 能重新打开 + 页数一致，通过后
  才原子替换到 `--out`；校验失败不会留下半成品。
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
import zipfile
from typing import Dict, List, Optional, Sequence, Tuple

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "缺少依赖：lxml。请用 DSH 自带 Python 运行，或执行 `python3 -m pip install lxml`。\n"
    )
    raise SystemExit(3)

try:
    from pptx import Presentation
except ImportError:  # pragma: no cover
    Presentation = None  # type: ignore[assignment]


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
SLIDE_RE = r"^ppt/slides/slide(\d+)\.xml$"

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# 每种模式：(转场元素名, 方向)；方向 None 表示该转场没有方向属性
TRANSITION_KIND = {
    "fade": ("fade", None),
    "push": ("push", "l"),
    "wipe": ("wipe", "r"),
}

# mix：按页轮换，相邻页观感有变化但不花哨
MIX_PATTERN: List[Tuple[str, Optional[str]]] = [
    ("fade", None),
    ("push", "l"),
    ("wipe", "r"),
    ("fade", None),
    ("push", "u"),
    ("wipe", "l"),
]


def qn(tag: str, ns: str = P_NS) -> str:
    return "{%s}%s" % (ns, tag)


def resolve_transition(mode: str, slide_index: int) -> Tuple[str, Optional[str]]:
    """按模式和第几页（1 起）决定用哪种转场。"""
    if mode == "mix":
        return MIX_PATTERN[(slide_index - 1) % len(MIX_PATTERN)]
    if mode not in TRANSITION_KIND:
        raise ValueError("未知转场模式：%s" % mode)
    return TRANSITION_KIND[mode]


def build_transition(mode: str, speed: str, slide_index: int) -> Tuple[etree._Element, str, Optional[str]]:
    """新建一个 <p:transition> 节点。"""
    kind, direction = resolve_transition(mode, slide_index)
    transition = etree.Element(qn("transition"))
    transition.set("spd", speed)
    child = etree.SubElement(transition, qn(kind))
    if direction:
        child.set("dir", direction)
    return transition, kind, direction


def patch_slide_xml(xml_bytes: bytes, mode: str, speed: str, slide_index: int) -> Tuple[bytes, str, Optional[str]]:
    """替换/插入一页的转场，返回 (新 XML, 转场名, 方向)。"""
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    root = etree.fromstring(xml_bytes, parser=parser)

    # 去掉旧的（含 AlternateContent 里的）转场，避免同一页出现两个
    for transition in root.findall(qn("transition")):
        root.remove(transition)
    for alternate in root.findall("{%s}AlternateContent" % "http://schemas.openxmlformats.org/markup-compatibility/2006"):
        for transition in alternate.iter(qn("transition")):
            parent = transition.getparent()
            if parent is not None:
                parent.remove(transition)

    transition, kind, direction = build_transition(mode, speed, slide_index)

    # OOXML 顺序：transition 必须在 timing / extLst 之前
    insert_at = len(root)
    for tag in ("timing", "extLst"):
        node = root.find(qn(tag))
        if node is not None:
            insert_at = min(insert_at, root.index(node))
    root.insert(insert_at, transition)

    body = etree.tostring(root, xml_declaration=False, encoding="UTF-8")
    return XML_DECL + body, kind, direction


def list_slide_entries(zf: zipfile.ZipFile) -> List[Tuple[int, str]]:
    """返回 [(页号, zip 内路径)]，按页号升序。"""
    import re

    entries: List[Tuple[int, str]] = []
    for name in zf.namelist():
        match = re.match(SLIDE_RE, name)
        if match:
            entries.append((int(match.group(1)), name))
    entries.sort(key=lambda item: item[0])
    return entries


def patch_pptx(pptx_path: str, out_path: str, mode: str, speed: str) -> int:
    if not os.path.exists(pptx_path):
        sys.stderr.write("找不到 pptx：%s\n" % pptx_path)
        return 2
    if os.path.abspath(pptx_path) == os.path.abspath(out_path):
        sys.stderr.write("--out 不能和 --pptx 相同（避免原地改写源文件）\n")
        return 2
    if mode not in ("fade", "push", "wipe", "mix"):
        sys.stderr.write("--mode 只支持 fade / push / wipe / mix\n")
        return 2
    if speed not in ("fast", "med", "slow"):
        sys.stderr.write("--speed 只支持 fast / med / slow\n")
        return 2

    with zipfile.ZipFile(pptx_path, "r") as src:
        bad = src.testzip()
        if bad is not None:
            sys.stderr.write("输入文件 zip 已损坏（%s），请先用 PPT 另存一份。\n" % bad)
            return 2
        entries = list_slide_entries(src)
        if not entries:
            sys.stderr.write("这个 pptx 里没有找到 ppt/slides/slideN.xml，可能不是演示文稿。\n")
            return 2

        infos = {info.filename: info for info in src.infolist()}
        patched: Dict[str, bytes] = {}
        results: List[Tuple[int, str, Optional[str]]] = []
        for slide_no, name in entries:
            raw = src.read(name)
            try:
                new_xml, kind, direction = patch_slide_xml(raw, mode, speed, slide_no)
            except etree.XMLSyntaxError as exc:
                sys.stderr.write("第 %d 页 XML 解析失败：%s\n" % (slide_no, exc))
                return 2
            patched[name] = new_xml
            results.append((slide_no, kind, direction))

        out_dir = os.path.dirname(os.path.abspath(out_path))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        tmp_fd, tmp_path = tempfile.mkstemp(prefix=".add_motion_", suffix=".pptx", dir=out_dir)
        os.close(tmp_fd)
        try:
            with zipfile.ZipFile(tmp_path, "w") as dst:
                for name in src.namelist():
                    info = infos[name]
                    if name in patched:
                        dst.writestr(info, patched[name])
                    else:
                        dst.writestr(info, src.read(name))
        finally:
            src.close()

    # ---- 校验 ----
    ok_zip = False
    try:
        with zipfile.ZipFile(tmp_path, "r") as check:
            ok_zip = check.testzip() is None
    except zipfile.BadZipFile:
        ok_zip = False
    if not ok_zip:
        os.unlink(tmp_path)
        sys.stderr.write("校验失败：写出的 zip 不完整，已丢弃临时文件。\n")
        return 1

    ok_pptx = False
    slide_count = -1
    if Presentation is not None:
        try:
            prs = Presentation(tmp_path)
            slide_count = len(prs.slides)
            ok_pptx = slide_count == len(entries)
        except Exception as exc:  # noqa: BLE001 - 校验失败就报中文
            sys.stderr.write("校验失败：python-pptx 打不开写出的文件（%s）。\n" % exc)
            os.unlink(tmp_path)
            return 1

    if not ok_pptx:
        os.unlink(tmp_path)
        sys.stderr.write("校验失败：python-pptx 能打开但页数对不上，已丢弃临时文件。\n")
        return 1

    os.replace(tmp_path, out_path)

    mode_text = {"fade": "淡出淡入", "push": "推入", "wipe": "擦除", "mix": "混合"}.get(mode, mode)
    print("给 %s 补页面转场" % os.path.basename(pptx_path))
    for slide_no, kind, direction in results:
        extra = "（方向 %s）" % direction if direction else ""
        print("  第 %d 页：%s%s → 已写入" % (slide_no, kind, extra))
    print("模式：%s（%s）  速度：%s" % (mode, mode_text, speed))
    print("已写入：%s" % os.path.abspath(out_path))
    print(
        "校验：zip 完整 ✓  python-pptx 可重新打开 ✓  页数 %d ✓"
        % (slide_count if slide_count >= 0 else len(entries))
    )
    print("提示：本脚本只写页面转场，不写元素级动画；元素动画请在 DSH 里用 ppt_edit 的 motion 字段。")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="add_motion.py",
        description="给 .pptx 补页面转场（fade / push / wipe / mix），轻量兜底，不写元素级动画",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  add_motion.py --pptx deck.pptx --out deck-motion.pptx --mode fade\n"
            "  add_motion.py --pptx deck.pptx --out deck-motion.pptx --mode mix --speed fast\n"
        ),
    )
    parser.add_argument("--pptx", required=True, help="输入 .pptx 路径")
    parser.add_argument("--out", required=True, help="输出 .pptx 路径（不能与输入相同）")
    parser.add_argument("--mode", default="fade", choices=["fade", "push", "wipe", "mix"], help="转场模式")
    parser.add_argument(
        "--speed", default="med", choices=["fast", "med", "slow"], help="转场速度（默认 med）"
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return patch_pptx(args.pptx, args.out, args.mode, args.speed)
    except KeyboardInterrupt:  # pragma: no cover
        sys.stderr.write("\n已取消。\n")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())

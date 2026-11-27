#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cover_art.py — 封面装饰图生成（PIL，确定性、离线）。

给封面/章节页生成一张「不抢文字」的装饰背景：透明底或浅色底，图形集中在右侧，
左侧 55%~60% 留白放标题。同一 `--seed` 重复运行结果完全一致。

用法：

    python3 cover_art.py --style arcs --primary "#00367F" --accent "#0055AA" \
            --size 960x540 --out media/cover_bg.png --seed 7

    python3 cover_art.py --style waves --bg light --out media/cover_bg_light.png
    python3 cover_art.py --style grid --anchor left --intensity subtle --out media/bg.png

    # 打印给 DSH generate_image 用的英文提示词（不生成图片）
    python3 cover_art.py --style arcs --ai-prompt --school 示例大学 \
            --school-en "Northeastern University" --topic "低空 ISAC 通感一体"

风格：
  arcs    右下角同心弧线，安静、学院
  grid    极淡网格 + 少量高亮格，秩序、技术感
  waves   底部叠层正弦波带，流动、现代
  blocks  右侧几何色块，有力量、适合 bold 皮肤

产物可直接当 PPT 背景图（透明 PNG 垫在白底上，或 --bg light 直接铺满）。
"""

from __future__ import annotations

import argparse
import math
import os
import random
import re
import sys
from typing import Callable, List, Optional, Sequence, Tuple

try:
    from PIL import Image, ImageDraw
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "缺少依赖：Pillow。请用 DSH 自带 Python 运行，或执行 `python3 -m pip install Pillow`。\n"
    )
    raise SystemExit(3)


RGB = Tuple[int, int, int]
STYLES = ("arcs", "grid", "waves", "blocks")

NEUTRAL_PRIMARY = "#1F3A5F"
NEUTRAL_ACCENT = "#3A6EA5"

# 英文提示词里每种风格的图形描述
STYLE_PROMPT = {
    "arcs": ("thin concentric arc rings anchored at one corner", "calm, elegant, academic"),
    "grid": ("a very faint architectural grid with a few softly highlighted cells", "orderly, technical, precise"),
    "waves": ("layered soft sine-wave bands along one edge", "flowing, modern, research-poster"),
    "blocks": ("bold flat geometric rectangles and one rotated square in one corner", "confident, modern, poster-like"),
}

SUPERSAMPLE = 3  # 先画 3 倍大再缩回，边缘干净


# ---------------------------------------------------------------------------
# 颜色与绘图小工具
# ---------------------------------------------------------------------------


def hex_to_rgb(value: str) -> RGB:
    if not isinstance(value, str):
        raise ValueError("颜色必须是 #RRGGBB 字符串")
    s = value.strip().lstrip("#")
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    if len(s) != 6:
        raise ValueError("颜色格式不对：%r（要 #RRGGBB）" % value)
    try:
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except ValueError:
        raise ValueError("颜色格式不对：%r（含非十六进制字符）" % value)


def mix(rgb1: Sequence[int], rgb2: Sequence[int], t: float) -> RGB:
    t = max(0.0, min(1.0, t))
    return tuple(int(round(rgb1[i] + (rgb2[i] - rgb1[i]) * t)) for i in range(3))  # type: ignore[return-value]


def darken(rgb: Sequence[int], t: float) -> RGB:
    return mix(rgb, (0, 0, 0), t)


def lighten(rgb: Sequence[int], t: float) -> RGB:
    return mix(rgb, (255, 255, 255), t)


def parse_size(text: str) -> Tuple[int, int]:
    match = re.match(r"^\s*(\d{2,4})\s*[x×X*]\s*(\d{2,4})\s*$", text or "")
    if not match:
        raise ValueError("--size 格式应为 960x540，收到：%r" % text)
    width, height = int(match.group(1)), int(match.group(2))
    if width < 64 or height < 64:
        raise ValueError("--size 太小（宽高至少 64）：%s" % text)
    if width > 4096 or height > 4096:
        raise ValueError("--size 太大（宽高上限 4096）：%s" % text)
    return width, height


def stamp(base: Image.Image, draw_fn: Callable[[ImageDraw.ImageDraw, tuple], None], rgb: RGB, alpha: int) -> Image.Image:
    """把一层半透明图形叠到 base 上（PIL 的 RGBA 直接画会覆盖而不是混合，所以分层合成）。"""
    layer = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw_fn(ImageDraw.Draw(layer), rgb + (max(3, min(255, int(alpha))),))
    return Image.alpha_composite(base, layer)


def scale_alpha(alpha: float, intensity: str) -> int:
    factor = {"subtle": 0.62, "normal": 1.0, "bold": 1.65}.get(intensity, 1.0)
    return int(round(max(3, min(255, alpha * factor))))


# ---------------------------------------------------------------------------
# 四种风格
# ---------------------------------------------------------------------------


def render_arcs(base: Image.Image, size: Tuple[int, int], primary: RGB, accent: RGB,
                rng: random.Random, intensity: str, bg_mode: str) -> Image.Image:
    width, height = size
    main_rgb, second_rgb = (255, 255, 255) if bg_mode == "deep" else primary, accent
    unit = min(width, height)

    # 角落光晕
    def halo(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        r = unit * 0.30
        d.ellipse((width - r, -r, width + r, r), fill=fill)

    base = stamp(base, halo, second_rgb, scale_alpha(16, intensity))

    rings = [
        (unit * 0.52, max(1, int(unit * 0.0030)), main_rgb, 46),
        (unit * 0.70, max(1, int(unit * 0.0024)), second_rgb, 34),
        (unit * 0.88, max(1, int(unit * 0.0020)), main_rgb, 30),
        (unit * 1.10, max(1, int(unit * 0.0016)), second_rgb, 20),
    ]
    for radius, line_w, rgb, alpha in rings:
        radius *= rng.uniform(0.96, 1.04)

        def draw_arc(d: ImageDraw.ImageDraw, fill: tuple, r: float = radius, lw: int = line_w) -> None:
            d.arc((width - r, -r, width + r, r), start=90, end=180, fill=fill, width=lw)

        base = stamp(base, draw_arc, rgb, scale_alpha(alpha, intensity))

    # 沿最大弧线点三颗小点，增加细节但不喧宾
    dot_radius = max(2.0, unit * 0.006)
    for angle_deg in (104.0, 128.0, 152.0):
        theta = math.radians(angle_deg)
        radius = unit * 1.04
        cx = width + radius * math.cos(theta)
        cy = radius * math.sin(theta)  # PIL 角度顺时针、y 向下；圆心在右上角

        def draw_dot(d: ImageDraw.ImageDraw, fill: tuple, x: float = cx, y: float = cy, rr: float = dot_radius) -> None:
            d.ellipse((x - rr, y - rr, x + rr, y + rr), fill=fill)

        base = stamp(base, draw_dot, second_rgb, scale_alpha(92, intensity))

    return base


def render_grid(base: Image.Image, size: Tuple[int, int], primary: RGB, accent: RGB,
                rng: random.Random, intensity: str, bg_mode: str) -> Image.Image:
    width, height = size
    main_rgb = (255, 255, 255) if bg_mode == "deep" else primary
    line_rgb = lighten(primary, 0.15) if bg_mode != "deep" else (255, 255, 255)

    cols = 24
    rows = 14
    cell_w = width / cols
    cell_h = height / rows

    def grid_lines(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        for c in range(cols + 1):
            x = min(width - 1, c * cell_w)
            d.line((x, 0, x, height), fill=fill, width=1)
        for r in range(rows + 1):
            y = min(height - 1, r * cell_h)
            d.line((0, y, width, y), fill=fill, width=1)

    base = stamp(base, grid_lines, line_rgb, scale_alpha(16, intensity))

    # 右侧少量高亮格
    highlights = [
        (cols - 6, 1, 3, 2, primary, 30),
        (cols - 4, 4, 2, 2, accent, 38),
        (cols - 8, 7, 3, 1, primary, 22),
    ]
    for c0, r0, cw, rh, rgb, alpha in highlights:
        if bg_mode == "deep" and rgb == primary:
            rgb = accent
        box = (c0 * cell_w, r0 * cell_h, (c0 + cw) * cell_w, (r0 + rh) * cell_h)

        def draw_cell(d: ImageDraw.ImageDraw, fill: tuple, b=box) -> None:
            d.rectangle(b, fill=fill)

        base = stamp(base, draw_cell, rgb, scale_alpha(alpha, intensity))

    # 右侧竖装饰条
    bar_w = max(3, int(width * 0.008))
    bar_rgb = accent if bg_mode != "deep" else (255, 255, 255)

    def draw_bar(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        d.rectangle((width - bar_w, 0, width, height), fill=fill)

    base = stamp(base, draw_bar, bar_rgb, scale_alpha(46, intensity))
    return base


def render_waves(base: Image.Image, size: Tuple[int, int], primary: RGB, accent: RGB,
                 rng: random.Random, intensity: str, bg_mode: str) -> Image.Image:
    width, height = size
    main_rgb, second_rgb = (255, 255, 255) if bg_mode == "deep" else primary, accent

    # 远处柔和大圆
    def halo(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        r = min(width, height) * 0.42
        d.ellipse((width - r * 1.6, -r * 0.6, width + r * 0.4, r * 1.4), fill=fill)

    base = stamp(base, halo, second_rgb, scale_alpha(14, intensity))

    # (基准高度比, 波幅占画布高, 跨画布周期数, 颜色, alpha, 相位)
    bands = [
        (0.84, 0.018, 0.7, main_rgb, 22, 0.0),
        (0.88, 0.014, 1.1, second_rgb, 28, 1.7),
        (0.93, 0.022, 0.5, main_rgb, 36, 3.1),
    ]
    for base_ratio, amp_ratio, cycles, rgb, alpha, phase in bands:
        amp = height * amp_ratio * rng.uniform(0.85, 1.15)
        base_y = height * base_ratio
        steps = 160
        wave_pts: List[Tuple[float, float]] = []
        for i in range(steps + 1):
            x = width * i / steps
            y = base_y + amp * math.sin(2.0 * math.pi * (x / width * cycles + phase))
            wave_pts.append((x, y))
        wave_pts.append((width, height))
        wave_pts.append((0, height))

        def draw_wave(d: ImageDraw.ImageDraw, fill: tuple, pts=tuple(wave_pts)) -> None:
            d.polygon(pts, fill=fill)

        base = stamp(base, draw_wave, rgb, scale_alpha(alpha, intensity))

    # 波带前沿一条细线
    front_pts: List[Tuple[float, float]] = []
    steps = 160
    for i in range(steps + 1):
        x = width * i / steps
        y = height * 0.93 + height * 0.022 * math.sin(2.0 * math.pi * (x / width * 0.5 + 3.1))
        front_pts.append((x, y))

    def draw_front(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        d.line(front_pts, fill=fill, width=max(1, int(min(width, height) * 0.0022)))

    base = stamp(base, draw_front, second_rgb, scale_alpha(72, intensity))
    return base


def render_blocks(base: Image.Image, size: Tuple[int, int], primary: RGB, accent: RGB,
                  rng: random.Random, intensity: str, bg_mode: str) -> Image.Image:
    width, height = size
    main_rgb = (255, 255, 255) if bg_mode == "deep" else primary
    second_rgb = accent

    # 右上大色块
    def big_block(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        d.rectangle((width * 0.74, 0, width, height * 0.40), fill=fill)

    base = stamp(base, big_block, main_rgb, scale_alpha(22, intensity))

    # 中段窄条
    def mid_block(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        d.rectangle((width * 0.80, height * 0.45, width, height * 0.66), fill=fill)

    base = stamp(base, mid_block, second_rgb, scale_alpha(38, intensity))

    # 旋转 45° 的小方块
    cx = width * 0.79
    cy = height * 0.30
    half = min(width, height) * 0.055 * rng.uniform(0.92, 1.08)

    def diamond(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        d.polygon(
            [(cx - half, cy), (cx, cy - half), (cx + half, cy), (cx, cy + half)],
            fill=fill,
        )

    base = stamp(base, diamond, main_rgb if bg_mode != "deep" else second_rgb, scale_alpha(72, intensity))

    # 右下角一块浅块 + 一条细线
    def corner_block(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        d.rectangle((width * 0.62, height * 0.86, width, height), fill=fill)

    base = stamp(base, corner_block, main_rgb, scale_alpha(18, intensity))

    line_y = height * 0.80

    def thin_line(d: ImageDraw.ImageDraw, fill: tuple) -> None:
        d.rectangle((width * 0.68, line_y, width * 0.92, line_y + max(2, int(height * 0.004))), fill=fill)

    base = stamp(base, thin_line, second_rgb, scale_alpha(80, intensity))
    return base


RENDERERS = {
    "arcs": render_arcs,
    "grid": render_grid,
    "waves": render_waves,
    "blocks": render_blocks,
}


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------


def render_cover_art(
    style: str,
    primary: str,
    accent: str,
    size: Tuple[int, int] = (960, 540),
    seed: int = 7,
    bg: str = "transparent",
    anchor: str = "right",
    intensity: str = "normal",
) -> Image.Image:
    """生成装饰图，返回 RGBA Image。"""
    if style not in RENDERERS:
        raise ValueError("未知风格：%s（可选 %s）" % (style, " / ".join(STYLES)))
    if bg not in ("transparent", "light", "deep"):
        raise ValueError("--bg 只支持 transparent / light / deep")
    if anchor not in ("right", "left"):
        raise ValueError("--anchor 只支持 right / left")
    if intensity not in ("subtle", "normal", "bold"):
        raise ValueError("--intensity 只支持 subtle / normal / bold")

    primary_rgb = hex_to_rgb(primary)
    accent_rgb = hex_to_rgb(accent)
    rng = random.Random(int(seed))

    if bg == "transparent":
        base = Image.new("RGBA", size, (0, 0, 0, 0))
    elif bg == "light":
        base = Image.new("RGBA", size, lighten(primary_rgb, 0.945) + (255,))
    else:  # deep
        base = Image.new("RGBA", size, darken(primary_rgb, 0.25) + (255,))

    base = RENDERERS[style](base, size, primary_rgb, accent_rgb, rng, intensity, bg)

    if anchor == "left":
        base = base.transpose(Image.FLIP_LEFT_RIGHT)

    # 超采样缩回目标尺寸
    if SUPERSAMPLE > 1:
        base = base.resize(size, Image.LANCZOS)
    return base


def build_ai_prompt(args) -> str:
    motif, mood = STYLE_PROMPT.get(args.style or "arcs", STYLE_PROMPT["arcs"])
    school_en = (args.school_en or "").strip()
    school = (args.school or "").strip()
    if school_en and school:
        subject = "%s (%s)" % (school_en, school)
    elif school:
        subject = school
    elif school_en:
        subject = school_en
    else:
        subject = "a Chinese university"

    topic = (args.topic or "").strip()
    bg_desc = {
        "transparent": "a clean very light off-white background",
        "light": "a very light off-white paper background",
        "deep": "a deep dark navy background",
    }[args.bg]

    lines = [
        "Minimal academic presentation cover background for %s." % subject,
    ]
    if topic:
        lines.append("Topic: %s." % topic)
    lines += [
        "Composition: 16:9 landscape; keep the left 55-60% almost empty as negative space "
        "for a large Chinese title and a one-line subtitle; push all decoration to the right edge "
        "and the bottom; low contrast, calm, nothing competing with the text.",
        "Motif: %s, %s." % (motif, mood),
        "Colors: brand primary %s and accent %s on %s; flat areas only, no harsh gradients, "
        "subtle tonal wash allowed." % (args.primary, args.accent, bg_desc),
        "Style: flat vector illustration, precise geometry, professional academic conference aesthetic, "
        "clean edges, print-ready, 4K quality.",
        "No text, no letters, no numbers, no logos, no watermark, no people, no photo, no clutter.",
    ]
    return " ".join(lines)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="cover_art.py",
        description="封面装饰图生成（arcs / grid / waves / blocks），确定性离线，可用 --ai-prompt 打英文提示词",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  cover_art.py --style arcs --primary '#00367F' --accent '#0055AA' --size 960x540 --out media/cover_bg.png --seed 7\n"
            "  cover_art.py --style waves --bg light --out media/cover_light.png\n"
            "  cover_art.py --style grid --ai-prompt --school 示例大学 --topic 低空 ISAC\n"
        ),
    )
    parser.add_argument("--style", choices=list(STYLES), help="风格：arcs / grid / waves / blocks")
    parser.add_argument("--primary", default=NEUTRAL_PRIMARY, help="主色（默认 %s）" % NEUTRAL_PRIMARY)
    parser.add_argument("--accent", default=NEUTRAL_ACCENT, help="辅色（默认 %s）" % NEUTRAL_ACCENT)
    parser.add_argument("--size", default="960x540", help="画布尺寸，默认 960x540（16:9）")
    parser.add_argument("--out", help="输出 PNG 路径")
    parser.add_argument("--seed", type=int, default=7, help="随机种子（默认 7，保证确定性）")
    parser.add_argument(
        "--bg", default="transparent", choices=["transparent", "light", "deep"], help="底色（默认透明）"
    )
    parser.add_argument("--anchor", default="right", choices=["right", "left"], help="装饰在图右侧还是左侧（默认 right）")
    parser.add_argument(
        "--intensity", default="normal", choices=["subtle", "normal", "bold"], help="装饰强度（默认 normal）"
    )
    parser.add_argument("--ai-prompt", action="store_true", help="打印给 generate_image 的英文提示词")
    parser.add_argument("--school", default="", help="学校中文名（--ai-prompt 用）")
    parser.add_argument("--school-en", dest="school_en", default="", help="学校英文名（--ai-prompt 用）")
    parser.add_argument("--topic", default="", help="汇报主题（--ai-prompt 用）")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        size = parse_size(args.size)
    except ValueError as exc:
        sys.stderr.write(str(exc) + "\n")
        return 2

    if args.ai_prompt and not args.out:
        try:
            hex_to_rgb(args.primary)
            hex_to_rgb(args.accent)
        except ValueError as exc:
            sys.stderr.write(str(exc) + "\n")
            return 2
        print(build_ai_prompt(args))
        return 0

    if not args.out:
        sys.stderr.write("缺少 --out（输出 PNG 路径），或使用 --ai-prompt 只打印英文提示词。\n")
        return 2
    if not args.style:
        sys.stderr.write("缺少 --style（arcs / grid / waves / blocks）。\n")
        return 2

    try:
        image = render_cover_art(
            style=args.style,
            primary=args.primary,
            accent=args.accent,
            size=size,
            seed=args.seed,
            bg=args.bg,
            anchor=args.anchor,
            intensity=args.intensity,
        )
    except ValueError as exc:
        sys.stderr.write(str(exc) + "\n")
        return 2

    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    try:
        image.save(args.out, "PNG")
    except OSError as exc:
        sys.stderr.write("保存失败：%s\n" % exc)
        return 2
    if not os.path.exists(args.out) or os.path.getsize(args.out) == 0:
        sys.stderr.write("保存失败：文件为空或未生成 %s\n" % args.out)
        return 2

    bg_text = {"transparent": "透明底", "light": "浅色底", "deep": "深色底"}[args.bg]
    print(
        "已生成封面装饰图：%s（%d×%d，%s，%s，主色 %s / 辅色 %s，seed=%d）"
        % (os.path.abspath(args.out), size[0], size[1], args.style, bg_text, args.primary, args.accent, args.seed)
    )
    safe_side = "右" if args.anchor == "right" else "左"
    print("安全区提示：文字放另一侧（装饰在%s侧，另一侧约 60%% 留白）" % safe_side)

    if args.ai_prompt:
        print()
        print("给 DSH generate_image 的英文提示词：")
        print(build_ai_prompt(args))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

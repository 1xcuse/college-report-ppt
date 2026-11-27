#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""brand_from_logo.py — 校徽取色、文字校徽、品牌包拼装。

三个子命令：

1. 从校徽/院徽图片提取品牌配色（排除近白/近黑/低饱和灰，保证 primary 与白对比度 >= 4.5）：

       python3 brand_from_logo.py palette --image crest.png --json
       → {"primary": "#00367F", "accent": "#0055AA", "deep": "#002B5F", "soft": "#F0F4FA"}

2. 没有校徽图片时，生成干净的「校名文字标识」PNG（品牌色底 + 白字，字号自适应）：

       python3 brand_from_logo.py text-crest --school 示例大学 --primary "#00367F" --out crest.png
       python3 brand_from_logo.py text-crest --school 示例大学 --primary "#00367F" \
               --sub NORTHEASTERN --width 190 --height 44 --out crest-footer.png

3. 按 references/brand_schema.md §2 拼一个品牌包 JSON（供 brand_packs.json 合并）：

       python3 brand_from_logo.py make-pack --school 示例大学 --slug neu --primary "#00367F" \
               --school-en "Northeastern University" --motto "自强不息 知行合一" --out neu.json

依赖：Pillow（DSH 自带 Python 已内置）。字体探测：macOS 用 PingFang / Heiti SC，
Windows 用微软雅黑，Linux 用 Noto Sans CJK / 文泉驿；找不到中文字体时给中文报错。
"""

from __future__ import annotations

import argparse
import colorsys
import json
import math
import os
import sys
from typing import List, Optional, Sequence, Tuple

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:  # pragma: no cover - 环境缺依赖时给中文提示
    sys.stderr.write(
        "缺少依赖：Pillow。请用 DSH 自带 Python 运行，或执行 "
        "`python3 -m pip install Pillow`。\n"
    )
    raise SystemExit(3)


# ---------------------------------------------------------------------------
# 基础颜色工具
# ---------------------------------------------------------------------------

RGB = Tuple[int, int, int]

NEUTRAL_PRIMARY = "#1F3A5F"
NEUTRAL_ACCENT = "#3A6EA5"
INK_DEFAULT = "#1A1A1A"
WHITE: RGB = (255, 255, 255)
BLACK: RGB = (0, 0, 0)


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def hex_to_rgb(value: str) -> RGB:
    """'#00367F' / '00367F' / '#03F' → (r, g, b)。"""
    if not isinstance(value, str):
        raise ValueError("颜色必须是字符串，如 #00367F")
    s = value.strip().lstrip("#")
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    if len(s) != 6:
        raise ValueError("颜色格式不对：%r（要 6 位十六进制，如 #00367F）" % value)
    try:
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except ValueError:
        raise ValueError("颜色格式不对：%r（含非十六进制字符）" % value)


def rgb_to_hex(rgb: Sequence[int]) -> str:
    r, g, b = (int(clamp(round(float(c)), 0, 255)) for c in rgb[:3])
    return "#%02X%02X%02X" % (r, g, b)


def _channel_luminance(c: int) -> float:
    value = c / 255.0
    return value / 12.92 if value <= 0.03928 else ((value + 0.055) / 1.055) ** 2.4


def relative_luminance(rgb: Sequence[int]) -> float:
    r, g, b = rgb[:3]
    return (
        0.2126 * _channel_luminance(int(r))
        + 0.7152 * _channel_luminance(int(g))
        + 0.0722 * _channel_luminance(int(b))
    )


def contrast_ratio(rgb1: Sequence[int], rgb2: Sequence[int]) -> float:
    """WCAG 对比度，返回 1.0 ~ 21.0。"""
    l1 = relative_luminance(rgb1)
    l2 = relative_luminance(rgb2)
    high, low = (l1, l2) if l1 >= l2 else (l2, l1)
    return (high + 0.05) / (low + 0.05)


def mix_rgb(rgb1: Sequence[int], rgb2: Sequence[int], t: float) -> RGB:
    """线性混合：t=0 取 rgb1，t=1 取 rgb2。"""
    t = clamp(t, 0.0, 1.0)
    return (
        int(round(rgb1[0] + (rgb2[0] - rgb1[0]) * t)),
        int(round(rgb1[1] + (rgb2[1] - rgb1[1]) * t)),
        int(round(rgb1[2] + (rgb2[2] - rgb1[2]) * t)),
    )


def darken_rgb(rgb: Sequence[int], t: float = 0.25) -> RGB:
    return mix_rgb(rgb, BLACK, t)


def lighten_rgb(rgb: Sequence[int], t: float = 0.25) -> RGB:
    return mix_rgb(rgb, WHITE, t)


def ensure_white_contrast(rgb: Sequence[int], min_ratio: float = 4.5) -> RGB:
    """把颜色逐步压暗，直到与白色的对比度 >= min_ratio。"""
    current: RGB = (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    if contrast_ratio(current, WHITE) >= min_ratio:
        return current
    for _ in range(80):
        current = darken_rgb(current, 0.08)
        if contrast_ratio(current, WHITE) >= min_ratio:
            return current
        if current == (0, 0, 0):
            break
    return current


def _rgb_distance(a: Sequence[int], b: Sequence[int]) -> float:
    return math.sqrt(sum((float(a[i]) - float(b[i])) ** 2 for i in range(3)))


def _hue_distance(h1: float, h2: float) -> float:
    d = abs(h1 - h2) % 1.0
    return min(d, 1.0 - d)


def hsv(rgb: Sequence[int]) -> Tuple[float, float, float]:
    return colorsys.rgb_to_hsv(rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)


# ---------------------------------------------------------------------------
# 取色
# ---------------------------------------------------------------------------

# 过滤阈值：太亮（白底/反白）太暗（黑边/黑字）不饱和（灰）都不算品牌色
_V_MAX = 0.96
_V_MIN = 0.07
_S_MIN = 0.14


def _open_image(source) -> Image.Image:
    """接受路径、bytes、PIL.Image。返回 RGB/RGBA 图像。"""
    if isinstance(source, Image.Image):
        return source
    if isinstance(source, (bytes, bytearray)):
        from io import BytesIO

        return Image.open(BytesIO(bytes(source)))
    if isinstance(source, str):
        if not os.path.exists(source):
            raise FileNotFoundError("找不到图片：%s" % source)
        return Image.open(source)
    raise TypeError("不支持的图片来源：%r" % type(source))


def _candidate_colors(img: Image.Image) -> List[Tuple[float, RGB, float, float, float]]:
    """量化后返回候选色 [(权重, rgb, h, s, v)]，按权重降序。"""
    img = img.convert("RGBA")
    if max(img.size) > 240:
        img.thumbnail((240, 240), Image.LANCZOS)
    if min(img.size) < 8:
        return []

    alpha = img.getchannel("A")
    # 用直方图数不透明像素，避免 getdata() 的弃用告警
    opaque = sum(count for value, count in enumerate(alpha.histogram()) if value > 128)
    if opaque == 0:
        return []

    flat = Image.alpha_composite(Image.new("RGBA", img.size, (255, 255, 255, 255)), img).convert("RGB")
    try:
        quantize_method = Image.Quantize.MEDIANCUT
    except AttributeError:  # Pillow < 9.1
        quantize_method = 0  # type: ignore[assignment]
    quant = flat.quantize(colors=64, method=quantize_method)
    palette = quant.getpalette() or []
    counts = quant.getcolors() or []

    candidates = []
    for count, index in counts:
        base = index * 3
        if base + 2 >= len(palette):
            continue
        rgb = (palette[base], palette[base + 1], palette[base + 2])
        h, s, v = hsv(rgb)
        if v > _V_MAX or v < _V_MIN or s < _S_MIN:
            continue
        weight = float(count) * (0.30 + 0.70 * s) * (0.55 + 0.45 * (1.0 - v))
        candidates.append((weight, rgb, h, s, v))

    candidates.sort(key=lambda item: item[0], reverse=True)

    deduped: List[Tuple[float, RGB, float, float, float]] = []
    for item in candidates:
        rgb = item[1]
        if any(_rgb_distance(rgb, kept[1]) < 34 for kept in deduped):
            continue
        deduped.append(item)
    return deduped


def _pick_accent(
    candidates: List[Tuple[float, RGB, float, float, float]],
    primary: RGB,
    primary_hue: float,
) -> RGB:
    """挑一个和主色区分度够、能当点缀的辅色；挑不到就由主色提亮。"""
    for weight, rgb, hue, sat, val in candidates:
        if _rgb_distance(rgb, primary) < 52:
            continue
        hue_gap = _hue_distance(hue, primary_hue)
        # 同色系要明显更亮，异色系要色相差 >= 20°
        if hue_gap < 20.0 / 360.0 and val - hsv(primary)[2] < 0.18:
            continue
        if contrast_ratio(rgb, WHITE) < 3.0:
            continue
        return rgb

    # 兜底：主色提亮一档，控制在对比度 >= 3.0
    for step in (0.28, 0.24, 0.20, 0.16, 0.12, 0.10):
        derived = lighten_rgb(primary, step)
        if contrast_ratio(derived, WHITE) >= 3.0:
            return derived
    return lighten_rgb(primary, 0.10)


def extract_palette(source, min_contrast: float = 4.5, verbose: bool = False) -> dict:
    """从图片提取品牌色。

    返回 {"primary","accent","deep","soft"}，全部大写十六进制；
    primary 与白色的对比度保证 >= min_contrast（默认 4.5）。
    """
    img = _open_image(source)
    candidates = _candidate_colors(img)

    if candidates:
        primary = candidates[0][1]
        primary_hue = candidates[0][2]
        accent = _pick_accent(candidates, primary, primary_hue)
        if verbose:
            sys.stderr.write("候选色（权重降序，最多 8 个）：\n")
            for weight, rgb, h, s, v in candidates[:8]:
                sys.stderr.write(
                    "  %s  h=%.2f s=%.2f v=%.2f weight=%.1f\n"
                    % (rgb_to_hex(rgb), h, s, v, weight)
                )
    else:
        primary = hex_to_rgb(NEUTRAL_PRIMARY)
        accent = hex_to_rgb(NEUTRAL_ACCENT)
        if verbose:
            sys.stderr.write("图片里没有可用品牌色（全白/全黑/近灰），改用中性色。\n")

    primary = ensure_white_contrast(primary, min_contrast)
    # accent 只做点缀：能看清即可（>= 2.5），不强行压到 4.5
    if contrast_ratio(accent, WHITE) < 2.5:
        accent = ensure_white_contrast(accent, 2.5)

    deep = darken_rgb(primary, 0.25)
    soft = lighten_rgb(primary, 0.92)

    return {
        "primary": rgb_to_hex(primary),
        "accent": rgb_to_hex(accent),
        "deep": rgb_to_hex(deep),
        "soft": rgb_to_hex(soft),
    }


# ---------------------------------------------------------------------------
# 字体探测（中文）
# ---------------------------------------------------------------------------

# (路径, 首选 index, 备注)。ttc 里 index 可能变，_open_font 会再扫一遍挑 SC/中文族。
FONT_CANDIDATES = [
    # macOS
    ("/System/Library/Fonts/PingFang.ttc", 0, "PingFang SC"),
    ("/System/Library/Fonts/STHeiti Medium.ttc", 1, "Heiti SC Medium"),
    ("/System/Library/Fonts/Hiragino Sans GB.ttc", 2, "Hiragino Sans GB W6"),
    ("/Library/Fonts/Arial Unicode.ttf", 0, "Arial Unicode MS"),
    ("/System/Library/Fonts/Supplemental/Arial Unicode.ttf", 0, "Arial Unicode MS"),
    # Windows
    ("C:/Windows/Fonts/msyhbd.ttc", 0, "Microsoft YaHei Bold"),
    ("C:/Windows/Fonts/msyh.ttc", 0, "Microsoft YaHei"),
    ("C:/Windows/Fonts/simhei.ttf", 0, "SimHei"),
    ("C:/Windows/Fonts/simsun.ttc", 0, "SimSun"),
    # Linux
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 0, "Noto Sans CJK"),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0, "Noto Sans CJK"),
    ("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc", 0, "Noto Sans CJK"),
    ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0, "WenQuanYi Zen Hei"),
    ("/usr/share/fonts/wenquanyi/wqy-zenhei/wqy-zenhei.ttc", 0, "WenQuanYi Zen Hei"),
]


def _font_works(path: str, index: int, size: int = 40) -> Optional[ImageFont.FreeTypeFont]:
    try:
        return ImageFont.truetype(path, size, index=index)
    except (OSError, ValueError):
        return None


def find_cjk_font() -> Tuple[str, int, str]:
    """返回 (字体路径, ttc index, 字体族名)。找不到时抛 RuntimeError（中文）。"""
    # 1. 先按候选顺序找「族名像中文黑体」的字重；对 ttc 额外扫 index
    for path, index, label in FONT_CANDIDATES:
        if not os.path.exists(path):
            continue
        font = _font_works(path, index)
        if font is not None:
            try:
                family = " ".join(font.getname())
            except Exception:  # pragma: no cover
                family = label
            low = family.lower()
            if "sc" in low or "cjk" in low or "heiti" in low or "hei" in low:
                return path, index, family

        # PingFang.ttc 里 SC 往往不是 0 号，扫一遍其他 index
        for idx in range(0, 18):
            if idx == index:
                continue
            probe = _font_works(path, idx)
            if probe is None:
                continue
            try:
                fam = " ".join(probe.getname())
            except Exception:  # pragma: no cover
                continue
            low = fam.lower()
            if ("sc" in low or "cjk" in low or "heiti" in low) and (
                "bold" in low or "medium" in low or "semibold" in low or "regular" in low or "w6" in low or "w5" in low
            ):
                return path, idx, fam

    # 2. 兜底：任何能开的中文字体候选
    for path, index, label in FONT_CANDIDATES:
        if not os.path.exists(path):
            continue
        font = _font_works(path, index)
        if font is not None:
            try:
                return path, index, " ".join(font.getname())
            except Exception:  # pragma: no cover
                return path, index, label
    raise RuntimeError(
        "没有找到可用的中文字体。\n"
        "  macOS：确认 /System/Library/Fonts/STHeiti Medium.ttc 存在\n"
        "  Windows：确认 C:/Windows/Fonts/msyhbd.ttc 存在\n"
        "  Linux：安装 fonts-noto-cjk 或文泉驿正黑\n"
        "  也可以把中文字体文件放到 /Library/Fonts 后重试。"
    )


def _load_font(path: str, index: int, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, max(6, int(size)), index=index)


def _text_bbox(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont):
    return draw.textbbox((0, 0), text, font=font)


def _fit_font_size(
    draw: ImageDraw.ImageDraw,
    text: str,
    path: str,
    index: int,
    max_width: int,
    target_size: int,
    min_size: int = 8,
) -> Tuple[ImageFont.FreeTypeFont, int, int]:
    """把字号从 target_size 往下收，直到文字宽度 <= max_width。返回 (font, size, text_width)。"""
    size = max(min_size, int(target_size))
    while size > min_size:
        font = _load_font(path, index, size)
        bbox = _text_bbox(draw, text, font)
        width = bbox[2] - bbox[0]
        if width <= max_width:
            return font, size, width
        size -= 1
    font = _load_font(path, index, min_size)
    bbox = _text_bbox(draw, text, font)
    return font, min_size, bbox[2] - bbox[0]


def _draw_centered(
    draw: ImageDraw.ImageDraw, center_x: float, center_y: float, text: str, font, fill
) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    width = bbox[2] - bbox[0]
    height = bbox[3] - bbox[1]
    x = center_x - width / 2.0 - bbox[0]
    y = center_y - height / 2.0 - bbox[1]
    draw.text((x, y), text, font=font, fill=fill)


def make_text_crest(
    school: str,
    primary: str,
    out_path: str,
    sub: str = "",
    width: int = 500,
    height: int = 110,
    bg_color: Optional[str] = None,
    text_color: str = "#FFFFFF",
) -> dict:
    """生成文字校徽 PNG（品牌色底 + 白字，字号自适应）。

    返回 {"path","width","height","font","primary","contrast"}。
    """
    if not school or not school.strip():
        raise ValueError("--school 不能为空")
    if width < 60 or height < 24:
        raise ValueError("尺寸太小：width 至少 60、height 至少 24")
    if width > 2000 or height > 2000:
        raise ValueError("尺寸太大：width/height 上限 2000")

    primary_rgb = hex_to_rgb(primary)
    bg_rgb = hex_to_rgb(bg_color) if bg_color else ensure_white_contrast(primary_rgb, 4.5)
    text_rgb = hex_to_rgb(text_color)

    font_path, font_index, family = find_cjk_font()
    scale = 2  # 2 倍超采样，缩回后边缘更干净
    canvas_w, canvas_h = width * scale, height * scale
    img = Image.new("RGB", (canvas_w, canvas_h), bg_rgb)
    draw = ImageDraw.Draw(img)

    pad_x = max(6, int(canvas_w * 0.07))
    max_text_w = canvas_w - 2 * pad_x
    school = school.strip()
    sub = (sub or "").strip()

    if sub:
        main_target = int(canvas_h * 0.46)
        sub_target = int(canvas_h * 0.22)
        main_font, main_size, main_w = _fit_font_size(
            draw, school, font_path, font_index, max_text_w, main_target, min_size=10
        )
        sub_font, sub_size, sub_w = _fit_font_size(
            draw, sub, font_path, font_index, max_text_w, sub_target, min_size=8
        )
        main_bb = _text_bbox(draw, school, main_font)
        sub_bb = _text_bbox(draw, sub, sub_font)
        main_h = main_bb[3] - main_bb[1]
        sub_h = sub_bb[3] - sub_bb[1]
        gap = max(2, int(canvas_h * 0.05))
        total_h = main_h + gap + sub_h
        main_cy = canvas_h / 2.0 - total_h / 2.0 + main_h / 2.0
        sub_cy = canvas_h / 2.0 + total_h / 2.0 - sub_h / 2.0
        _draw_centered(draw, canvas_w / 2.0, main_cy, school, main_font, text_rgb)
        _draw_centered(draw, canvas_w / 2.0, sub_cy, sub.upper() if sub.isascii() else sub, sub_font, text_rgb)
    else:
        main_target = int(canvas_h * 0.50)
        main_font, main_size, main_w = _fit_font_size(
            draw, school, font_path, font_index, max_text_w, main_target, min_size=10
        )
        _draw_centered(draw, canvas_w / 2.0, canvas_h / 2.0, school, main_font, text_rgb)

    img = img.resize((width, height), Image.LANCZOS)
    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    img.save(out_path, "PNG")
    if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
        raise RuntimeError("写文件失败：%s" % out_path)

    return {
        "path": os.path.abspath(out_path),
        "width": width,
        "height": height,
        "font": family,
        "primary": rgb_to_hex(bg_rgb),
        "text": rgb_to_hex(text_rgb),
        "contrast": round(contrast_ratio(bg_rgb, text_rgb), 2),
    }


# ---------------------------------------------------------------------------
# 品牌包
# ---------------------------------------------------------------------------


def build_brand_pack(args) -> dict:
    """按 brand_schema.md §2 拼品牌包 dict。"""
    colors: dict = {}
    source_note_parts: List[str] = []

    if args.from_image:
        palette = extract_palette(args.from_image)
        colors.update(palette)
        source_note_parts.append("配色由校徽取色")

    if args.primary:
        colors["primary"] = rgb_to_hex(hex_to_rgb(args.primary))
    if args.accent:
        colors["accent"] = rgb_to_hex(hex_to_rgb(args.accent))
    if args.deep:
        colors["deep"] = rgb_to_hex(hex_to_rgb(args.deep))
    if args.soft:
        colors["soft"] = rgb_to_hex(hex_to_rgb(args.soft))
    if args.ink:
        colors["ink"] = rgb_to_hex(hex_to_rgb(args.ink))
    if args.gold:
        colors["gold"] = rgb_to_hex(hex_to_rgb(args.gold))

    if "primary" not in colors:
        raise ValueError("缺少 primary：用 --primary #xxxxxx，或 --from-image 从校徽取色")
    primary_rgb = hex_to_rgb(colors["primary"])
    colors["primary"] = rgb_to_hex(ensure_white_contrast(primary_rgb, 4.5))
    primary_rgb = hex_to_rgb(colors["primary"])
    colors.setdefault("accent", rgb_to_hex(lighten_rgb(primary_rgb, 0.24)))
    colors.setdefault("deep", rgb_to_hex(darken_rgb(primary_rgb, 0.25)))
    colors.setdefault("soft", rgb_to_hex(lighten_rgb(primary_rgb, 0.92)))
    colors.setdefault("ink", INK_DEFAULT)
    # 字段顺序稳定：primary/accent/deep/soft/ink/gold
    ordered_colors = {key: colors[key] for key in ("primary", "accent", "deep", "soft", "ink", "gold") if key in colors}

    pack: dict = {
        "school": args.school,
        "school_en": args.school_en or "",
        "colors": ordered_colors,
    }

    crest_kind = args.crest_kind or ("image" if args.crest else "text")
    if args.crest:
        pack["crest"] = args.crest
    else:
        pack["crest"] = None
    pack["crest_kind"] = crest_kind
    pack["crest_source"] = args.crest_source or ("estimated" if not args.crest else "official-website")
    pack["crest_verified"] = bool(args.crest_verified) and bool(args.crest and args.crest_verified)

    if args.motto:
        pack["motto"] = args.motto
    if args.footer_left:
        pack["footer_left"] = args.footer_left
    elif args.motto:
        pack["footer_left"] = args.motto
    if args.footer_right:
        pack["footer_right"] = args.footer_right
    elif args.school_en:
        pack["footer_right"] = args.school_en

    pack["decor"] = args.decor or "rule"
    pack["skin_hint"] = args.skin_hint or "academic"
    if args.source:
        pack["source"] = args.source
    elif source_note_parts:
        pack["source"] = "；".join(source_note_parts)
    pack["verified"] = bool(args.verified)
    return pack


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def cmd_palette(args) -> int:
    try:
        palette = extract_palette(args.image, min_contrast=args.min_contrast, verbose=args.verbose)
    except FileNotFoundError as exc:
        sys.stderr.write(str(exc) + "\n")
        return 2
    except Exception as exc:  # 图片打不开、不是图片等
        sys.stderr.write("取色失败：%s\n" % exc)
        return 2

    if args.out:
        out_dir = os.path.dirname(os.path.abspath(args.out))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(palette, fh, ensure_ascii=False, indent=2)
            fh.write("\n")

    if args.json:
        print(json.dumps(palette, ensure_ascii=False))
    else:
        primary_rgb = hex_to_rgb(palette["primary"])
        print("图片：%s" % args.image)
        print("主色 primary：%s（与白对比度 %.2f:1）" % (palette["primary"], contrast_ratio(primary_rgb, WHITE)))
        print("辅色 accent：%s" % palette["accent"])
        print("深色 deep：%s" % palette["deep"])
        print("浅色 soft：%s" % palette["soft"])
        if args.out:
            print("已写入：%s" % os.path.abspath(args.out))
    return 0


def cmd_text_crest(args) -> int:
    try:
        info = make_text_crest(
            school=args.school,
            primary=args.primary,
            out_path=args.out,
            sub=args.sub or "",
            width=args.width,
            height=args.height,
            bg_color=args.bg_color,
            text_color=args.text_color,
        )
    except (ValueError, RuntimeError, OSError) as exc:
        sys.stderr.write("生成文字校徽失败：%s\n" % exc)
        return 2
    print(
        "已生成文字校徽：%s（%d×%d，底色 %s，白字对比度 %.2f:1）"
        % (info["path"], info["width"], info["height"], info["primary"], info["contrast"])
    )
    print("字体：%s" % info["font"])
    if args.json:
        print(json.dumps(info, ensure_ascii=False))
    return 0


def cmd_make_pack(args) -> int:
    try:
        pack = build_brand_pack(args)
    except (ValueError, FileNotFoundError) as exc:
        sys.stderr.write("拼品牌包失败：%s\n" % exc)
        return 2
    text = json.dumps(pack, ensure_ascii=False, indent=2)
    if args.out:
        out_dir = os.path.dirname(os.path.abspath(args.out))
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
        print("已写入品牌包：%s" % os.path.abspath(args.out))
        print("slug 建议键：%s" % (args.slug or "(未提供 --slug，合并时自行指定)"))
    print(text)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="brand_from_logo.py",
        description="校徽取色 / 文字校徽 / 品牌包拼装（本地离线，Pillow）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  brand_from_logo.py palette --image assets/schools/crests/neu.png --json\n"
            "  brand_from_logo.py text-crest --school 示例大学 --primary '#00367F' --out crest.png\n"
            "  brand_from_logo.py make-pack --school 示例大学 --slug neu --primary '#00367F' --out neu.json\n"
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_palette = sub.add_parser("palette", help="从图片提取品牌配色")
    p_palette.add_argument("--image", required=True, help="校徽/院徽图片路径（PNG/JPG）")
    p_palette.add_argument("--json", action="store_true", help="只输出一行 JSON")
    p_palette.add_argument("--out", help="把配色写到 JSON 文件")
    p_palette.add_argument("--min-contrast", type=float, default=4.5, help="primary 与白的对比度下限（默认 4.5）")
    p_palette.add_argument("--verbose", action="store_true", help="把候选色打印到 stderr")
    p_palette.set_defaults(func=cmd_palette)

    p_crest = sub.add_parser("text-crest", help="生成文字校徽 PNG")
    p_crest.add_argument("--school", required=True, help="校名（中文），如 示例大学")
    p_crest.add_argument("--primary", required=True, help="品牌主色，如 #00367F")
    p_crest.add_argument("--out", required=True, help="输出 PNG 路径")
    p_crest.add_argument("--sub", default="", help="副标题（英文校名/简称），可选")
    p_crest.add_argument("--width", type=int, default=500, help="输出宽（默认 500；封面小图可用 190）")
    p_crest.add_argument("--height", type=int, default=110, help="输出高（默认 110；封面小图可用 44）")
    p_crest.add_argument("--bg-color", help="底色覆盖（默认品牌色）")
    p_crest.add_argument("--text-color", default="#FFFFFF", help="文字颜色（默认 #FFFFFF）")
    p_crest.add_argument("--json", action="store_true", help="额外输出 JSON 摘要")
    p_crest.set_defaults(func=cmd_text_crest)

    p_pack = sub.add_parser("make-pack", help="拼一个品牌包 JSON")
    p_pack.add_argument("--school", required=True, help="学校中文全称")
    p_pack.add_argument("--slug", help="学校 slug（只用于提示，不写进包）")
    p_pack.add_argument("--school-en", default="", help="英文校名")
    p_pack.add_argument("--primary", help="主色 #RRGGBB")
    p_pack.add_argument("--accent", help="辅色")
    p_pack.add_argument("--deep", help="深色（缺省 primary 压暗 25%）")
    p_pack.add_argument("--soft", help="浅色（缺省 primary 提亮 92%）")
    p_pack.add_argument("--ink", help="正文色（缺省 #1A1A1A）")
    p_pack.add_argument("--gold", help="点缀金色，可选")
    p_pack.add_argument("--from-image", help="从校徽图片取色（--primary 仍可覆盖）")
    p_pack.add_argument("--crest", help="校徽相对 assets/schools/ 的路径，如 crests/neu.png")
    p_pack.add_argument("--crest-kind", choices=["image", "text"], help="校徽类型")
    p_pack.add_argument(
        "--crest-source",
        choices=["official-website", "wikimedia", "dataset", "user", "estimated"],
        help="校徽来源",
    )
    p_pack.add_argument("--crest-verified", action="store_true", help="人工看过校徽（像素/透明底/无白边）")
    p_pack.add_argument("--motto", default="", help="校训")
    p_pack.add_argument("--footer-left", default="", help="页脚左文字（缺省用 motto）")
    p_pack.add_argument("--footer-right", default="", help="页脚右文字（缺省用英文校名）")
    p_pack.add_argument("--decor", choices=["rule", "block", "arc", "minimal"], help="装饰风格（默认 rule）")
    p_pack.add_argument(
        "--skin-hint", choices=["academic", "campus", "modern", "bold"], help="推荐皮肤（默认 academic）"
    )
    p_pack.add_argument("--source", default="", help="数据来源说明")
    p_pack.add_argument("--verified", action="store_true", help="配色经人工核对（VI 标准色）")
    p_pack.add_argument("--out", help="写入 JSON 文件")
    p_pack.set_defaults(func=cmd_make_pack)
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args))
    except KeyboardInterrupt:  # pragma: no cover
        sys.stderr.write("\n已取消。\n")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())

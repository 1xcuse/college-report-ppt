#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""学院汇报 PPT 生成器 v2（全国本科院校通用）。

content.json（+ profile.json）→ PPTD 工程（deck.pptd + pages/*.page + media/），
再用 DSH 的 ppt_export 导出 .pptx；没有 DSH 时用 build_report.py 走本地 python-pptx 轨道。

用法：
  python build_deck.py --content content.json --out 工程目录 \
      [--profile profile.json] [--school 示例大学] [--college 示例学院] \
      [--crest 校徽.png] [--logo 院徽.png] [--media-dir 图片目录] \
      [--skin auto|academic|campus|modern|bold] [--motion none|gentle|lively] \
      [--reporter 张三] [--date "2026 年 11 月"] [--cn-font "Microsoft YaHei"]

其它：
  --list-schools [关键字]   查模板库里的学校（需要 assets/schools/schools.json）
  --school-info 学校名      打印匹配到的学校与品牌包
  --check-only              只做内容检查，不写工程
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import sys
from datetime import date

try:
    import yaml
except ImportError:  # pragma: no cover
    sys.exit("缺少 PyYAML：请先 pip install -r requirements.txt（DSH 自带 Python 已内置）")

# ------------------------------------------------------------------ 常量
W, H = 960, 540
MARGIN = 56
RIGHT = W - MARGIN  # 904
SZ = {"title": 27, "sub": 18, "body": 18.5, "note": 13.5, "cap": 11.5, "foot": 9.5}

DEFAULT_PRIMARY = "#1F3A5F"
DEFAULT_ACCENT = "#2E5C8A"

PURPOSE_DEFAULTS = {
    "course_report": {"skin": "campus", "duration": 10, "pages": "10–14"},
    "group_paper": {"skin": "academic", "duration": 20, "pages": "15–20"},
    "group_progress": {"skin": "modern", "duration": 12, "pages": "8–12"},
    "proposal": {"skin": "academic", "duration": 18, "pages": "12–18"},
    "defense": {"skin": "academic", "duration": 25, "pages": "18–25"},
    "project": {"skin": "bold", "duration": 15, "pages": "12–18"},
}

SKIN_NAMES = ("academic", "campus", "modern", "bold")

# ------------------------------------------------------------------ 颜色工具
def _rgb(c):
    c = str(c).strip().lstrip("#")
    if len(c) == 3:
        c = "".join(ch * 2 for ch in c)
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _hex(t):
    return "#%02X%02X%02X" % tuple(max(0, min(255, int(round(v)))) for v in t)


def mix(c1, c2, t):
    a, b = _rgb(c1), _rgb(c2)
    return _hex(tuple(a[i] * (1 - t) + b[i] * t for i in range(3)))


def darken(c, t=0.25):
    return mix(c, "#000000", t)


def lighten(c, t=0.92):
    return mix(c, "#FFFFFF", t)


def _lin(v):
    v = v / 255.0
    return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4


def contrast(c1, c2):
    l1 = sum(k * _lin(v) for k, v in zip((0.2126, 0.7152, 0.0722), _rgb(c1)))
    l2 = sum(k * _lin(v) for k, v in zip((0.2126, 0.7152, 0.0722), _rgb(c2)))
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def ensure_contrast(c, bg="#FFFFFF", ratio=4.5):
    out = c
    for _ in range(12):
        if contrast(out, bg) >= ratio:
            return out
        out = darken(out, 0.07)
    return out


# ------------------------------------------------------------------ 小工具
def est_lines(text, width, size):
    """粗估中文排版行数：中文按 1 字宽、西文按 0.55 字宽。"""
    unit = sum(1.0 if ord(ch) > 127 else 0.55 for ch in str(text))
    per = max(1.0, float(width) / max(1.0, float(size)))
    return max(1, int(math.ceil(unit / per)))


def read_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def write_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


class Warnings:
    def __init__(self):
        self.items = []

    def warn(self, msg):
        self.items.append(("WARN", msg))
        print("[build_deck][警告] " + msg, file=sys.stderr)

    def error(self, msg):
        self.items.append(("ERROR", msg))
        print("[build_deck][错误] " + msg, file=sys.stderr)


WARNINGS = Warnings()

# ------------------------------------------------------------------ 学校模板库
SCHOOL_CACHE = {}


def _skill_dir():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_school_db(skill_dir):
    if "db" in SCHOOL_CACHE:
        return SCHOOL_CACHE["db"]
    path = os.path.join(skill_dir, "assets", "schools", "schools.json")
    db = []
    if os.path.isfile(path):
        try:
            data = read_json(path)
            if isinstance(data, dict):
                data = data.get("schools", [])
            if isinstance(data, list):
                db = [x for x in data if isinstance(x, dict) and x.get("name")]
        except Exception as e:  # noqa: BLE001
            WARNINGS.warn(f"学校名单读取失败：{path}（{e}）")
    SCHOOL_CACHE["db"] = db
    return db


def load_brand_packs(skill_dir):
    if "brands" in SCHOOL_CACHE:
        return SCHOOL_CACHE["brands"]
    path = os.path.join(skill_dir, "assets", "schools", "brand_packs.json")
    packs = {}
    if os.path.isfile(path):
        try:
            data = read_json(path)
            if isinstance(data, dict):
                packs = data
        except Exception as e:  # noqa: BLE001
            WARNINGS.warn(f"品牌包读取失败：{path}（{e}）")
    SCHOOL_CACHE["brands"] = packs
    return packs


def resolve_school(db, query):
    """精确 → 别名 → slug → 拼音/首字母 → 子串 → 模糊。返回 (entry, score)。"""
    if not query:
        return None, 0
    q = str(query).strip()
    ql = q.lower()

    def name_match(entry, field, val):
        return str(entry.get(field) or "").lower() == ql

    for e in db:
        if name_match(e, "name", q):
            return e, 100
    for e in db:
        if name_match(e, "slug", q) or name_match(e, "pinyin", q) or name_match(e, "abbr", q):
            return e, 96
    for e in db:
        aliases = [str(a).lower() for a in (e.get("aliases") or [])]
        if ql in aliases:
            return e, 94
    for e in db:
        if q in str(e.get("name") or ""):
            return e, 88
    for e in db:
        for a in (e.get("aliases") or []):
            if a and q in str(a):
                return e, 82
    # 模糊
    best, score = None, 0
    for e in db:
        s = _sim(q, e.get("name", ""))
        s = max(s, _sim(ql, str(e.get("pinyin") or "")))
        if s > score:
            best, score = e, s
    if score >= 0.55:
        return best, int(score * 80)
    return None, 0


def _sim(a, b):
    import difflib
    return difflib.SequenceMatcher(None, str(a), str(b)).ratio()


DEFAULT_BRAND = {
    "school": "", "school_en": "", "primary": DEFAULT_PRIMARY, "accent": DEFAULT_ACCENT,
    "deep": darken(DEFAULT_PRIMARY, .35), "soft": lighten(DEFAULT_PRIMARY, .93),
    "ink": "#1A1A1A", "gray": "#5A5A5A", "lightgray": "#8A8A8A", "rule": "#C9D2DE",
    "white": "#FFFFFF", "gold": "#C8A15A", "motto": "", "footer_left": "",
    "footer_right": "", "decor": "rule", "skin_hint": "academic", "crest": "",
    "college_default": "",
}


def build_brand(skill_dir, profile, content):
    """合并：profile.brand_colors > 品牌包 > 中性默认。"""
    school = (profile.get("school") or "").strip()
    slug = (profile.get("school_slug") or "").strip()
    db = load_school_db(skill_dir)
    entry = None
    if slug:
        entry, _ = resolve_school(db, slug)
        # slug 与校名冲突时（拼音重码/旧档案），以校名为准，避免串到别的学校品牌包
        if entry is not None and school and str(entry.get("name") or "") != school:
            WARNINGS.warn("档案里的 school_slug（%s）与校名（%s）不一致，按校名匹配"
                          % (slug, school))
            entry = None
    if entry is None and school:
        entry, score = resolve_school(db, school)
    if entry is not None:
        slug = entry.get("slug")
        school = entry.get("name") or school
        if not profile.get("college"):
            profile["college"] = profile.get("college") or ""
    packs = load_brand_packs(skill_dir)
    bp = packs.get(slug or "", {}) if slug else {}
    if not bp and school:
        # 兜底：按学校名匹配品牌包（防止 schools.json 与 brand_packs.json 的 slug 不一致）
        for _k, _v in packs.items():
            if str(_v.get("school") or "") == school:
                bp = _v
                slug = slug or _k
                break
    b = dict(DEFAULT_BRAND)
    b["school"] = school or (bp.get("school") or "")
    b["school_en"] = (entry or {}).get("name_en") or bp.get("school_en") or ""
    b["slug"] = slug or ""
    b["province"] = (entry or {}).get("province", "")
    b["city"] = (entry or {}).get("city", "")
    b["level"] = (entry or {}).get("level", [])
    colors = dict(bp.get("colors") or {})
    user_colors = profile.get("brand_colors") or {}
    for key, val in user_colors.items():
        if isinstance(val, str) and re.match(r"^#?[0-9a-fA-F]{3,6}$", val.strip()):
            colors[key] = val
            colors["verified"] = True
    primary = colors.get("primary") or DEFAULT_PRIMARY
    primary = ensure_contrast(primary)
    if entry is None and not bp and school:
        WARNINGS.warn(f"模板库里没有匹配到「{school}」，使用中性配色 + 文字标识；"
                      f"可运行 --list-schools 查看已收录学校")
    accent = colors.get("accent") or mix(primary, "#FFFFFF", 0.28)
    if contrast(accent, "#FFFFFF") < 3.0:
        accent = darken(accent, 0.12)
    b.update({
        "primary": primary,
        "accent": accent,
        "deep": colors.get("deep") or darken(primary, 0.35),
        "soft": colors.get("soft") or lighten(primary, 0.93),
        "ink": colors.get("ink") or "#1A1A1A",
        "gold": colors.get("gold") or "#C8A15A",
        "motto": (profile.get("motto") or bp.get("motto") or "").strip(),
        "footer_left": (profile.get("footer_left") or bp.get("footer_left") or "").strip(),
        "footer_right": (profile.get("footer_right") or bp.get("footer_right") or "").strip(),
        "decor": bp.get("decor") or "rule",
        "skin_hint": bp.get("skin_hint") or "academic",
        "college_default": bp.get("college_default") or "",
        "crest_rel": bp.get("crest") or "",
        "crest_kind": bp.get("crest_kind") or ("image" if bp.get("crest") else "none"),
        "verified": bool(bp.get("verified")),
    })
    # 兜底：品牌包没登记校徽，但 assets/schools/crests/<slug>.png 存在时直接使用
    if not b.get("crest_rel") and b.get("slug"):
        guess = "crests/%s.png" % b["slug"]
        if os.path.isfile(os.path.join(skill_dir, "assets", "schools", guess)):
            b["crest_rel"] = guess
            b["crest_kind"] = "image"
    return b, entry, bp


# ------------------------------------------------------------------ 校徽 / 文字标识
IMG_SIZE_CACHE = {}


def img_size(path):
    if path in IMG_SIZE_CACHE:
        return IMG_SIZE_CACHE[path]
    size = None
    try:
        from PIL import Image
        with Image.open(path) as im:
            size = im.size
    except Exception:  # noqa: BLE001
        size = None
    IMG_SIZE_CACHE[path] = size
    return size


def fit_box(px, py, max_w, max_h, size, halign="center", valign="middle"):
    """按图片比例在给定框内等比缩放，返回 (x, y, w, h)。"""
    if not size:
        return px, py, max_w, max_h
    iw, ih = size
    if iw <= 0 or ih <= 0:
        return px, py, max_w, max_h
    scale = min(max_w / iw, max_h / ih)
    w, h = iw * scale, ih * scale
    if halign == "center":
        x = px + (max_w - w) / 2
    elif halign == "right":
        x = px + max_w - w
    else:
        x = px
    y = py + (max_h - h) / 2 if valign == "middle" else py
    return round(x, 2), round(y, 2), round(w, 2), round(h, 2)


FONT_CANDIDATES = [
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
]


def find_cjk_font():
    for p in FONT_CANDIDATES:
        if os.path.isfile(p):
            return p
    try:
        from matplotlib import font_manager  # type: ignore
        for f in font_manager.fontManager.ttflist:
            name = (f.name or "").lower()
            if any(k in name for k in ("noto sans cjk", "source han sans", "microsoft yahei", "simhei", "pingfang")):
                return f.fname
    except Exception:  # noqa: BLE001
        pass
    return None


def make_text_crest(school, primary, out_path, sub="", text_color=None, with_bar=True):
    """没有校徽时生成校名文字标识。返回 True/False。"""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        return False
    font_path = find_cjk_font()
    if not font_path:
        return False
    name = (school or sub or "学院汇报").strip()
    text = name[:14]
    color = text_color or primary
    size = 56 if len(text) <= 6 else 46 if len(text) <= 9 else 36 if len(text) <= 12 else 30
    big = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
    d0 = ImageDraw.Draw(big)
    font = ImageFont.truetype(font_path, size)
    bbox = d0.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad = 14
    acc_w = 10 if with_bar else 0
    canvas_w = tw + pad * 2 + acc_w + (16 if with_bar else 0)
    canvas_h = th + pad * 2
    canvas_w = max(canvas_w, 150)
    img = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    if with_bar:
        d.rectangle([0, 0, acc_w, canvas_h], fill=color)
    d.text((pad + acc_w + (4 if with_bar else 0) - bbox[0], pad - bbox[1]), text, font=font, fill=color)
    img.save(out_path)
    return True


def prepare_crest(src, dst):
    """把校徽整理成统一规格：白底转透明、裁到内容 bbox、限高。失败时直接复制。"""
    try:
        from PIL import Image
    except ImportError:
        shutil.copy(src, dst)
        return dst
    try:
        im = Image.open(src)
    except Exception:  # noqa: BLE001
        shutil.copy(src, dst)
        return dst
    try:
        im = im.convert("RGBA")
        w, h = im.size
        if w > 1400:
            im = im.resize((1400, max(1, int(h * 1400 / w))), Image.LANCZOS)
        px = im.load()
        iw, ih = im.size
        corners = [px[0, 0], px[iw - 1, 0], px[0, ih - 1], px[iw - 1, ih - 1]]
        has_alpha = im.getchannel("A").getextrema()[0] < 250
        if not has_alpha and all(c[:3] and min(c[:3]) > 235 for c in corners):
            out = im.copy()
            op = out.load()
            for y in range(ih):
                for x in range(iw):
                    r, g, b, _a = op[x, y]
                    if r > 238 and g > 238 and b > 238:
                        op[x, y] = (r, g, b, 0)
            bbox = out.getbbox()
            if bbox:
                out = out.crop(bbox)
            op = out.load()
            ow, oh = out.size
            # 透明包围盒外沿修一圈：避免白边
            out.save(dst)
            return dst
        bbox = im.getbbox()
        if bbox:
            im = im.crop(bbox)
        iw, ih = im.size
        if ih > 480:
            im = im.resize((max(1, int(iw * 480 / ih)), 480), Image.LANCZOS)
        im.save(dst)
        return dst
    except Exception:  # noqa: BLE001
        shutil.copy(src, dst)
        return dst


def resolve_crest(skill_dir, brand, profile, args, out_media_dir):
    """返回 media 目录里的 crest.png 路径（不存在则生成/复制）。"""
    candidates = []
    if getattr(args, "crest", None):
        candidates.append(os.path.abspath(args.crest))
    if profile.get("crest_override"):
        candidates.append(os.path.abspath(profile["crest_override"]))
    if brand.get("crest_rel"):
        candidates.append(os.path.join(skill_dir, "assets", "schools", brand["crest_rel"]))
    for c in candidates:
        if c and os.path.isfile(c):
            dst = os.path.join(out_media_dir, "crest.png")
            try:
                if img_size(c) is None:
                    WARNINGS.warn(f"校徽文件打不开，忽略：{c}")
                    continue
            except Exception:  # noqa: BLE001
                continue
            prepare_crest(c, dst)
            return dst, "image"
    if brand.get("school"):
        dst = os.path.join(out_media_dir, "crest.png")
        if make_text_crest(brand["school"], brand["primary"], dst):
            # 深色底用白色文字标识
            try:
                make_text_crest(brand["school"], brand["primary"],
                                os.path.join(out_media_dir, "crest_white.png"),
                                text_color="#FFFFFF")
            except Exception:  # noqa: BLE001
                pass
            return dst, "text"
    legacy = os.path.join(skill_dir, "assets", "college_crest.png")
    if os.path.isfile(legacy):
        dst = os.path.join(out_media_dir, "crest.png")
        prepare_crest(legacy, dst)
        return dst, "legacy"
    WARNINGS.warn("没有校徽也没有可用字体，封面将不显示校徽")
    return None, "none"


# ------------------------------------------------------------------ PPTD 元素
BRAND = dict(DEFAULT_BRAND)
SKIN = "academic"
CN = "Microsoft YaHei"
EN = "Arial"
MATH = "Times New Roman"
MOTTO_FONT = "Georgia"
PAGE_BG = "#FFFFFF"
MOTION_TARGETS = []


def P(key):
    return BRAND.get(key, DEFAULT_BRAND.get(key, "#000000"))


def ff(kind="cn"):
    if kind == "math":
        return {"latin": MATH, "ea": CN}
    if kind == "motto":
        return {"latin": MOTTO_FONT, "ea": CN}
    return {"latin": EN, "ea": CN}


def el(eid, etype, x, y, w, h, **kw):
    d = {"elementId": eid, "elementType": etype, "bounds": [round(x, 2), round(y, 2), round(w, 2), round(h, 2)]}
    d.update(kw)
    return d


def txt(eid, x, y, w, h, s, size, color=None, bold=False, align=("left", "top"),
        font="cn", lh=None, ls=None, wrap=True, motion=None):
    ha, va = align
    text = "" if s is None else str(s)
    if ha != "left" and not text.lstrip().startswith("<p"):
        text = '<p style="text-align:%s">%s</p>' % (ha, text)
    c = {"fontSize": size, "color": color or P("ink"), "bold": bold,
         "align": [ha, va], "fontFamily": ff(font)}
    if text:
        c["text"] = text
    if lh is not None:
        c["lineHeight"] = lh
    if ls is not None:
        c["letterSpacing"] = ls
    if wrap:
        c["wrap"] = True
    if motion:
        MOTION_TARGETS.append((eid, motion))
    return el(eid, "text", x, y, w, h, content=c)


def rect(eid, x, y, w, h, color, border=None, radius=None):
    kw = {"shapeName": "roundRect" if radius else "rect",
          "fill": {"type": "solid", "color": color}}
    if radius:
        kw["cornerRadius"] = radius
    if border:
        kw["border"] = {"style": "solid", "width": border[1], "color": border[0]}
    return el(eid, "shape", x, y, w, h, **kw)


def ellipse(eid, x, y, w, h, color, border=None):
    kw = {"shapeName": "ellipse", "fill": {"type": "solid", "color": color}}
    if border:
        kw["border"] = {"style": "solid", "width": border[1], "color": border[0]}
    return el(eid, "shape", x, y, w, h, **kw)


def line(eid, x, y, w, h, color, width=1, arrow=False):
    end = "arrow" if arrow else None
    return el(eid, "line", x, y, w, h, viewBox=[max(1, int(w)), max(1, int(h))],
              points="0,0 %d,%d" % (int(w), int(h)),
              border={"style": "solid", "width": width, "color": color}, arrow=[None, end])


def img(eid, x, y, w, h, src, motion=None, fit_mode="contain"):
    if motion:
        MOTION_TARGETS.append((eid, motion))
    return el(eid, "image", x, y, w, h, src=src, fit={"mode": fit_mode})


def img_fit(eid, box_x, box_y, box_w, box_h, src, halign="center", valign="middle", motion=None):
    size = img_size(os.path.join(OUT_MEDIA, os.path.basename(src))) if src else None
    if not size:
        size = img_size(src) if src and os.path.isfile(src) else None
    x, y, w, h = fit_box(box_x, box_y, box_w, box_h, size, halign, valign)
    if not size:
        # 图片缺失：画一个浅色占位框，页面不崩
        return rect(eid + "-missing", x, y, w, h, "#F2F4F8",
                    border=(WARNINGS and "#C9D2DE" or "#C9D2DE", 1))
    return img(eid, x, y, w, h, src, motion=motion)


def page(elements, ptype="content", notes="", background=None):
    return {"pageType": ptype,
            "background": {"type": "solid", "color": background or PAGE_BG},
            "notes": "" if notes is None else str(notes),
            "elements": elements}


# ------------------------------------------------------------------ 皮肤样式
def skin_metrics():
    return {
        "academic": {"body_top": 96, "title_size": 27, "title_x": 268, "title_y": 22,
                     "title_w": 636, "title_h": 40, "show_eyebrow": False},
        "campus": {"body_top": 96, "title_size": 26, "title_x": 244, "title_y": 20,
                   "title_w": 656, "title_h": 44, "show_eyebrow": False},
        "modern": {"body_top": 96, "title_size": 29, "title_x": 56, "title_y": 52,
                   "title_w": 848, "title_h": 64, "show_eyebrow": True},
        "bold": {"body_top": 106, "title_size": 26, "title_x": 232, "title_y": 20,
                 "title_w": 668, "title_h": 44, "show_eyebrow": False},
    }[SKIN]


def _crest_in(box_x, box_y, box_w, box_h, eid="crest", halign="center"):
    if not CREST_REL:
        return None
    return img_fit(eid, box_x, box_y, box_w, box_h, CREST_REL, halign=halign)


def _crest_with_w(box_x, box_y, box_w, box_h, eid="crest", halign="left"):
    """返回 (元素, 实际宽度)；用于按校徽宽度动态定位标题。"""
    if not CREST_REL:
        return None, 0
    p = os.path.join(OUT_MEDIA, os.path.basename(CREST_REL))
    size = img_size(p) or img_size(CREST_REL)
    x, y, w, h = fit_box(box_x, box_y, box_w, box_h, size, halign, "middle")
    return img(eid, x, y, w, h, CREST_REL), w


def _crest_in_plate(box_x, box_y, box_w, box_h, plate_color="#FFFFFF", pad=10,
                    radius=4, eid="crest"):
    """深色底上用：有图校徽放进白色底板；文字标识直接用白色版本，不加底板。返回元素列表。"""
    if not CREST_REL:
        return []
    if CREST_KIND == "text" and os.path.isfile(os.path.join(OUT_MEDIA, "crest_white.png")):
        return [img_fit(eid, box_x, box_y, box_w, box_h, "media/crest_white.png")]
    p = os.path.join(OUT_MEDIA, os.path.basename(CREST_REL))
    size = img_size(p) or img_size(CREST_REL)
    x, y, w, h = fit_box(box_x, box_y, box_w, box_h, size, "center", "middle")
    return [rect(eid + "bg", x - pad, y - pad, w + 2 * pad, h + 2 * pad, plate_color,
                 radius=radius),
            img(eid, x, y, w, h, CREST_REL)]


def _crest_aspect():
    p = os.path.join(OUT_MEDIA, os.path.basename(CREST_REL)) if CREST_REL else ""
    size = img_size(p) or img_size(CREST_REL)
    if not size or size[1] <= 0:
        return 1.0
    return size[0] / size[1]


def _title_text(pg):
    eb = str(pg.get("eyebrow") or "").strip()
    ti = str(pg.get("title") or "").strip()
    if not ti:
        return eb
    if eb and not re.match(r"^\d+(\.\d+)?[\s　]", ti):
        return f"{eb}　{ti}"
    return ti


def chrome(pg, num, total, show_num=True):
    """统一页眉：左侧「编号 + 标题」，右侧校徽，标题下一条细线；正文区从 96 开始。"""
    title = _title_text(pg)
    n = len(title)
    size = 28 if n <= 20 else 25 if n <= 28 else 22
    e = []
    e.append(txt("title", MARGIN, 24, 648, 44, title, size, P("primary"), bold=True, lh=1.15))
    if CREST_REL:
        c = _crest_in(716, 18, 176, 40, halign="right")
        if c:
            e.append(c)
    e.append(rect("hrule", MARGIN, 74, 848, 1, mix(P("rule"), "#FFFFFF", 0.15)))
    e += footer(pg, num, total, show_num)
    return e


def footer(pg, num, total, show_num=True, dark=False, bg=None):
    left = (pg.get("_footer_left") or BRAND.get("footer_left") or BRAND.get("motto") or
            BRAND.get("school"))
    right = (pg.get("_footer_right") or BRAND.get("footer_right") or BRAND.get("school_en") or "")
    if not right and profile_get("college_en"):
        right = profile_get("college_en")
    e = []
    if dark:
        base = bg or P("deep")
        foot_color = mix("#FFFFFF", base, 0.3)
    else:
        foot_color = mix(P("lightgray"), P("ink"), 0.35)
    if SKIN == "academic":
        if left:
            e.append(txt("motto1", MARGIN, 512, 260, 12, left, SZ["foot"], foot_color,
                         font="motto", ls=0.5))
        if right:
            e.append(txt("motto2", 440, 512, 464, 12, right, SZ["foot"], foot_color,
                         font="motto", align=("right", "top")))
    elif SKIN == "campus":
        if left:
            e.append(txt("motto1", MARGIN, 510, 400, 12, left, SZ["foot"], foot_color, ls=0.3))
        if right:
            e.append(txt("motto2", 500, 510, 404, 12, right, SZ["foot"], foot_color,
                         align=("right", "top")))
    elif SKIN == "modern":
        if left:
            e.append(txt("motto1", MARGIN, 510, 400, 12, left, SZ["foot"], foot_color))
        if right:
            e.append(txt("motto2", 500, 510, 404, 12, right, SZ["foot"], foot_color,
                         align=("right", "top")))
        if show_num:
            e.append(txt("pn", 846, 510, 58, 12, f"{num} / {total}", SZ["foot"] + 1, foot_color,
                         align=("right", "top")))
        return e
    else:
        pass
        if left:
            e.append(txt("motto1", 40, 508, 420, 12, left, SZ["foot"], foot_color))
        if right:
            e.append(txt("motto2", 480, 508, 440, 12, right, SZ["foot"], foot_color,
                         align=("right", "top")))
    if show_num:
        e.append(txt("pn", 846, 486, 58, 12, f"{num} / {total}", SZ["foot"] + 1, foot_color,
                     align=("right", "top")))
    return e


# ------------------------------------------------------------------ 页面构建
def p_cover(pg, total, num):
    """封面：典型汇报模板——校徽校名在上，色带压主标题，报告人/日期在下。"""
    title = pg.get("title") or ""
    subtitle = pg.get("subtitle") or ""
    college = pg.get("_college") or ""
    school = BRAND.get("school") or ""
    reporter = pg.get("_reporter") or ""
    date_str = pg.get("_date") or ""
    citation = pg.get("citation") or ""
    school_line = " · ".join([x for x in (school, college) if x])
    text_crest = CREST_KIND == "text" and bool(CREST_REL)
    if text_crest:
        school_line = college or ""
    e = []
    if SKIN == "academic":
        if CREST_REL and _crest_aspect() >= 2.2:
            c = _crest_in(240, 46, 480, 92)
        elif CREST_REL:
            c = _crest_in(412, 42, 136, 96)
        else:
            c = None
        if c:
            e.append(c)
        if school_line:
            e.append(txt("school", 0, 152, W, 24, school_line, 14, P("gray"),
                         align=("center", "top")))
        band_top = 196
        e.append(rect("titleband", 0, band_top, W, 122, P("primary")))
        e.append(txt("cover-title", 70, band_top + 8, 820, 106, title, 31, "#FFFFFF", bold=True,
                     align=("center", "middle"), lh=1.3, motion="zoom"))
        y = band_top + 140
        if subtitle:
            e.append(txt("cover-sub", 70, y, 820, 28, subtitle, 15, P("gray"),
                         align=("center", "top"), lh=1.4, motion="fade"))
            y += 42
        meta = "报告人：%s　　　%s" % (reporter or "＿＿＿＿", date_str) if date_str else \
            ("报告人：%s" % (reporter or "＿＿＿＿"))
        e.append(txt("cover-meta", 0, y, W, 26, meta, 16, P("ink"), bold=True,
                     align=("center", "top")))
        if citation:
            e.append(txt("cite", MARGIN, 484, 848, 24, citation, 8, P("lightgray"),
                         align=("left", "top"), lh=1.4))
    elif SKIN == "campus":
        e.append(rect("cover-band", 0, 0, W, 196, P("primary")))
        e += _crest_in_plate(64, 48, 180, 92, plate_color="#FFFFFF", pad=8, radius=4)
        if school_line:
            if text_crest:
                e.append(txt("college", 300, 82, 600, 26, college, 15,
                             mix("#FFFFFF", P("primary"), 0.25)))
            else:
                e.append(txt("school", 300, 64, 600, 30, school, 24, "#FFFFFF", bold=True))
                e.append(txt("college", 300, 102, 600, 22, college, 13,
                             mix("#FFFFFF", P("primary"), 0.25)))
        e.append(txt("cover-title", 56, 240, 848, 84, title, 31, P("ink"), bold=True, lh=1.3,
                     motion="zoom"))
        if subtitle:
            e.append(txt("cover-sub", 56, 330, 848, 28, subtitle, 15, P("gray"), lh=1.4,
                         motion="fade"))
        meta = "报告人：%s　　　%s" % (reporter or "＿＿＿＿", date_str)
        e.append(txt("cover-meta", 56, 384, 848, 24, meta, 14, P("gray")))
        if citation:
            e.append(txt("cite", MARGIN, 484, 848, 24, citation, 8, P("lightgray"), lh=1.4))
    elif SKIN == "modern":
        if CREST_REL:
            c = _crest_in(770, 42, 110, 100, halign="right")
            if c:
                e.append(c)
        if school_line:
            e.append(txt("school", MARGIN, 92, 700, 22, school_line, 12, P("gray"), ls=0.6))
        e.append(txt("cover-title", MARGIN, 140, 848, 130, title, 36, P("ink"), bold=True,
                     lh=1.3, motion="zoom"))
        if subtitle:
            e.append(txt("cover-sub", MARGIN, 292, 848, 28, subtitle, 16, P("gray"), lh=1.4,
                         motion="fade"))
        meta = "报告人：%s" % (reporter or "＿＿＿＿")
        e.append(txt("cover-meta", MARGIN, 360, 480, 60, meta + "\n" + date_str, 13, P("ink"),
                     lh=1.8))
        if citation:
            e.append(txt("cite", 500, 360, 404, 60, citation, 8.5, P("lightgray"),
                         align=("right", "top"), lh=1.5))
    else:  # bold
        e.append(rect("bg", 0, 0, W, H, P("deep")))
        e += _crest_in_plate(64, 56, 190, 92, plate_color="#FFFFFF", pad=8, radius=4)
        if school_line:
            if text_crest:
                e.append(txt("college", 300, 86, 620, 26, college, 15,
                             mix("#FFFFFF", P("deep"), 0.2)))
            else:
                e.append(txt("school", 300, 66, 620, 30, school, 24, "#FFFFFF", bold=True))
                e.append(txt("college", 300, 104, 620, 22, college, 13,
                             mix("#FFFFFF", P("deep"), 0.2)))
        e.append(txt("cover-title", 56, 210, 700, 150, title, 34, "#FFFFFF", bold=True, lh=1.3,
                     motion="zoom"))
        if subtitle:
            e.append(txt("cover-sub", 56, 366, 700, 30, subtitle, 15,
                         mix("#FFFFFF", P("deep"), 0.2), lh=1.4, motion="fade"))
        meta = "报告人：%s　　　%s" % (reporter or "＿＿＿＿", date_str)
        e.append(txt("cover-meta", 56, 426, 700, 22, meta, 13, mix("#FFFFFF", P("deep"), 0.3)))
        if citation:
            e.append(txt("cite", 56, 484, 848, 24, citation, 8, mix("#FFFFFF", P("deep"), 0.4),
                         lh=1.4))
    e += footer(pg, num, total, show_num=False, dark=(SKIN == "bold"),
                bg=P("deep") if SKIN == "bold" else None)
    return page(e, "cover", pg.get("notes", ""),
                background=P("deep") if SKIN == "bold" else "#FFFFFF")


def p_paper_card(pg, total, num):
    p = pg.get("paper") or {}
    title = pg.get("title") or "论文简介"
    e = chrome(pg, num, total)
    x, w = MARGIN, 484
    y = 96
    e.append(txt("l1", x, y, w, 16, "论文题目", 10.5, P("gray")))
    e.append(txt("ltcn", x, y + 14, w, 44, p.get("title_cn", ""), 15, P("ink"), bold=True, lh=1.4))
    y += 60
    if p.get("title_en"):
        e.append(txt("lten", x, y, w, 34, p["title_en"], 9.5, P("gray"), lh=1.45))
        y += 38
    for label, key, base in (("作者", "authors", 30), ("单位", "affiliations", 30),
                             ("关键词", "keywords", 18)):
        if not p.get(key):
            continue
        if key == "keywords":
            size = 10.5
        elif key == "affiliations":
            size = 10.5
        else:
            size = 11.5
        lines = est_lines(str(p[key]), w - 56, size)
        hgt = max(base, int(lines * size * 1.55) + 6)
        if y + hgt > 430:
            hgt = max(24, 430 - y)
        e.append(txt("lb-" + key, x, y, 56, 16, label, 10.5, P("gray")))
        e.append(txt(key, x + 56, y - 2, w - 56, hgt, p[key], size, P("ink"), lh=1.55))
        y += hgt + 4
    if p.get("summary"):
        e.append(txt("sumlb", x, 352, w, 16, "一句话", 10.5, P("gray")))
        e.append(txt("sum", x, 370, w, 86, p["summary"], 12.5, P("ink"),
                     align=("left", "top"), lh=1.75))
    # 右栏
    rx, rw = 584, 320
    ry = 96
    e.append(rect("vline", 556, 96, 1, 344, P("rule")))
    e.append(txt("r1", rx, ry, rw, 16, "发表期刊", 10.5, P("gray")))
    e.append(txt("jname", rx, ry + 14, rw, 26, p.get("venue") or p.get("journal") or "", 12,
                 P("primary"), bold=True))
    ry += 48
    if p.get("volume"):
        e.append(txt("jvol", rx, ry, rw, 24, p["volume"], 10.5, P("ink"), lh=1.4))
        ry += 28
    if p.get("doi"):
        e.append(txt("jdoi", rx, ry, rw, 30, "DOI：" + p["doi"], 9, P("gray"), lh=1.45))
        ry += 34
    if p.get("ranks"):
        e.append(txt("r2", rx, ry, rw, 16, "期刊级别", 10.5, P("gray")))
        ry += 16
        for i, s in enumerate(p["ranks"][:4]):
            e.append(txt("rk%d" % i, rx, ry + i * 21, rw, 18, s, 10.5, P("ink")))
        ry += 21 * min(4, len(p["ranks"])) + 6
    if p.get("metrics"):
        e.append(txt("r3", rx, ry, rw, 16, "影响因子与指标", 10.5, P("gray")))
        e.append(txt("ifv", rx, ry + 14, rw, 22, p["metrics"], 12, P("primary"), bold=True))
        ry += 40
        if p.get("metrics_note"):
            e.append(txt("ifn", rx, ry, rw, 30, p["metrics_note"], 9, P("gray"), lh=1.45))
            ry += 34
    if p.get("funding"):
        fy = min(420, ry + 8)
        e.append(txt("fund", rx, fy, rw, 40, p["funding"], 8.5, P("gray"), lh=1.4))
    if p.get("source_note"):
        e.append(txt("src", MARGIN, 466, 848, 14, p["source_note"], 8.5, P("lightgray")))
    return page(e, "content", pg.get("notes", ""))


def p_toc(pg, total, num):
    items = pg.get("items") or []
    if not pg.get("title"):
        pg = dict(pg)
        pg["title"] = "目录"
    e = chrome(pg, num, total)
    n = max(1, len(items))
    top = 112
    avail = 436 - top
    row = min(112, avail / n)
    for i, it in enumerate(items):
        y = top + i * row
        if SKIN == "modern":
            e.append(txt(f"n{i}", MARGIN, y - 8, 90, 40, it.get("num", str(i + 1)), 28,
                         mix(P("primary"), "#FFFFFF", 0.9), bold=True))
            e.append(txt(f"t{i}", 126, y, 600, 28, it.get("title", ""), 19, P("ink"), bold=True))
        else:
            e.append(txt(f"n{i}", MARGIN, y, 40, 24, it.get("num", str(i + 1)), 16, P("accent"), bold=True))
            e.append(txt(f"t{i}", 96, y - 2, 560, 26, it.get("title", ""), 17, P("primary"), bold=True))
        if it.get("pages"):
            e.append(txt(f"p{i}", 760, y + 2, 144, 18, it["pages"], 10.5, P("gray"),
                         align=("right", "top")))
        if it.get("subs"):
            e.append(txt(f"d{i}", 120 if SKIN != "modern" else 126, y + 28, 680, 20,
                         it["subs"], 10.5, P("gray"), lh=1.35))

    if pg.get("note"):
        e.append(txt("note", MARGIN, 448, 848, 18, pg["note"], 10, P("gray")))
    return page(e, "table_of_contents", pg.get("notes", ""))


def p_chapter(pg, total, num):
    """章节过渡：只放「第 N 节 / 标题 / 一句副题 / 本节提纲」，不做大数字水印和装饰。"""
    num_s = str(pg.get("num") or pg.get("chapter") or "")
    title = pg.get("title") or ""
    sub = pg.get("sub") or ""
    outline = [str(x) for x in (pg.get("bullets") or []) if x][:4]
    dark = SKIN in ("campus", "bold")
    bg = P("deep") if SKIN == "bold" else (P("primary") if SKIN == "campus" else "#FFFFFF")
    e = []
    if dark:
        e.append(rect("bg", 0, 0, W, H, bg))
        e += _crest_in_plate(56, 34, 150, 40, plate_color="#FFFFFF", pad=7, radius=3)
        if num_s:
            e.append(txt("sec", MARGIN, 150, 640, 24, "第 %s 节" % num_s, 13,
                         mix("#FFFFFF", bg, 0.35), ls=1.2))
        e.append(txt("title", MARGIN, 186, 800, 90, title, 34, "#FFFFFF", bold=True, lh=1.3,
                     motion="fade"))
        if sub:
            e.append(txt("sub", MARGIN, 288, 720, 40, sub, 15, mix("#FFFFFF", bg, 0.3), lh=1.6))
        y = 360
        for i, b in enumerate(outline):
            e.append(txt(f"o{i}", MARGIN, y, 720, 26, "·　" + b, 13, mix("#FFFFFF", bg, 0.18), lh=1.5))
            y += 34
    else:
        if CREST_KIND == "text" and CREST_REL:
            e += _crest_in_plate(56, 26, 150, 40, plate_color="#FFFFFF", pad=6, radius=3)
        elif CREST_REL:
            e.append(_crest_in(56, 26, 150, 40))
        if num_s:
            e.append(txt("sec", MARGIN, 150, 640, 24, "第 %s 节" % num_s, 13, P("lightgray"), ls=1.2))
        e.append(txt("title", MARGIN, 184, 800, 90, title, 34, P("ink"), bold=True, lh=1.3,
                     motion="fade"))
        if sub:
            e.append(txt("sub", MARGIN, 284, 720, 40, sub, 15, P("gray"), lh=1.6))
        y = 356
        for i, b in enumerate(outline):
            e.append(rect(f"d{i}", MARGIN, y + 9, 6, 6, P("accent")))
            e.append(txt(f"o{i}", MARGIN + 16, y, 720, 26, b, 14, P("ink"), lh=1.5))
            y += 36
    e += footer(pg, num, total, show_num=False, dark=dark, bg=bg if dark else None)
    return page(e, "chapter", pg.get("notes", ""), background=bg)


def _lead_strip(e, pg, y=114):
    lead = pg.get("lead")
    if not lead:
        return y
    e.append(rect("leadbg", MARGIN, y - 4, 848, 38, P("soft"), radius=3))

    e.append(txt("lead", MARGIN + 14, y + 1, 820, 28, lead, 14, P("deep") if SKIN == "modern" else P("primary"),
                 lh=1.4))
    return y + 52


def _item_marker(e, i, x, y, layout):
    if SKIN == "campus":
        e.append(ellipse(f"m{i}", x, y + 3, 18, 18, P("soft")))
        e.append(txt(f"mn{i}", x, y, 18, 20, str(i + 1), 10, P("primary"), bold=True,
                     align=("center", "middle")))
    elif SKIN == "modern":
        e.append(rect(f"m{i}", x, y + 4, 10, 10, P("accent")))
    elif SKIN == "bold":
        e.append(rect(f"m{i}", x - 18, y - 4, 5, 28, P("primary")))
    else:
        e.append(rect(f"m{i}", x, y + 7, 9, 9, P("accent")))


def p_bullets(pg, total, num):
    e = chrome(pg, num, total)
    layout = pg.get("layout") or "rows"
    items = [it for it in (pg.get("items") or []) if it]
    if not items:
        WARNINGS.warn("bullets 页没有 items：" + str(pg.get("title")))
        return page(e, "content", pg.get("notes", ""))
    if layout == "split" and pg.get("fig"):
        return _bullets_split(pg, total, num, e, items)
    if layout == "cards":
        return _bullets_cards(pg, total, num, e, items)
    body = skin_metrics()["body_top"]
    y = _lead_strip(e, pg, body)
    n = len(items)
    k_size = 19 if n <= 4 else 17.5
    t_size = 17.5 if n <= 4 else 16
    kw = 180
    tx = MARGIN + 34 + kw
    tw = 904 - tx

    def _plan(ks, ts):
        lines = [max(1, est_lines(str(it.get("t") or it.get("p") or ""), tw, ts)) for it in items]
        return [max(58, 30 + 23 * ln) for ln in lines]

    row_hs = _plan(k_size, t_size)
    avail = 452 - y
    if sum(row_hs) > avail:
        k_size, t_size = 16.5, 15
        row_hs = _plan(k_size, t_size)
    if sum(row_hs) > avail:
        # 仍超出：压缩到均分，字号再降一档
        k_size, t_size = 15.5, 14
        row_hs = [avail / n] * n
        tight = True
        tight = False
    # 行距保持自然（不把剩余空间均摊到每行），底部留白；需要时用 takeaway 收一句
    start = y
    yy = start
    for i, it in enumerate(items):
        k = str(it.get("k") or it.get("h") or "")
        t = str(it.get("t") or it.get("p") or "")
        _item_marker(e, i, MARGIN + (18 if SKIN == "bold" else 3), yy + 4, layout)
        mx = MARGIN + (34 if SKIN != "bold" else 40)
        extra_s = str(it.get("num") or "")
        e.append(txt(f"k{i}", mx, yy, kw + 1, 26,
                     f'<span style="font-weight:bold">{k}</span>', k_size,
                     P("primary"), lh=1.35, motion="fade"))
        if extra_s:
            e.append(txt(f"num{i}", mx, yy + row_hs[i] - 24, kw, 18, extra_s, 12, P("accent"), bold=True))
        e.append(txt(f"t{i}", tx, yy + 1, tw, max(24, row_hs[i] - 8), t, t_size, P("ink"),
                     align=("left", "top"), lh=1.55, motion="fade"))
        yy += row_hs[i]
    if pg.get("tag"):
        e.append(txt("tag", 760, 92, 144, 16, pg["tag"], 10, P("gray"), align=("right", "top")))
    if pg.get("takeaway"):
        e.append(txt("takeaway", MARGIN, 448, 848, 22, str(pg["takeaway"]), 12.5,
                     mix(P("gray"), P("ink"), 0.25), lh=1.4))
    return page(e, "content", pg.get("notes", ""))


def _bullets_cards(pg, total, num, e, items):
    body = skin_metrics()["body_top"]
    y = _lead_strip(e, pg, body)
    n = min(4, len(items))
    gap = 16
    cw = (848 - gap * (n - 1)) / n
    k_texts = [str(it.get("k") or it.get("h") or "") for it in items[:n]]
    t_texts = [str(it.get("t") or it.get("p") or "") for it in items[:n]]
    t_lines = max(1, max(est_lines(t, cw - 32, 13.5) for t in t_texts))
    k_lines = max(1, max(est_lines(k, cw - 32, 16) for k in k_texts))
    ch = min(260, max(150, 34 + 24 * k_lines + 26 + 22 * t_lines))
    avail = 452 - y
    y = y + max(0, (avail - ch) / 2)
    for i, it in enumerate(items[:n]):
        x = MARGIN + i * (cw + gap)
        e.append(rect(f"card{i}", x, y, cw, ch, P("soft"), radius=6))
        e.append(rect(f"cardtop{i}", x, y, cw, 4, P("primary")))
        k = k_texts[i]
        t = t_texts[i]
        e.append(txt(f"k{i}", x + 16, y + 18, cw - 32, 30 + 24 * k_lines, k, 16, P("primary"),
                     bold=True, lh=1.35, motion="fade"))
        e.append(txt(f"t{i}", x + 16, y + 30 + 24 * k_lines + 8, cw - 32, ch - 60 - 24 * k_lines,
                     t, 13.5, P("ink"), align=("left", "top"), lh=1.6))
    return page(e, "content", pg.get("notes", ""))


def _bullets_split(pg, total, num, e, items):
    body = skin_metrics()["body_top"]
    e = e[:len(e) - 0] if False else e
    y = _lead_strip(e, pg, body)
    fw = pg.get("fig_w", 400)
    fh = pg.get("fig_h", 288)
    fx = 904 - fw
    e.append(img_fit("fig", fx, y + 4, fw, fh, "media/" + os.path.basename(pg["fig"]), motion="zoom"))
    if pg.get("caption"):
        e.append(txt("cap", fx, y + fh + 8, fw, 16, pg["caption"], SZ["cap"], P("gray"),
                     align=("center", "top")))
    n = len(items)
    roww = min(84, (452 - y) / max(1, n))
    for i, it in enumerate(items):
        yy = y + i * roww
        _item_marker(e, i, MARGIN + 2, yy + 4, "rows")
        e.append(txt(f"k{i}", MARGIN + 34, yy, 330, 24, str(it.get("k", "")), 16, P("primary"),
                     bold=True, lh=1.35, motion="fade"))
        e.append(txt(f"t{i}", MARGIN + 34, yy + 24, 340, 40, str(it.get("t", "")), 13, P("ink"),
                     lh=1.5))
    return page(e, "content", pg.get("notes", ""))


def p_keynums(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    items = (pg.get("items") or [])[:4]
    n = max(1, len(items))
    gap = 20
    cw = (848 - gap * (n - 1)) / n
    avail = 452 - body
    top = body + max(20, (avail - 220) / 2) + 34
    for i, it in enumerate(items):
        x = MARGIN + i * (cw + gap)
        card = SKIN in ("campus", "modern")
        if card:
            e.append(rect(f"card{i}", x, top - 40, cw, 214, P("soft"), radius=6))
        num_s = str(it.get("num", ""))
        e.append(txt(f"num{i}", x + 8, top - 10, cw - 16, 84, num_s, 46 if len(num_s) <= 5 else 36,
                     P("primary"), bold=True, lh=1.1, motion="zoom"))
        if it.get("unit"):
            e.append(txt(f"unit{i}", x + 10, top + 70, cw - 20, 22, it["unit"], 13, P("accent"), bold=True))
        e.append(txt(f"t{i}", x + 10, top + 100, cw - 20, 70, it.get("t", ""), 13.5, mix(P("gray"), P("ink"), 0.2), lh=1.6))
        pass
    if pg.get("note"):
        e.append(txt("note", MARGIN, 412, 848, 32, pg["note"], 12, P("gray"), lh=1.6))
    return page(e, "content", pg.get("notes", ""))


def p_text(pg, total, num):
    e = chrome(pg, num, total)
    y = _lead_strip(e, pg, skin_metrics()["body_top"])
    blocks = pg.get("blocks") or []
    n = max(1, len(blocks))
    for i, b in enumerate(blocks):
        yy = y
        if b.get("h"):
            e.append(txt(f"h{i}", MARGIN, yy, 848, 26, b["h"], 18.5, P("primary"), bold=True))
            yy += 30
        body_text = str(b.get("p", ""))
        size = 17.5 if n <= 2 else 16
        lines = est_lines(body_text, 848, size)
        hgt = b.get("hgt") or int(lines * size * 1.55 + 6)
        e.append(txt(f"p{i}", MARGIN, yy, 848, hgt, body_text, size, P("ink"),
                     align=("justify", "top"), lh=1.8))
        y = yy + hgt + 18
    if pg.get("takeaway"):
        e.append(txt("takeaway", MARGIN, 448, 848, 22, str(pg["takeaway"]), 12.5,
                     mix(P("gray"), P("ink"), 0.25), lh=1.4))
    return page(e, "content", pg.get("notes", ""))


def p_left_fig(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    y = _lead_strip(e, pg, body)
    blocks = pg.get("blocks") or []
    fw = pg.get("fig_w", 444)
    fh = pg.get("fig_h", 250)
    fx = 904 - fw
    e.append(img_fit("fig", fx, y + 2, fw, fh, "media/" + os.path.basename(pg["fig"]), motion="zoom"))
    if pg.get("caption"):
        e.append(txt("cap", fx, y + fh + 10, fw, 18, pg["caption"], SZ["cap"], P("gray"),
                     align=("center", "top")))
    if pg.get("source"):
        e.append(txt("source", fx, y + fh + 28, fw, 14, pg["source"], 8, P("lightgray"),
                     align=("center", "top")))
    yb = y + 4
    for i, b in enumerate(blocks):
        e.append(txt(f"h{i}", MARGIN, yb, 380, 24, b.get("h", ""), SZ["sub"], P("primary"), bold=True))
        e.append(txt(f"p{i}", MARGIN, yb + 26, 384, b.get("h2", 64), b.get("p", ""), 16, P("ink"),
                     align=("left", "top"), lh=1.7))
        yb += 118
    return page(e, "content", pg.get("notes", ""))


def p_big_fig(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    src = "media/" + os.path.basename(pg["fig"])
    lines = pg.get("lines") or []
    size = img_size(os.path.join(OUT_MEDIA, os.path.basename(pg["fig"])))
    if not size:
        size = img_size(pg["fig"])
    avail_h = 452 - body - (22 if pg.get("caption") else 0) - (16 if pg.get("source") else 0) \
        - 24 * len(lines) - 8
    max_w = 780
    w, h = (max_w, max_w * size[1] / size[0]) if size else (720, 300)
    if h > avail_h:
        h = max(120, avail_h)
        w = h * size[0] / size[1] if size else max_w
    if w > max_w:
        w = max_w
        h = w * size[1] / size[0] if size else h
    fx = (W - w) / 2
    if size:
        e.append(img("fig", fx, body, w, h, src, motion="zoom"))
    else:
        e.append(rect("fig-missing", fx, body, w, h, "#F2F4F8", border=("#C9D2DE", 1)))
        WARNINGS.warn("图片不存在（已用占位框）：%s" % pg["fig"])
    yy = body + h + 10
    if pg.get("caption"):
        e.append(txt("cap", 56, yy, 848, 18, pg["caption"], SZ["cap"], P("gray"), align=("center", "top")))
        yy += 22
    if pg.get("source"):
        e.append(txt("source", 56, yy, 848, 14, pg["source"], 8, P("lightgray"), align=("center", "top")))
        yy += 16
    for i, ln in enumerate(lines[:2]):
        e.append(txt(f"l{i}", 56, yy + i * 24, 848, 22, ln, 13, P("ink"), align=("center", "top"), lh=1.4))
    return page(e, "content", pg.get("notes", ""))


def p_two_figs(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    fw = pg.get("fig_w", 372)
    fh = pg.get("fig_h", 280)
    xs = [64, 524]
    for i in range(min(2, len(pg.get("figs") or []))):
        f = pg["figs"][i]
        e.append(img_fit(f"fig{i}", xs[i], body, fw, fh, "media/" + os.path.basename(f["file"]),
                         motion="zoom"))
        e.append(txt(f"cap{i}", xs[i], body + fh + 8, fw, 18, f.get("caption", ""), SZ["cap"],
                     P("gray"), align=("center", "top")))
    if pg.get("source"):
        e.append(txt("source", 64, body + fh + 28, 832, 14, pg["source"], 8, P("lightgray"),
                     align=("center", "top")))
    for i, ln in enumerate((pg.get("lines") or [])[:2]):
        e.append(txt(f"l{i}", 64, body + fh + 46 + i * 22, 832, 20, ln, 12.5, P("ink"),
                     align=("center", "top"), lh=1.4))
    return page(e, "content", pg.get("notes", ""))


def p_table(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    cols = pg.get("columns") or [0.4, 0.6]
    tw = pg.get("tbl_w", 520)
    th = pg.get("tbl_h", 240)
    rows = pg.get("rows") or []
    cells = []
    for row in rows:
        cells.append([{"text": str(v), "fill": {"type": "solid", "color": "#FFFFFF"},
                       "align": ["left", "middle"], "border": None} for v in row])
    e.append(el("tbl", "table", MARGIN, body, tw, th, columnWidths=cols,
                rowHeights=[1 / max(1, len(rows))] * max(1, len(rows)), rows=cells,
                style={
                    "cellStyle": {"fontSize": 12, "color": P("ink"), "align": ["left", "middle"],
                                  "fill": {"type": "solid", "color": "#FFFFFF"}, "border": None},
                    "firstRowStyle": {"fontSize": 11.5, "color": P("primary"), "bold": True,
                                      "fill": {"type": "solid", "color": P("soft")},
                                      "border": [{"style": "solid", "width": 1.1, "color": P("primary")},
                                                 None,
                                                 {"style": "solid", "width": 0.8, "color": P("primary")},
                                                 None]},
                    "lastRowStyle": {"fill": {"type": "solid", "color": "#FFFFFF"},
                                     "border": [None, None,
                                                {"style": "solid", "width": 1.1, "color": P("primary")},
                                                None]},
                }))
    if pg.get("tnote"):
        e.append(txt("tnote", MARGIN, body + th + 12, tw + 40, 34, pg["tnote"], SZ["note"],
                     P("gray"), lh=1.6))
    if pg.get("side"):
        e.append(rect("vline", 612, body, 1, 300, P("rule")))
        e.append(txt("sh", 644, body - 4, 260, 20, pg["side"].get("h", ""), 13, P("accent"), bold=True))
        y = body + 30
        for i, it in enumerate(pg["side"].get("items") or []):
            e.append(txt(f"sk{i}", 644, y, 260, 18, str(it.get("k") or it.get("h") or ""), 12.5,
                         P("primary"), bold=True))
            e.append(txt(f"sd{i}", 644, y + 20, 260, 28, str(it.get("t") or it.get("d") or ""), 10.5,
                         P("gray"), lh=1.5))
            y += 58
    return page(e, "content", pg.get("notes", ""))


def p_formula(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    y0 = _lead_strip(e, pg, body)
    formulas = pg.get("formulas") or []
    gap = 42
    block_h = sum((46 if not f.get("label") else 22) for f in formulas)
    note_h = 0
    if pg.get("note"):
        note_h = int(est_lines(str(pg["note"]), 848, SZ["note"]) * SZ["note"] * 1.75) + 16
    avail = 452 - y0
    y = y0 + max(0, (avail - block_h - note_h) / 2)
    for i, f in enumerate(formulas):
        if f.get("label"):
            e.append(txt(f"lb{i}", MARGIN, y, 848, 18, f["t"], 11.5, P("gray")))
            y += 22
            continue
        if f.get("img"):
            # LaTeX 公式图：按 288dpi 渲染，显示时缩到 1/4（1pt=1px），与正文字号一致
            src = "media/" + os.path.basename(str(f["img"]))
            size = img_size(os.path.join(OUT_MEDIA, os.path.basename(str(f["img"]))))
            if size and size[0]:
                w, h = size[0] / 4.0, size[1] / 4.0
                if w > 760:
                    k = 760.0 / w
                    w, h = w * k, h * k
            else:
                w, h = 600.0, 40.0
            e.append(img(f"f{i}", MARGIN, y, w, h, src))
        else:
            e.append(txt(f"f{i}", MARGIN, y, 750, 34, f["t"], 19, P("ink"), font="math"))
        if f.get("no"):
            # 编号统一贴右边界右对齐（真 LaTeX 排版习惯）
            e.append(txt(f"n{i}", 790, y - 2, 114, 22, f["no"], 12.5, P("gray"), font="math",
                         align=("right", "top")))
        y += gap
    if pg.get("note"):
        e.append(txt("note", MARGIN, y + gap - 6, 848, 54, pg["note"], SZ["note"], P("gray"), lh=1.75))
    return page(e, "content", pg.get("notes", ""))


def p_compare(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    y = _lead_strip(e, pg, body)
    cols = (pg.get("cols") or [])[:3]
    n = max(1, len(cols))
    gap = 18
    cw = (848 - gap * (n - 1)) / n
    max_items = max([min(5, len(c.get("items") or [])) for c in cols] + [1])
    ch = min(320, 62 + 46 * max_items + 16)
    avail = 452 - y
    y = y + max(0, (avail - ch) / 2)
    for i, col in enumerate(cols):
        x = MARGIN + i * (cw + gap)
        hl = bool(col.get("highlight"))
        e.append(rect(f"col{i}", x, y, cw, ch, "#F7F9FC" if not hl else P("soft")))
        e.append(rect(f"colh{i}", x, y, cw, 44, P("primary") if hl else P("soft")))
        e.append(txt(f"colt{i}", x + 14, y + 10, cw - 28, 26, col.get("h", ""), 14,
                     "#FFFFFF" if hl else P("primary"), bold=True))
        yy = y + 62
        for j, it in enumerate((col.get("items") or [])[:5]):
            e.append(rect(f"dot{i}_{j}", x + 16, yy + 8, 6, 6, P("accent")))
            e.append(txt(f"it{i}_{j}", x + 30, yy, cw - 44, 40, str(it), 13, P("ink"), lh=1.55))
            yy += 46
    return page(e, "content", pg.get("notes", ""))


def p_timeline(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    steps = (pg.get("steps") or [])[:5]
    n = max(1, len(steps))
    y = body + 110
    e.append(rect("tl", MARGIN, y, 848, 2, P("rule")))
    seg = 848 / n
    for i, s in enumerate(steps):
        cx = MARGIN + seg * (i + 0.5)
        e.append(ellipse(f"node{i}", cx - 9, y - 8, 18, 18, P("primary")))
        e.append(txt(f"h{i}", cx - seg / 2 + 6, y - 62, seg - 12, 40, str(s.get("h", "")), 14,
                     P("primary"), bold=True, align=("center", "bottom"), lh=1.4))
        e.append(txt(f"t{i}", cx - seg / 2 + 6, y + 22, seg - 12, 90, str(s.get("t", "")), 12,
                     P("ink"), align=("center", "top"), lh=1.6))
    if pg.get("note"):
        e.append(txt("note", MARGIN, 408, 848, 30, pg["note"], 12, P("gray"), lh=1.6))
    return page(e, "content", pg.get("notes", ""))


def p_pipeline(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    steps = (pg.get("steps") or [])[:5]
    n = max(1, len(steps))
    gap = 26
    bw = (848 - gap * (n - 1)) / n
    bh = 96
    avail = 452 - body
    y = body + max(0, (avail - bh - 90) / 2)
    for i, s in enumerate(steps):
        x = MARGIN + i * (bw + gap)
        e.append(rect(f"box{i}", x, y, bw, bh, P("soft"), radius=5))
        e.append(rect(f"boxtop{i}", x, y, bw, 3, P("primary")))
        e.append(txt(f"h{i}", x + 10, y + 14, bw - 20, 24, str(s.get("h", "")), 13.5, P("primary"),
                     bold=True, lh=1.35))
        e.append(txt(f"t{i}", x + 10, y + 42, bw - 20, 46, str(s.get("t", "")), 11, P("ink"), lh=1.5))
        if i < n - 1:
            e.append(line(f"ar{i}", x + bw + 4, y + bh / 2, gap - 8, 1, P("accent"), 1.4, arrow=True))
    if pg.get("caption"):
        e.append(txt("cap", MARGIN, y + bh + 16, 848, 18, pg["caption"], SZ["cap"], P("gray"),
                     align=("center", "top")))
    if pg.get("note"):
        e.append(txt("note", MARGIN, y + bh + 42, 848, 44, pg["note"], 12, P("gray"), lh=1.7))
    return page(e, "content", pg.get("notes", ""))


def p_art(pg, total, num):
    """装饰图页：生成/提供的装饰图铺满页面，标题放在预留白区。"""
    src = pg.get("img") or pg.get("fig")
    e = []
    if not src:
        WARNINGS.warn("art 页没有给 img，退化为普通文字页")
        return p_text(pg, total, num)
    path = "media/" + os.path.basename(str(src))
    exists = os.path.isfile(os.path.join(OUT_MEDIA, os.path.basename(str(src)))) or \
        os.path.isfile(str(src))
    if exists:
        e.append(img("artbg", 0, 0, W, H, path, fit_mode="cover"))
    else:
        e.append(rect("artbg-missing", 0, 0, W, H, P("soft")))
        WARNINGS.warn("art 页图片不存在：%s" % src)
    style = pg.get("overlay") or "panel"
    if style == "cover":
        # 全图 + 底部渐深色条（用实色条近似）
        e.append(rect("artband", 0, 396, W, 144, P("deep")))
        e.append(txt("title", 56, 418, 720, 60, pg.get("title", ""), 30, "#FFFFFF", bold=True, lh=1.3,
                     motion="fade"))
        if pg.get("lead"):
            e.append(txt("lead", 56, 482, 780, 28, pg["lead"], 13, mix("#FFFFFF", P("deep"), 0.25)))
    else:
        panel_w = min(640, max(420, int(W * 0.55)))
        e.append(rect("artpanel", 0, 0, panel_w, H, "#FFFFFF"))
        e.append(rect("artline", panel_w - 3, 0, 3, H, P("primary")))
        if pg.get("eyebrow"):
            e.append(txt("eyebrow", 56, 82, panel_w - 110, 20, pg["eyebrow"], 11, P("accent"), ls=1.4))
        e.append(txt("title", 56, 118, panel_w - 120, 140, pg.get("title", ""), 30, P("primary"),
                     bold=True, lh=1.3, motion="fade"))
        e.append(rect("artrule", 56, 276, 88, 3, P("accent")))
        if pg.get("lead"):
            e.append(txt("lead", 56, 298, panel_w - 130, 90, pg["lead"], 15, P("gray"), lh=1.7))
        if pg.get("items"):
            yy = 410
            for i, it in enumerate(pg["items"][:3]):
                e.append(rect(f"dot{i}", 56, yy + 7, 7, 7, P("accent")))
                e.append(txt(f"it{i}", 74, yy, panel_w - 150, 26,
                             str(it.get("k") or it.get("t") or it), 12, P("ink"), lh=1.5))
                yy += 32
    if pg.get("caption"):
        e.append(txt("cap", 56, 512, 848, 14, pg["caption"], 8.5,
                     "#FFFFFF" if style == "cover" else P("lightgray"),
                     align=("left", "top")))
    return page(e, "content", pg.get("notes", ""), background="#FFFFFF")


def p_quote(pg, total, num):
    e = chrome(pg, num, total)
    e.append(rect("qbg", 90, 140, 780, 240, P("soft"), radius=6))
    e.append(txt("qmark", 104, 150, 80, 80, "“", 56, P("accent"), bold=True))
    e.append(txt("qtext", 140, 190, 680, 130, pg.get("text", ""), 23,
                 P("deep") if SKIN == "modern" else P("primary"),
                 bold=True, align=("center", "middle"), lh=1.55, motion="fade"))
    if pg.get("from"):
        e.append(txt("qfrom", 120, 338, 720, 22, "—— " + pg["from"], 12, P("gray"),
                     align=("center", "top")))
    return page(e, "content", pg.get("notes", ""))


def p_summary(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    y = _lead_strip(e, pg, body)
    cols = (pg.get("cols") or [])[:2]
    n = max(1, len(cols))
    gap = 24
    cw = (848 - gap * (n - 1)) / n
    for ci, col in enumerate(cols):
        x = MARGIN + ci * (cw + gap)
        e.append(txt(f"ch{ci}", x, y, cw, 26, col.get("h", ""), SZ["sub"], P("primary"), bold=True))

        yy = y + 48
        for i, ln in enumerate((col.get("items") or [])[:5]):
            e.append(rect(f"dot{ci}_{i}", x + 2, yy + 9, 7, 7, P("accent")))
            e.append(txt(f"it{ci}_{i}", x + 18, yy, cw - 22, 44, str(ln), 13.5, P("ink"), lh=1.65))
            yy += 58
    return page(e, "content", pg.get("notes", ""))


def p_roadmap(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    items = (pg.get("items") or [])[:4]
    y = body + 10
    for i, it in enumerate(items):
        yy = y + i * 74
        e.append(rect(f"when{i}", MARGIN, yy, 110, 30, P("primary"), radius=3))
        e.append(txt(f"wt{i}", MARGIN, yy + 3, 110, 24, str(it.get("when", "")), 12, "#FFFFFF",
                     bold=True, align=("center", "middle")))
        e.append(txt(f"rt{i}", 190, yy + 1, 670, 30, str(it.get("t", "")), 15, P("ink"), lh=1.5))

    return page(e, "content", pg.get("notes", ""))


def p_qa(pg, total, num):
    e = chrome(pg, num, total)
    body = skin_metrics()["body_top"]
    items = (pg.get("items") or [])[:4]
    y = body
    for i, it in enumerate(items):
        yy = y + i * 86
        e.append(txt(f"q{i}", MARGIN, yy, 848, 24, "Q：" + str(it.get("q", "")), 13.5,
                     P("primary"), bold=True, lh=1.4))
        e.append(txt(f"a{i}", MARGIN + 20, yy + 26, 828, 44, "A：" + str(it.get("a", "")), 12,
                     P("ink"), lh=1.6))

    return page(e, "content", pg.get("notes", ""))


def p_thanks(pg, total, num):
    title = pg.get("title") or "谢谢"
    subtitle = pg.get("subtitle") or "请老师和同学批评指正"
    dark = SKIN in ("campus", "bold")
    bg = P("deep") if SKIN == "bold" else (P("primary") if SKIN == "campus" else "#FFFFFF")
    fg = "#FFFFFF" if dark else P("primary")
    e = []
    if dark:
        e.append(rect("bg", 0, 0, W, H, bg))
    if CREST_REL:
        if dark:
            e += _crest_in_plate(400, 92, 160, 110, plate_color="#FFFFFF", pad=8, radius=4)
        elif _crest_aspect() >= 2.2:
            c = _crest_in(330, 96, 300, 76)
            if c:
                e.append(c)
        else:
            c = _crest_in(420, 96, 120, 112)
            if c:
                e.append(c)
    e.append(txt("t1", 0, 232, W, 50, title, 34, fg, bold=True, align=("center", "middle"),
                 motion="fade"))
    e.append(txt("t2", 0, 288, W, 24, subtitle, 15, mix(fg, bg, 0.25) if dark else P("gray"),
                 align=("center", "top")))

    if pg.get("citation"):
        e.append(txt("cite", 120, 440, 720, 30, pg["citation"], 8,
                     mix(fg, bg, 0.5) if dark else P("lightgray"), align=("center", "top"), lh=1.5))
    e += footer(pg, num, total, show_num=False, dark=dark, bg=bg if dark else None)
    return page(e, "final", pg.get("notes", ""), background=bg)


def p_unknown(pg, total, num):
    WARNINGS.error("未知页型：" + str(pg.get("type")))
    return None


MC = {
    "cover": p_cover,
    "paper_card": p_paper_card,
    "paper_info": p_paper_card,
    "toc": p_toc,
    "table_of_contents": p_toc,
    "chapter": p_chapter,
    "bullets": p_bullets,
    "keynums": p_keynums,
    "text": p_text,
    "left_fig": p_left_fig,
    "big_fig": p_big_fig,
    "two_figs": p_two_figs,
    "fig_bullets": p_left_fig,
    "table": p_table,
    "formula": p_formula,
    "compare": p_compare,
    "timeline": p_timeline,
    "pipeline": p_pipeline,
    "art": p_art,
    "quote": p_quote,
    "summary": p_summary,
    "roadmap": p_roadmap,
    "qa": p_qa,
    "thanks": p_thanks,
    "final": p_thanks,
}

# ------------------------------------------------------------------ 自动页面
NUM_RE = re.compile(r"^\s*(\d+)(?:[.．](\d+))?\s*[　\s:：-]*(.*)$")


def page_numbering(pg):
    eb = str(pg.get("eyebrow") or "").strip()
    ti = str(pg.get("title") or "").strip()
    if pg.get("type") == "chapter":
        num = str(pg.get("num") or pg.get("chapter") or "").strip()
        if num:
            return num, num, ti
    for raw in (eb, ti):
        m = NUM_RE.match(raw)
        if m:
            major = m.group(1)
            sub = m.group(2)
            rest = m.group(3).strip()
            return major, (f"{major}.{sub}" if sub else major), rest or ti
    return None, None, ti


def infer_sections(content, body):
    explicit = []
    for s in content.get("sections") or []:
        explicit.append({"num": str(s.get("num", "")), "title": s.get("title", ""),
                         "summary": s.get("summary", "")})
    groups = []
    for idx, pg in enumerate(body):
        if pg.get("type") in ("thanks", "final"):
            continue
        major, no, sub = page_numbering(pg)
        if not major:
            continue
        if not groups or groups[-1]["num"] != major:
            title = next((x["title"] for x in explicit if x["num"] == major), "")
            if not title:
                title = sub if not re.match(r"^\d", sub or "") else "第 %s 节" % major
            groups.append({"num": major, "title": title, "items": []})
        groups[-1]["items"].append({"idx": idx, "sub": pg.get("title", "") or sub, "no": no})
    for g in groups:
        t = next((x["title"] for x in explicit if x["num"] == g["num"]), "")
        if t:
            g["title"] = t
    return groups


def auto_cover(content, profile):
    paper = content.get("paper") or {}
    title = content.get("title") or paper.get("title_cn") or paper.get("title_en") or "汇报"
    subtitle = content.get("subtitle") or ""
    citation = content.get("citation") or ""
    if not citation and paper:
        parts = []
        if paper.get("authors"):
            parts.append(paper["authors"])
        if paper.get("title_en"):
            parts.append("“" + paper["title_en"] + "”")
        if paper.get("venue"):
            parts.append(paper["venue"])
        citation = ", ".join(parts)
    return {"type": "cover", "title": title, "subtitle": subtitle, "citation": citation,
            "notes": content.get("cover_notes") or content.get("meta", {}).get("cover_notes", "")
            or "各位老师、同学好，我汇报的题目是《%s》。" % title}


def auto_toc(content, body, body_start):
    groups = infer_sections(content, body)
    exp = {str(s.get("num", "")): s for s in (content.get("sections") or [])}
    items = []
    for g in groups:
        shorts = []
        for it in g.get("items", [])[:4]:
            if it.get("no") and it.get("sub") and "." in str(it.get("no")):
                shorts.append(f"{it['no']} {it['sub']}")
        subs = "　".join(shorts)
        if not subs:
            subs = str((exp.get(g["num"]) or {}).get("summary") or "")
        if len(subs) > 72:
            subs = subs[:70] + "…"
        a = body_start + g["items"][0]["idx"]
        b = body_start + g["items"][-1]["idx"]
        rng = "P%02d–P%02d" % (a, b) if b > a else "P%02d" % a
        items.append({"num": g["num"], "title": g["title"], "subs": subs, "pages": rng})
    return {"type": "toc", "title": "目录", "items": items,
            "notes": content.get("toc_notes") or "报告分 %d 节展开：%s。" % (
                len(items), "、".join(x["title"] for x in items))}


def auto_thanks(content):
    return {"type": "thanks", "title": "谢谢", "subtitle": "请老师和同学批评指正",
            "notes": content.get("thanks_notes") or "我的汇报到这里，谢谢各位老师和同学，欢迎提问和批评。"}


# ------------------------------------------------------------------ 文案卫生
INTERNAL_NOTE_PAT = re.compile(
    r"请(?:再)?核对|引用前|待确认|待补充|待补|待定|TODO|FIXME|TBD|占位|草稿|自检|核对清单|勿外传|"
    r"(?:本页|此处)(?:待|需|要)(?:补|改|确认)"
)


def _split_zh_sentences(text):
    parts = re.split(r"(?<=[。；！？])", str(text or ""))
    return [x for x in parts if x.strip()]


def sanitize_visible_text(text, sink=None, where=""):
    """把「工作备注/内部提示」从观众可见文案里剔出来，转存到讲稿备注。"""
    raw = str(text or "")
    if not raw or not INTERNAL_NOTE_PAT.search(raw):
        return raw
    kept, moved = [], []
    for sent in _split_zh_sentences(raw):
        if INTERNAL_NOTE_PAT.search(sent):
            moved.append(sent.strip())
        else:
            kept.append(sent)
    if not moved:
        return raw
    if sink is not None:
        note = " ".join(moved)
        sink[0] = (sink[0] + " " + note).strip()
    WARNINGS.warn("从页面文案里移除了写手备注（%s）：%s" % (where or "未标注", " ".join(moved)[:60]))
    return "".join(kept).strip()


def sanitize_content(content):
    """扫描 content.json 的观众可见字段，把内部备注挪到 notes；标题/导语只告警不改。"""
    paper = content.get("paper") or {}
    if isinstance(paper, dict):
        sink = [str(content.get("paper_notes") or "")]
        for key in ("source_note", "metrics_note", "volume", "venue", "citation"):
            if isinstance(paper.get(key), str) and INTERNAL_NOTE_PAT.search(paper[key]):
                paper[key] = sanitize_visible_text(paper[key], sink, "paper.%s" % key)
        content["paper_notes"] = sink[0]
    for i, pg in enumerate(content.get("pages") or []):
        if not isinstance(pg, dict):
            continue
        sink = [str(pg.get("notes") or "")]
        touched = False
        for key in ("source", "tnote", "caption", "caption2", "note", "tag", "lead"):
            if isinstance(pg.get(key), str) and INTERNAL_NOTE_PAT.search(pg[key]):
                pg[key] = sanitize_visible_text(pg[key], sink, "第 %d 页 %s" % (i + 1, key))
                touched = True
        for key in ("items", "blocks", "bullets", "lines", "steps", "rows", "cols", "formulas", "figs"):
            val = pg.get(key)
            if isinstance(val, list):
                for item in val:
                    if isinstance(item, dict):
                        for k2 in ("t", "p", "k", "h", "caption", "file", "text", "q", "a", "when"):
                            if isinstance(item.get(k2), str) and INTERNAL_NOTE_PAT.search(item[k2]):
                                item[k2] = sanitize_visible_text(item[k2], sink, "第 %d 页 %s.%s" % (i + 1, key, k2))
                                touched = True
                    elif isinstance(item, str) and INTERNAL_NOTE_PAT.search(item):
                        new = sanitize_visible_text(item, sink, "第 %d 页 %s" % (i + 1, key))
                        val[val.index(item)] = new
                        touched = True
        if touched:
            pg["notes"] = sink[0]


# ------------------------------------------------------------------ 内容检查
def _plain(s):
    return re.sub(r"<[^>]+>", "", str(s or ""))


def validate_content(content):
    pages = content.get("pages") or []
    if not pages:
        WARNINGS.error("content.pages 为空")
    for i, pg in enumerate(pages, 1):
        t = pg.get("type")
        if t not in MC:
            WARNINGS.error("第 %d 个正文页页型未知：%s" % (i, t))
            continue
        if t in ("cover", "toc", "thanks"):
            continue
        title = str(pg.get("title") or "")
        if title and len(_plain(title)) > 26:
            WARNINGS.warn("第 %d 页标题偏长（%d 字）：%s" % (i, len(title), title))
        for key in ("items", "bullets"):
            for j, it in enumerate(pg.get(key) or []):
                if isinstance(it, str):
                    if len(_plain(it)) > 66:
                        WARNINGS.warn("第 %d 页 %s[%d] 文本过长（%d 字）" % (i, key, j, len(it)))
                    continue
                if not isinstance(it, dict):
                    continue
                k = str(it.get("k") or it.get("h") or "")
                tt = str(it.get("t") or it.get("p") or "")
                if len(_plain(k)) > 14:
                    WARNINGS.warn("第 %d 页 %s[%d] 核心词过长（%d 字）：%s" % (i, key, j, len(k), k))
                if len(_plain(tt)) > 66:
                    WARNINGS.warn("第 %d 页 %s[%d] 说明过长（%d 字）" % (i, key, j, len(tt)))
        notes = str(pg.get("notes") or "")
        if len(_plain(notes)) > 240:
            WARNINGS.warn("第 %d 页讲稿备注偏长（%d 字）" % (i, len(notes)))
        for fkey in ("fig", "fig2"):
            f = pg.get(fkey)
            if f and not os.path.isfile(os.path.join(MEDIA_DIR, os.path.basename(str(f)))) \
                    and not os.path.isfile(str(f)):
                WARNINGS.warn("第 %d 页图片不存在（已跳过该图）：%s" % (i, f))
        for f in pg.get("figs") or []:
            if isinstance(f, dict) and f.get("file"):
                if not os.path.isfile(os.path.join(MEDIA_DIR, os.path.basename(str(f["file"])))) \
                        and not os.path.isfile(str(f["file"])):
                    WARNINGS.warn("第 %d 页图片不存在（已跳过该图）：%s" % (i, f["file"]))

    # 整篇节奏：标题/导语不要「整齐得不像人」
    body = [p for p in pages if isinstance(p, dict) and p.get("type") not in ("cover", "toc", "thanks")]
    if len(body) >= 6:
        strong_re = re.compile(
            r"^(?:[一二三四五六七八九十两]\s*[个条句种步块处点件篇])"
            r"|(?:[一二三四五六七八九十两]\s*[个条句种步块处点件篇][^，。；]{0,10}[，、][^，。；]{0,6}"
            r"[一二三四五六七八九十两]\s*[个条句种步块处点件篇])")
        known = re.compile(r"各缺一块|各管一段|各去掉一块|正好是下一步|一组数字|三句话|三处可迁移|"
                           r"四个缺口|三条结论|两个期望|六条约束|五个对比|几个要素")
        titles = [_plain(p.get("title")) for p in body if _plain(p.get("title"))]
        leads = [_plain(p.get("lead")) for p in body if _plain(p.get("lead"))]
        para = [t for t in titles if strong_re.search(t) or known.search(t)] + \
            [l for l in leads if strong_re.search(l) or known.search(l)]
        if len(para) >= 4:
            WARNINGS.warn("数字对仗标题/导语偏多（%d 处）：%s；同类句式全篇最多留 2–3 处"
                          % (len(para), "；".join(para[:4])))
        if leads and len(leads) / len(body) > 0.6:
            WARNINGS.warn("导语覆盖率偏高（%d/%d 页）；建议只给真正需要一句结论的页面加 lead"
                          % (len(leads), len(body)))
        tlen = [len(t) for t in titles]
        if len(tlen) >= 8:
            sd = (sum((x - sum(tlen) / len(tlen)) ** 2 for x in tlen) / len(tlen)) ** 0.5
            if sd < 3.5:
                WARNINGS.warn("标题长度太整齐（标准差 %.1f）；让长短错开一些" % sd)

        # 页面骨架：与 humanize_lint.check_deck_structure 同源的轻量检查（一键管线也能拦）
        stop_char = set("的了在与和及或上个")

        def _shape_s(text):
            out = []
            for ch in text:
                if "\u4e00" <= ch <= "\u9fff":
                    out.append("C")
                elif ch.isdigit():
                    out.append("9")
                elif ch.isascii() and ch.isalpha():
                    out.append("X")
                elif ch.strip():
                    out.append(ch)
            return "".join(out)

        def _labels_of(pg):
            labels = []
            for key in ("items", "bullets", "blocks", "steps"):
                for it in pg.get(key) or []:
                    if isinstance(it, dict):
                        v = _plain(it.get("k") or it.get("h"))
                        if v and len(v) <= 8:
                            labels.append(v)
            return labels

        def _visible_text(pg):
            parts = [_plain(pg.get("title")), _plain(pg.get("lead")), _plain(pg.get("sub"))]
            for key in ("items", "bullets"):
                for it in pg.get(key) or []:
                    if isinstance(it, dict):
                        parts += [_plain(it.get("k")), _plain(it.get("t") or it.get("p"))]
                    else:
                        parts.append(_plain(it))
            for key in ("blocks", "steps", "lines"):
                val = pg.get(key)
                if isinstance(val, list):
                    for it in val:
                        if isinstance(it, dict):
                            parts += [_plain(it.get("h")), _plain(it.get("p") or it.get("t"))]
                        else:
                            parts.append(_plain(it))
                elif isinstance(val, str):
                    parts.append(_plain(val))
            parts.append(_plain(pg.get("takeaway")))
            return " ".join(x for x in parts if x)

        def _parts_of(pg):
            out = [_plain(pg.get("title")), _plain(pg.get("lead")), _plain(pg.get("sub"))]
            for key in ("items", "bullets"):
                for it in pg.get(key) or []:
                    if isinstance(it, dict):
                        out += [_plain(it.get("k")), _plain(it.get("t") or it.get("p"))]
                    else:
                        out.append(_plain(it))
            for key in ("lines", "blocks", "steps", "takeaway", "note", "tnote"):
                val = pg.get(key)
                if isinstance(val, list):
                    for it in val:
                        if isinstance(it, dict):
                            out += [_plain(it.get("h")), _plain(it.get("p") or it.get("t"))]
                        else:
                            out.append(_plain(it))
                elif isinstance(val, str):
                    out.append(_plain(val))
            return out

        num_re = re.compile(r"[一二三四五六七八九十两0-9]+\s*[个条句种步块处点类项篇次组张页所台套]")
        clause_pages = {}
        n_takeaway = 0
        for i, pg in enumerate(body, 1):
            labels = _labels_of(pg)
            if len(labels) >= 4:
                cnt = {}
                for v in labels:
                    s = _shape_s(v)
                    cnt[s] = cnt.get(s, 0) + 1
                top, num = max(cnt.items(), key=lambda kv: kv[1])
                if num >= 4:
                    WARNINGS.warn("第 %d 页同构短标签 %d 个（形状 %s）；标签长短/句式要错开" % (i, num, top))
            if len(labels) >= 3:
                # 编号式标签（贡献一/步骤2/方案③）共享词干是正常写法，先排除；
                # 纯 ASCII 前缀（MODE/MODE-w/MODE-o）是方法名，不算
                enum_re = re.compile(r"[一二三四五六七八九十\d①②③④⑤⑥⑦⑧⑨⑩]$")
                plain_labels = [v for v in labels if not enum_re.search(v)]
                big_cnt = {}
                for v in plain_labels:
                    for bg in {v[j:j + 2] for j in range(len(v) - 1)
                               if any("\u4e00" <= c <= "\u9fff" for c in v[j:j + 2])}:
                        big_cnt[bg] = big_cnt.get(bg, 0) + 1
                bg2, num2 = max(big_cnt.items(), key=lambda kv: kv[1]) if big_cnt else ("", 0)
                if num2 >= 3:
                    WARNINGS.warn("第 %d 页 %d 个标签共用词缀「%s」；换掉一半的写法" % (i, num2, bg2))
                char_cnt = {}
                for v in plain_labels:
                    for c in {c for c in v if "\u4e00" <= c <= "\u9fff" and c not in stop_char}:
                        char_cnt[c] = char_cnt.get(c, 0) + 1
                c3, num3 = max(char_cnt.items(), key=lambda kv: kv[1]) if char_cnt else ("", 0)
                if num3 >= 4:
                    WARNINGS.warn("第 %d 页 %d 个标签都围着「%s」字（同句式套出来）；保留一处，其余改具体判断"
                                  % (i, num3, c3))
            heads = [_plain(pg.get("title")), _plain(pg.get("lead")), _plain(pg.get("sub"))] + labels
            n_num = sum(len(num_re.findall(x)) for x in heads if x)
            if n_num >= 3:
                WARNINGS.warn("第 %d 页计数词 %d 个；「三个/五种」式表述一页最多留 2 个" % (i, n_num))
            if _plain(pg.get("takeaway")):
                n_takeaway += 1
            # 逐字段分句：整页拼接后 split 会把全页粘成一个巨型子句，跨页重复永远查不出来
            for part in _parts_of(pg):
                for clause in re.split(r"[，。；、：\n]", part):
                    clause = re.sub(r"\s+", "", clause)
                    if len(clause) >= 8:
                        clause_pages.setdefault(clause, set()).add(i)
            items = [it for it in (pg.get("items") or pg.get("bullets") or []) if it]
            lead = _plain(pg.get("lead"))
            if lead and len(items) >= 2:
                item_txt = " ".join(
                    (_plain(it.get("k")) + _plain(it.get("t") or it.get("p")))
                    if isinstance(it, dict) else _plain(it) for it in items)
                if len(item_txt) >= 10:
                    stop2 = {"我们", "可以", "这个", "一个", "的是", "不是", "就是", "以及", "同时", "因此"}
                    lb = {lead[j:j + 2] for j in range(len(lead) - 1)} - stop2
                    ib = {item_txt[j:j + 2] for j in range(len(item_txt) - 1)}
                    if len(lb & ib) >= 5:
                        WARNINGS.warn("第 %d 页导语在复述要点名词；导语要么给结论/数字，要么删" % i)
            texts = [(_plain(it.get("k")) + " " + _plain(it.get("t") or it.get("p")))
                     if isinstance(it, dict) else _plain(it) for it in items]
            if len(texts) >= 3:
                bgs = [{s[j:j + 2] for j in range(len(s) - 1)} for s in texts if len(s) >= 4]
                worst, pair = 0.0, ""
                for a in range(len(bgs)):
                    for b in range(a + 1, len(bgs)):
                        union = len(bgs[a] | bgs[b]) or 1
                        jac = len(bgs[a] & bgs[b]) / union
                        if jac > worst:
                            worst, pair = jac, "%s / %s" % (texts[a][:16], texts[b][:16])
                if worst >= 0.55:
                    WARNINGS.warn("第 %d 页两条要点高度相似（%.2f）：%s；合并或删一条" % (i, worst, pair))
        if n_takeaway >= 4:
            WARNINGS.warn("%d 页都带收束句（takeaway）；只给重点页收 2–3 页，其余删掉" % n_takeaway)
        repeats = sorted((c for c, ps in clause_pages.items() if len(ps) >= 2), key=len, reverse=True)
        if repeats:
            WARNINGS.warn("同一句出现在多页（%d 处）：%s；金句全篇只说一次" % (len(repeats), repeats[0][:40]))
        suf_cnt = {}
        for t in titles:
            m = re.search(r"(的[\u4e00-\u9fff]{1,3})$", t)
            if m:
                suf_cnt[m.group(1)] = suf_cnt.get(m.group(1), 0) + 1
        for sfx, c in sorted(suf_cnt.items(), key=lambda kv: -kv[1]):
            if c >= 3:
                WARNINGS.warn("「…%s」式标题 %d 个；同后缀标题全篇最多 2 个" % (sfx, c))
        chaps = [_plain(p.get("sub")) for p in body
                 if p.get("type") in ("chapter", "section") and _plain(p.get("sub"))]
        tmpl = [s for s in chaps if s.count("→") >= 2 or strong_re.search(s) or known.search(s)]
        if len(tmpl) >= 2:
            WARNINGS.warn("%d 个过渡页副题是同一模板（箭头链/数字串）；几章的导语别都写成一种格式" % len(tmpl))


# ------------------------------------------------------------------ 输出
OUT_MEDIA = ""
MEDIA_DIR = ""
CREST_REL = ""
CREST_KIND = "none"
PROFILE = {}


def profile_get(key, default=None):
    return PROFILE.get(key, default)


def copy_media(files, media_dir, out_media):
    os.makedirs(out_media, exist_ok=True)
    copied = []
    for f in files:
        if not f:
            continue
        src = f if os.path.isfile(str(f)) else os.path.join(media_dir, os.path.basename(str(f)))
        if not os.path.isfile(src):
            WARNINGS.warn("图片找不到（已跳过）：%s" % f)
            continue
        dst = os.path.join(out_media, os.path.basename(src))
        try:
            if os.path.abspath(src) != os.path.abspath(dst):
                shutil.copy(src, dst)
            copied.append(dst)
        except Exception as e:  # noqa: BLE001
            WARNINGS.error("图片复制失败：%s（%s）" % (src, e))
    return copied


def apply_motion(pg, ptype, motion_mode):
    """把本页登记的动画目标转成 PPTD animations。"""
    targets = [t for t in MOTION_TARGETS if t[0]]
    if not targets:
        return []
    motion_mode = (pg.get("motion") or motion_mode or "gentle").lower()
    if motion_mode == "none":
        return []
    out = []
    effect_map = {"zoom": "zoomIn", "fade": "fadeIn"}
    for i, (eid, kind) in enumerate(targets[:6]):
        if motion_mode == "lively":
            eff = "flyIn" if kind == "fade" else effect_map.get(kind, "fadeIn")
            item = {"elementId": eid, "effect": eff, "durationMs": 450}
            if eff == "flyIn":
                item["direction"] = "bottom"
        else:
            eff = effect_map.get(kind, "fadeIn")
            item = {"elementId": eid, "effect": eff, "durationMs": 450}
        if i == 0:
            item["trigger"] = "withPrevious"
        else:
            item["trigger"] = "afterPrevious"
        out.append(item)
    return out


def dump_page(path, pg):
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(pg, f, allow_unicode=True, sort_keys=False, width=4000)


def main(argv=None):
    global BRAND, SKIN, CN, EN, MOTTO_FONT, OUT_MEDIA, MEDIA_DIR, CREST_REL, CREST_KIND, PROFILE
    ap = argparse.ArgumentParser(description="学院汇报 PPT 生成器 v2")
    ap.add_argument("--content", help="content.json")
    ap.add_argument("--profile", default="", help="profile.json（向导输出）")
    ap.add_argument("--out", default="", help="PPTD 工程目录（需已由 ppt_create 建好）")
    ap.add_argument("--school", default="", help="学校名（覆盖 profile）")
    ap.add_argument("--college", default="", help="学院名（覆盖 profile）")
    ap.add_argument("--reporter", default="", help="汇报人（覆盖 profile）")
    ap.add_argument("--date", default="", help="封面日期")
    ap.add_argument("--crest", default="", help="校徽图片路径（覆盖模板库）")
    ap.add_argument("--logo", default="", help="院徽/实验室 logo 路径")
    ap.add_argument("--media-dir", default="", help="论文图目录")
    ap.add_argument("--skin", default="", choices=["", "auto", "academic", "campus", "modern", "bold"])
    ap.add_argument("--motion", default="", choices=["", "none", "gentle", "lively"])
    ap.add_argument("--cn-font", default="Microsoft YaHei")
    ap.add_argument("--en-font", default="Arial")
    ap.add_argument("--list-schools", nargs="?", const="", default=None, help="列模板库学校")
    ap.add_argument("--school-info", default="", help="查一所学校的模板匹配情况")
    ap.add_argument("--check-only", action="store_true", help="只检查内容，不写工程")
    args = ap.parse_args(argv)

    skill_dir = _skill_dir()

    if args.list_schools is not None:
        db = load_school_db(skill_dir)
        packs = load_brand_packs(skill_dir)
        kw = args.list_schools.strip().lower()
        rows = [e for e in db if not kw or kw in str(e.get("name", "")).lower()
                or kw in str(e.get("pinyin", "")).lower() or kw in str(e.get("abbr", "")).lower()]
        print("模板库共 %d 所（带品牌包 %d 所）" % (len(db), len(packs)))
        for e in rows[:80]:
            mark = "★" if e.get("slug") in packs else " "
            print("%s %-6s %-20s %-6s %s" % (mark, e.get("slug", ""), e.get("name", ""),
                                             e.get("province", ""), "/".join(e.get("level") or []) or "本科"))
        if len(rows) > 80:
            print("……共 %d 条，只显示前 80 条，可加关键字过滤" % len(rows))
        return 0

    if args.school_info:
        db = load_school_db(skill_dir)
        packs = load_brand_packs(skill_dir)
        entry, score = resolve_school(db, args.school_info)
        if not entry:
            print("没有匹配到：%s" % args.school_info)
            return 2
        print("匹配：%s（slug=%s，得分 %s）" % (entry.get("name"), entry.get("slug"), score))
        bp = packs.get(entry.get("slug"), {})
        if bp:
            print(json.dumps(bp, ensure_ascii=False, indent=2))
        else:
            print("没有品牌包，将用中性配色 + 文字标识")
        return 0

    if not args.content:
        ap.error("必须有 --content（或用 --list-schools / --school-info）")

    content = read_json(args.content)
    PROFILE = {}
    if args.profile:
        PROFILE.update(read_json(args.profile))
    if isinstance(content.get("profile"), dict):
        tmp = dict(content["profile"])
        tmp.update(PROFILE)
        PROFILE = tmp
    for k, v in (("school", args.school), ("college", args.college),
                 ("reporter", args.reporter), ("date", args.date)):
        if v:
            PROFILE[k] = v
    if args.crest:
        PROFILE["crest_override"] = args.crest
    if args.logo:
        PROFILE["logo_override"] = args.logo
    if args.skin:
        PROFILE["skin"] = args.skin
    if args.motion:
        PROFILE["motion"] = args.motion
    PROFILE.setdefault("role", "研究生")
    purpose = (PROFILE.get("purpose") or "group_paper").strip()
    defaults = PURPOSE_DEFAULTS.get(purpose, PURPOSE_DEFAULTS["group_paper"])
    SKIN = (PROFILE.get("skin") or "auto").strip()
    if SKIN in ("", "auto"):
        SKIN = defaults["skin"]
    if SKIN not in SKIN_NAMES:
        WARNINGS.warn("未知皮肤 %s，回落到 academic" % SKIN)
        SKIN = "academic"
    motion_mode = (PROFILE.get("motion") or "gentle").strip()
    if motion_mode not in ("none", "gentle", "lively"):
        motion_mode = "gentle"
    CN = args.cn_font or "Microsoft YaHei"
    EN = args.en_font or "Arial"

    brand, entry, bp = build_brand(skill_dir, PROFILE, content)
    BRAND.update(brand)

    media_dir = os.path.abspath(args.media_dir) if args.media_dir else ""
    MEDIA_DIR = media_dir
    if not media_dir:
        WARNINGS.warn("没有给 --media-dir，页面里的论文图将找不到")
    if not os.path.isdir(media_dir):
        if media_dir:
            WARNINGS.error("图片目录不存在：%s" % media_dir)
        MEDIA_DIR = ""

    sanitize_content(content)

    v1_style = "schema_version" not in content
    if v1_style:
        WARNINGS.warn("content.json 没有 schema_version:2，按 v1 兼容模式读取（建议按 format_spec v2 重写）")

    validate_content(content) if MEDIA_DIR else None
    if args.check_only:
        print("检查完成：%d 条问题" % len(WARNINGS.items))
        return 1 if any(x[0] == "ERROR" for x in WARNINGS.items) else 0

    out = os.path.abspath(args.out)
    if not out:
        ap.error("必须有 --out")
    os.makedirs(os.path.join(out, "pages"), exist_ok=True)
    os.makedirs(os.path.join(out, "media"), exist_ok=True)
    OUT_MEDIA = os.path.join(out, "media")

    crest_path, crest_kind = resolve_crest(skill_dir, brand, PROFILE, args, OUT_MEDIA)
    CREST_REL = "media/crest.png" if crest_path else ""
    CREST_KIND = crest_kind
    logo_path = ""
    logo_candidates = [PROFILE.get("logo_override") or "", getattr(args, "logo", "") or ""]
    for c in logo_candidates:
        if c and os.path.isfile(c):
            logo_path = os.path.join(OUT_MEDIA, "logo.png")
            shutil.copy(c, logo_path)
            break

    # 组织页面
    raw_pages = content.get("pages") or []
    has_cover = any(p.get("type") == "cover" for p in raw_pages)
    has_toc = any(p.get("type") in ("toc", "table_of_contents") for p in raw_pages)
    body = [p for p in raw_pages if p.get("type") not in ("cover", "toc", "table_of_contents")]
    paper = content.get("paper") or {}
    has_paper_card = any(p.get("type") in ("paper_card", "paper_info") for p in body)
    body = [p for p in body if p.get("type") not in ("paper_card", "paper_info")]

    cover = next((p for p in raw_pages if p.get("type") == "cover"), None) or auto_cover(content, PROFILE)
    cover["_reporter"] = PROFILE.get("reporter") or ""
    cover["_date"] = PROFILE.get("date") or args.date or ("%d 年 %d 月" % (date.today().year, date.today().month))
    cover["_college"] = PROFILE.get("college") or BRAND.get("college_default") or ""
    if not cover.get("title"):
        cover["title"] = content.get("title") or paper.get("title_cn") or "汇报"

    pages = [cover]
    if paper.get("title_cn") or paper.get("title_en") or has_paper_card:
        pages.append({"type": "paper_card", "paper": paper, "title": "论文简介",
                      "notes": content.get("paper_notes") or ""})
    toc_slot = None
    if content.get("sections") and len(body) >= 6 and not has_toc:
        toc_slot = len(pages)
        pages.append(None)
    pages += body
    if not has_toc and toc_slot is not None:
        body_start = toc_slot + 2
        pages[toc_slot] = auto_toc(content, body, body_start)
    elif has_toc:
        user_toc = next(p for p in raw_pages if p.get("type") in ("toc", "table_of_contents"))
        if toc_slot is None and not any(p.get("type") in ("toc",) for p in pages):
            pass
    # 用户自定义 toc 插到封面/论文简介之后
    if has_toc and not any(p and p.get("type") in ("toc", "table_of_contents") for p in pages):
        user_toc = next(p for p in raw_pages if p.get("type") in ("toc", "table_of_contents"))
        pages.insert(len(pages) - len(body), user_toc)
    # 结尾
    if content.get("meta", {}).get("auto_thanks", True) and not any(
            p and p.get("type") in ("thanks", "final") for p in pages):
        pages.append(auto_thanks(content))

    total = len(pages)

    # 图片收集
    figs = []
    for pg in pages:
        if not pg:
            continue
        for k in ("fig", "fig2", "img"):
            if pg.get(k):
                figs.append(pg[k])
        for f in pg.get("formulas") or []:
            if isinstance(f, dict) and f.get("img"):
                figs.append(f["img"])
        for f in pg.get("figs") or []:
            if isinstance(f, dict) and f.get("file"):
                figs.append(f["file"])
    copy_media(figs, MEDIA_DIR, OUT_MEDIA)

    # deck.pptd
    deck = {
        "profile": "generate",
        "title": cover.get("title") or content.get("title") or "",
        "size": [W, H],
        "export": {"fontEmbed": "subset",
                   "transition": "auto" if motion_mode == "lively" else "fade",
                   "animations": "auto" if motion_mode != "none" else "off"},
        "pages": [f"pages/{i+1:02d}_{(p or {}).get('type','unknown')}.page" for i, p in enumerate(pages)],
    }
    dump_page(os.path.join(out, "deck.pptd"), deck)

    written = []
    for i, pg in enumerate(pages):
        if pg is None:
            continue
        ptype = pg.get("type")
        builder = MC.get(ptype)
        if builder is None:
            WARNINGS.error("跳过未知页型：%s" % ptype)
            continue
        MOTION_TARGETS.clear()
        built = builder(pg, total, i + 1)
        if built is None:
            continue
        anims = apply_motion(pg, ptype, motion_mode)
        if anims:
            built["animations"] = anims
        name = f"{i+1:02d}_{ptype}.page"
        dump_page(os.path.join(out, "pages", name), built)
        written.append(name)

    report = {
        "built_at": date.today().isoformat(),
        "pages": len(written),
        "skin": SKIN,
        "motion": motion_mode,
        "purpose": purpose,
        "profile": PROFILE,
        "brand": {k: BRAND.get(k) for k in ("school", "school_en", "slug", "primary", "accent",
                                            "deep", "soft", "motto", "footer_left", "footer_right",
                                            "verified", "crest_kind")},
        "crest": crest_path or "",
        "logo": logo_path or "",
        "warnings": WARNINGS.items,
    }
    write_json(os.path.join(out, "build_report.json"), report)
    print(f"[build_deck] 完成：{len(written)} 页 → {out}（皮肤 {SKIN}，动画 {motion_mode}，"
          f"学校 {BRAND.get('school') or '未指定'}）")
    if any(x[0] == "ERROR" for x in WARNINGS.items):
        print("[build_deck] 有错误，请先修复；详见 build_report.json", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""profile_wizard.py — 学院汇报 · 模板匹配向导（本地网页，纯标准库）。

开始做 PPT 时收集：学校 / 学院 / 姓名 / 身份 / 汇报目的 / 时长 / 日期 / 皮肤 /
动画档位 / 「说人话」开关 / 校徽院徽，输出严格符合 references/profile_schema.md
的 profile.json。

两种用法：

1. 网页模式（推荐，人用）：

       python3 profile_wizard.py --skill-dir <技能根> --out profile.json [--port 0] [--no-open] [--timeout 900]

   起在 127.0.0.1 上（默认随机端口），自动用系统浏览器打开；提交后写 profile.json，
   服务 2 秒后自动退出。页面完全离线：无 CDN、无第三方字体，全部内嵌。

2. 纯文本模式（agent 用）：

       python3 profile_wizard.py --skill-dir <技能根> --out profile.json --plain
       python3 profile_wizard.py --skill-dir <技能根> --out profile.json --answers answers.json

   `--plain` 把问题逐条打印到 stdout（可以读给用户听）；`--answers` 直接按 JSON
   生成 profile（agent 在聊天里问完用户后写 answers.json 最省事）。

容错：
- schools.json 不存在 → 用内置 15 所样例并在页面顶部提示；
- brand_packs.json 缺某校 → 用中性色（#1F3A5F）；
- 指定端口被占用 → 自动顺延重试，最后退回系统随机端口；
- 浏览器打不开 → 打印 URL，手动打开即可。

隐私：只监听 127.0.0.1，不上传任何数据；上传的校徽/院徽写到
`<out 目录>/../user_assets/`，profile.json 里只记本机绝对路径。
"""

from __future__ import annotations

import argparse
import base64
import binascii
import datetime
import hashlib
import http.server
import json
import os
import re
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
from typing import Any, Dict, List, Optional, Tuple

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

# 取色函数复用 brand_from_logo.py；缺文件时降级（不影响写 profile）
try:
    from brand_from_logo import extract_palette as _extract_palette  # type: ignore
except Exception:  # noqa: BLE001 - 任何 import 失败都降级
    _extract_palette = None  # type: ignore[assignment]

try:
    from PIL import Image  # type: ignore
except Exception:  # noqa: BLE001
    Image = None  # type: ignore[assignment]


SCHEMA_VERSION = 2
WIZARD_TITLE = "学院汇报 · 模板匹配向导"

# 中性配色：没有品牌包时的降级色（brand_schema.md §3）
NEUTRAL_COLORS = {
    "primary": "#1F3A5F",
    "accent": "#3A6EA5",
    "deep": "#172B47",
    "soft": "#EDF1F7",
    "ink": "#1A1A1A",
}

ROLES = ["本科生", "研究生", "博士生", "教师", "其他"]

# 6 个汇报目的：卡片文案 + 默认时长 / 页数 / 皮肤 / 听众（对齐 profile_schema.md）
PURPOSES = [
    {
        "key": "course_report",
        "name": "个人课程汇报",
        "desc": "选题背景 → 做了什么 → 结果与收获，讲给课程班同学和老师。",
        "default_min": 9,
        "default_pages": 12,
        "skin": "campus",
        "audience": "课程班",
    },
    {
        "key": "group_paper",
        "name": "组会文献汇报",
        "desc": "论文简介 → 背景 → 方法 → 实验 → 结论与启发，讲给课题组。",
        "default_min": 22,
        "default_pages": 18,
        "skin": "academic",
        "audience": "课题组组会",
    },
    {
        "key": "group_progress",
        "name": "组会进展汇报",
        "desc": "上次结论 → 本期工作 → 结果 → 卡点 → 下一步，讲进度。",
        "default_min": 12,
        "default_pages": 10,
        "skin": "modern",
        "audience": "课题组组会",
    },
    {
        "key": "proposal",
        "name": "开题 / 中期检查",
        "desc": "背景与意义 → 现状不足 → 研究内容 → 技术路线 → 计划。",
        "default_min": 18,
        "default_pages": 15,
        "skin": "academic",
        "audience": "开题 / 中期专家组",
    },
    {
        "key": "defense",
        "name": "毕业答辩",
        "desc": "背景 → 方法 → 实验 → 结论 → 创新点与展望，讲完整工作。",
        "default_min": 26,
        "default_pages": 22,
        "skin": "academic",
        "audience": "答辩委员会",
    },
    {
        "key": "project",
        "name": "项目 / 竞赛汇报",
        "desc": "痛点 → 方案 → 亮点 → 数据 → 团队与计划，讲清价值和完成度。",
        "default_min": 15,
        "default_pages": 15,
        "skin": "bold",
        "audience": "项目评审 / 竞赛评委",
    },
]
PURPOSE_BY_KEY = {p["key"]: p for p in PURPOSES}

SKINS = [
    {"key": "auto", "name": "智能匹配", "desc": "按汇报目的自动选（推荐）"},
    {"key": "academic", "name": "学术 · 院系格式", "desc": "白底细线、编号标题；组会 / 开题 / 答辩"},
    {"key": "campus", "name": "校园 · 课程汇报", "desc": "封面大色块、品牌竖条；课程展示"},
    {"key": "modern", "name": "现代 · 组会进展", "desc": "无框眉标、细线条目；进展汇报"},
    {"key": "bold", "name": "撞色 · 项目竞赛", "desc": "深底大色块、粗色条；项目 / 竞赛"},
]
SKIN_KEYS = {"auto", "academic", "campus", "modern", "bold"}

MOTIONS = [
    {"key": "none", "name": "无动画", "desc": "只用 fade 转场，适合打印 / 发送版"},
    {"key": "gentle", "name": "轻淡入", "desc": "fade 转场 + 要点依次淡入（推荐）"},
    {"key": "lively", "name": "活泼", "desc": "混合转场 + 飞入，适合答辩 / 竞赛"},
]
MOTION_KEYS = {"none", "gentle", "lively"}

# 内置样例（schools.json 缺失时兜底）。字段按 brand_schema.md §1；配色只对示例大学
# 有实测标准色，其余走中性色，不编造。
BUILTIN_SCHOOLS = [
    {"slug": "neu", "name": "示例大学", "name_en": "Northeastern University",
     "aliases": ["东大", "NEU"], "province": "示例省", "city": "示例市",
     "level": ["985", "211", "双一流"], "type": "理工",
     "pinyin": "dongbeidaxue", "abbr": "dbdx"},
    {"slug": "tsinghua", "name": "清华大学", "name_en": "Tsinghua University",
     "aliases": ["清华"], "province": "北京", "city": "北京",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "qinghuadaxue", "abbr": "qhdx"},
    {"slug": "pku", "name": "北京大学", "name_en": "Peking University",
     "aliases": ["北大"], "province": "北京", "city": "北京",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "beijingdaxue", "abbr": "bjdx"},
    {"slug": "zju", "name": "浙江大学", "name_en": "Zhejiang University",
     "aliases": ["浙大"], "province": "浙江", "city": "杭州",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "zhejiangdaxue", "abbr": "zjdx"},
    {"slug": "sjtu", "name": "上海交通大学", "name_en": "Shanghai Jiao Tong University",
     "aliases": ["上海交大", "上交"], "province": "上海", "city": "上海",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "shanghaijiaotongdaxue", "abbr": "shjtdx"},
    {"slug": "fudan", "name": "复旦大学", "name_en": "Fudan University",
     "aliases": ["复旦"], "province": "上海", "city": "上海",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "fudandaxue", "abbr": "fddx"},
    {"slug": "nju", "name": "南京大学", "name_en": "Nanjing University",
     "aliases": ["南大"], "province": "江苏", "city": "南京",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "nanjingdaxue", "abbr": "njdx"},
    {"slug": "ustc", "name": "中国科学技术大学", "name_en": "University of Science and Technology of China",
     "aliases": ["中科大", "科大"], "province": "安徽", "city": "合肥",
     "level": ["985", "211", "双一流"], "type": "理工",
     "pinyin": "zhongguokexuejishudaxue", "abbr": "zgkxjsdx"},
    {"slug": "hit", "name": "哈尔滨工业大学", "name_en": "Harbin Institute of Technology",
     "aliases": ["哈工大"], "province": "黑龙江", "city": "哈尔滨",
     "level": ["985", "211", "双一流"], "type": "理工",
     "pinyin": "haerbingongyedaxue", "abbr": "hebgydx"},
    {"slug": "xjtu", "name": "西安交通大学", "name_en": "Xi'an Jiaotong University",
     "aliases": ["西安交大"], "province": "陕西", "city": "西安",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "xianjiaotongdaxue", "abbr": "xajtdx"},
    {"slug": "whu", "name": "武汉大学", "name_en": "Wuhan University",
     "aliases": ["武大"], "province": "湖北", "city": "武汉",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "wuhandaxue", "abbr": "whdx"},
    {"slug": "hust", "name": "华中科技大学", "name_en": "Huazhong University of Science and Technology",
     "aliases": ["华科", "华中科大"], "province": "湖北", "city": "武汉",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "huazhongkejidaxue", "abbr": "hzkjdx"},
    {"slug": "sysu", "name": "中山大学", "name_en": "Sun Yat-sen University",
     "aliases": ["中大"], "province": "广东", "city": "广州",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "zhongshandaxue", "abbr": "zsdx"},
    {"slug": "scu", "name": "四川大学", "name_en": "Sichuan University",
     "aliases": ["川大"], "province": "四川", "city": "成都",
     "level": ["985", "211", "双一流"], "type": "综合",
     "pinyin": "sichuandaxue", "abbr": "scdx"},
    {"slug": "uestc", "name": "电子科技大学", "name_en": "University of Electronic Science and Technology of China",
     "aliases": ["电子科大", "成电"], "province": "四川", "city": "成都",
     "level": ["985", "211", "双一流"], "type": "理工",
     "pinyin": "dianzikejidaxue", "abbr": "dzkjdx"},
]

# 内置样例里唯一有实测标准色的学校（来自品牌包契约示例）
BUILTIN_BRANDS = {
    "neu": {
        "school": "示例大学",
        "school_en": "Northeastern University",
        "colors": {
            "primary": "#00367F",
            "accent": "#0055AA",
            "deep": "#002B5C",
            "soft": "#EAF1FA",
            "ink": "#1A1A1A",
            "gold": "#C9A063",
        },
        "motto": "自强不息 知行合一",
        "footer_left": "MAKE GLOBAL CHANGES",
        "footer_right": "Sydney Smart Technology College",
        "skin_hint": "academic",
        "college_default": "示例学院",
    }
}

HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
SLUG_RE = re.compile(r"^[a-z0-9-]+$")
MIME_EXT = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/webp": "webp",
    "image/gif": "gif",
    "image/svg+xml": "svg",
}
EXT_MIME = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "webp": "image/webp",
    "gif": "image/gif",
    "svg": "image/svg+xml",
}
MAX_BODY_BYTES = 12 * 1024 * 1024  # 12 MB，含 base64 图片
MAX_ASSET_BYTES = 6 * 1024 * 1024


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------


def now_local() -> datetime.datetime:
    return datetime.datetime.now().astimezone()


def default_date_text() -> str:
    now = now_local()
    return "%d 年 %d 月" % (now.year, now.month)


def log(message: str = "") -> None:
    print(message, flush=True)


def warn(message: str) -> None:
    sys.stderr.write(message + "\n")
    sys.stderr.flush()


def normalize_hex(value: Any) -> Optional[str]:
    if not isinstance(value, str):
        return None
    text = value.strip().upper()
    if not text:
        return None
    if not text.startswith("#"):
        text = "#" + text
    if HEX_RE.match(text):
        return text
    return None


def normalize_colors(raw: Any) -> Dict[str, str]:
    """把品牌包 colors 归一成完整配色；缺字段用主色推。"""
    colors = dict(NEUTRAL_COLORS)
    if not isinstance(raw, dict):
        return colors
    primary = normalize_hex(raw.get("primary"))
    if not primary:
        return colors
    colors["primary"] = primary
    accent = normalize_hex(raw.get("accent")) or colors["accent"]
    colors["accent"] = accent
    # deep：主色压暗 25%；soft：主色提亮 92%
    def _mix(hex_a: str, hex_b: str, t: float) -> str:
        a = tuple(int(hex_a[i : i + 2], 16) for i in (1, 3, 5))
        b = tuple(int(hex_b[i : i + 2], 16) for i in (1, 3, 5))
        out = tuple(int(round(a[k] + (b[k] - a[k]) * t)) for k in range(3))
        return "#%02X%02X%02X" % out

    colors["deep"] = normalize_hex(raw.get("deep")) or _mix(primary, "#000000", 0.25)
    colors["soft"] = normalize_hex(raw.get("soft")) or _mix(primary, "#FFFFFF", 0.92)
    colors["ink"] = normalize_hex(raw.get("ink")) or "#1A1A1A"
    gold = normalize_hex(raw.get("gold"))
    if gold:
        colors["gold"] = gold
    return colors


def _brand_owns(brand: Any, school_name: str) -> bool:
    """品牌包是否属于这所学校（按包内 school 校名判定；包内没写校名时视为可接受）。"""
    if not isinstance(brand, dict):
        return False
    pack_school = str(brand.get("school") or "").strip()
    return not pack_school or pack_school == school_name


def _as_str_list(value: Any) -> List[str]:
    """把 ["985","211"] / "985,211" / "985 211" 统一成字符串列表（容错其它数据源）。"""
    if value is None:
        return []
    if isinstance(value, str):
        parts = re.split(r"[,，、/|\s]+", value)
        return [part.strip() for part in parts if part.strip()]
    if isinstance(value, (list, tuple, set)):
        return [str(item).strip() for item in value if str(item).strip()]
    return []


def level_rank(levels: Any) -> int:
    levels = levels if isinstance(levels, list) else []
    for rank, name in enumerate(("985", "211", "双一流", "部属", "省属")):
        if name in levels:
            return rank
    return 9 if levels else 6


def esc(text: Any) -> str:
    """给 HTML 用。"""
    s = "" if text is None else str(text)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# ---------------------------------------------------------------------------
# 学校 / 品牌数据加载
# ---------------------------------------------------------------------------


class Catalog:
    def __init__(self, skill_dir: str) -> None:
        self.skill_dir = skill_dir
        self.schools_path = os.path.join(skill_dir, "assets", "schools", "schools.json")
        self.brands_path = os.path.join(skill_dir, "assets", "schools", "brand_packs.json")
        self.crests_dir = os.path.join(skill_dir, "assets", "schools", "crests")
        self.schools: List[dict] = []
        self.brands: Dict[str, dict] = {}
        self.school_source = "builtin"
        self.school_source_path = ""
        self.fallback = False
        self.fallback_reason = ""
        self.crest_map: Dict[str, Tuple[str, str]] = {}  # slug -> (绝对路径, mime)
        self.by_slug: Dict[str, dict] = {}
        self.load()

    # -- 加载 --------------------------------------------------------------
    def load(self) -> None:
        raw_schools: List[dict] = []
        if os.path.isfile(self.schools_path):
            try:
                with open(self.schools_path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                if isinstance(data, dict):
                    data = data.get("schools") or data.get("items") or data.get("data") or []
                if isinstance(data, list):
                    raw_schools = [
                        item
                        for item in data
                        if isinstance(item, dict) and item.get("slug") and item.get("name")
                    ]
                if raw_schools:
                    self.school_source = "file"
                    self.school_source_path = self.schools_path
            except (OSError, ValueError) as exc:
                self.fallback_reason = "schools.json 读取失败（%s）" % exc
        else:
            self.fallback_reason = "未找到 schools.json"

        if not raw_schools:
            raw_schools = [dict(item) for item in BUILTIN_SCHOOLS]
            self.school_source = "builtin"
            self.fallback = True
            self.fallback_reason = self.fallback_reason or "schools.json 不存在"

        if os.path.isfile(self.brands_path):
            try:
                with open(self.brands_path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                if isinstance(data, dict):
                    if isinstance(data.get("brands"), dict):
                        data = data["brands"]
                    self.brands = {str(k): v for k, v in data.items() if isinstance(v, dict)}
            except (OSError, ValueError) as exc:
                warn("品牌包读取失败（%s），配色退回中性色。" % exc)

        builtin_brand = BUILTIN_BRANDS if self.fallback else {}
        # 品牌包的键不一定是 slug：school-brand 生产的数据用拼音全拼作键，
        # 契约示例用短 slug。这里做多路匹配（slug → 拼音 → 首字母 → 校名）。
        brand_by_name: Dict[str, dict] = {}
        for brand_value in self.brands.values():
            brand_school = str(brand_value.get("school") or "").strip()
            if brand_school:
                brand_by_name.setdefault(brand_school, brand_value)
        for raw in raw_schools:
            slug = str(raw.get("slug", "")).strip().lower()
            if not slug or not SLUG_RE.match(slug):
                continue
            name = str(raw.get("name", "")).strip()
            if not name:
                continue
            pinyin = str(raw.get("pinyin") or "").strip().lower().replace(" ", "")
            abbr = str(raw.get("abbr") or "").strip().lower().replace(" ", "")
            # 品牌包匹配：slug 直接命中；拼音/首字母兜底时必须确认品牌包属于本校
            # （拼音有重码：山西师范大学/陕西师范大学、济南大学/暨南大学 都是同拼，
            #   只按拼音查会把别校品牌包配错）。
            brand = self.brands.get(slug)
            if brand is None and pinyin:
                candidate = self.brands.get(pinyin)
                if candidate is not None and _brand_owns(candidate, name):
                    brand = candidate
            if brand is None and abbr:
                candidate = self.brands.get(abbr)
                if candidate is not None and _brand_owns(candidate, name):
                    brand = candidate
            if brand is None:
                brand = brand_by_name.get(name)
            if brand is None:
                brand = builtin_brand.get(slug)
            brand = brand or {}
            colors = normalize_colors(brand.get("colors"))
            entry = {
                "slug": slug,
                "name": name,
                "name_en": str(raw.get("name_en") or brand.get("school_en") or "").strip(),
                "aliases": _as_str_list(raw.get("aliases")),
                "province": str(raw.get("province") or "").strip(),
                "city": str(raw.get("city") or "").strip(),
                "level": _as_str_list(raw.get("level")),
                "type": str(raw.get("type") or "其他").strip() or "其他",
                "pinyin": pinyin,
                "abbr": abbr,
                "colors": colors,
                "motto": str(brand.get("motto") or "").strip(),
                "footer_left": str(brand.get("footer_left") or "").strip(),
                "footer_right": str(brand.get("footer_right") or "").strip(),
                "skin_hint": str(brand.get("skin_hint") or "").strip(),
                "college_default": str(brand.get("college_default") or "").strip(),
                "has_brand": bool(brand),
                "brand": brand,
            }
            self.schools.append(entry)

        if not self.schools:
            for item in BUILTIN_SCHOOLS:
                item = dict(item)
                item["colors"] = normalize_colors(BUILTIN_BRANDS.get(item["slug"], {}).get("colors"))
                item.setdefault("has_brand", False)
                self.schools.append(item)
            self.school_source = "builtin"
            self.fallback = True

        # 排序：知名层次优先，然后按校名；页面上默认先看到这批
        self.schools.sort(key=lambda s: (level_rank(s["level"]), s["name"]))
        self.by_slug = {s["slug"]: s for s in self.schools}

        # 校徽文件：品牌包指向 + crests/<slug>.* + <slug>-text.*
        crest_candidates: Dict[str, str] = {}
        crest_base = os.path.normpath(os.path.join(self.skill_dir, "assets", "schools"))
        for school in self.schools:
            rel = (school.get("brand") or {}).get("crest")
            if isinstance(rel, str) and rel.strip():
                full = os.path.normpath(os.path.join(crest_base, rel))
                if full.startswith(crest_base) and os.path.isfile(full):
                    crest_candidates[school["slug"]] = full
        # 拼音/首字母重码时（山西师范/陕西师范、济南大学/暨南大学 同拼），
        # 不能拿拼音文件名当自己的校徽，否则会串图。只允许唯一拼音/首字母做兜底。
        pinyin_counts: Dict[str, int] = {}
        abbr_counts: Dict[str, int] = {}
        for school in self.schools:
            if school.get("pinyin"):
                pinyin_counts[school["pinyin"]] = pinyin_counts.get(school["pinyin"], 0) + 1
            if school.get("abbr"):
                abbr_counts[school["abbr"]] = abbr_counts.get(school["abbr"], 0) + 1
        for slug, school in self.by_slug.items():
            if slug in crest_candidates:
                continue
            # 文件名优先 slug（唯一），其次唯一拼音全拼，最后唯一拼音首字母
            stems = [slug]
            pinyin = school.get("pinyin")
            if pinyin and pinyin not in stems and pinyin_counts.get(pinyin) == 1:
                stems.append(pinyin)
            abbr = school.get("abbr")
            if abbr and abbr not in stems and abbr_counts.get(abbr) == 1:
                stems.append(abbr)
            for ext in ("png", "svg", "jpg", "jpeg", "webp"):
                for stem in stems:
                    for pattern in ("%s.%s" % (stem, ext), "%s-text.%s" % (stem, ext)):
                        full = os.path.join(self.crests_dir, pattern)
                        if os.path.isfile(full):
                            crest_candidates[slug] = full
                            break
                    if slug in crest_candidates:
                        break
                if slug in crest_candidates:
                    break
        for slug, full in crest_candidates.items():
            ext = os.path.splitext(full)[1].lower().lstrip(".")
            self.crest_map[slug] = (os.path.abspath(full), EXT_MIME.get(ext, "image/png"))
        for school in self.schools:
            school["has_crest"] = school["slug"] in self.crest_map
        self.matched_brand_count = sum(1 for school in self.schools if school["has_brand"])

        self.page_data = {
            "schools": [
                {
                    "slug": s["slug"],
                    "name": s["name"],
                    "name_en": s["name_en"],
                    "aliases": s["aliases"],
                    "province": s["province"],
                    "city": s["city"],
                    "level": s["level"],
                    "type": s["type"],
                    "pinyin": s["pinyin"],
                    "abbr": s["abbr"],
                    "colors": {k: s["colors"][k] for k in ("primary", "accent", "deep", "soft")},
                    "crest_url": "/api/crest/%s" % s["slug"] if s["has_crest"] else "",
                    "motto": s["motto"],
                    "college_default": s["college_default"],
                    "has_brand": bool(s["has_brand"]),
                }
                for s in self.schools
            ],
            "purposes": PURPOSES,
            "roles": ROLES,
            "skins": SKINS,
            "motions": MOTIONS,
            "meta": {
                "title": WIZARD_TITLE,
                "school_count": len(self.schools),
                "brand_count": getattr(self, "matched_brand_count", 0),
                "school_source": self.school_source,
                "fallback": self.fallback,
                "fallback_reason": self.fallback_reason,
                "default_date": default_date_text(),
                "neutral_colors": dict(NEUTRAL_COLORS),
                "schema_version": SCHEMA_VERSION,
            },
        }

    # -- 查询 --------------------------------------------------------------
    def search(self, query: str, limit: int = 60) -> List[dict]:
        return search_schools(self.schools, query, limit=limit)

    def resolve(self, name: str = "", slug: str = "") -> Optional[dict]:
        slug = (slug or "").strip().lower()
        if slug and slug in self.by_slug:
            return self.by_slug[slug]
        name = (name or "").strip()
        if not name:
            return None
        lowered = name.lower()
        for school in self.schools:
            if school["name"] == name:
                return school
        for school in self.schools:
            pool = [school["name"].lower(), school["pinyin"], school["abbr"]]
            pool += [a.lower() for a in school["aliases"]]
            if lowered in pool:
                return school
        hits = self.search(name, limit=1)
        return hits[0] if hits else None


def search_schools(schools: List[dict], query: str, limit: int = 60) -> List[dict]:
    """按 名称 / 拼音 / 首字母 / 别名 / 省份 搜索，返回按相关度排序的结果。"""
    query = (query or "").strip().lower()
    if not query:
        return schools[:limit]
    tokens = [t for t in re.split(r"\s+", query) if t]
    scored: List[Tuple[int, int, str, dict]] = []
    for school in schools:
        haystack = " ".join(
            [
                school["name"],
                school.get("name_en", ""),
                school.get("province", ""),
                school.get("city", ""),
                school.get("pinyin", ""),
                school.get("abbr", ""),
            ]
            + list(school.get("aliases") or [])
        ).lower()
        if not all(token in haystack for token in tokens):
            continue
        score = _search_score(school, query)
        scored.append((score, level_rank(school["level"]), school["name"], school))
    scored.sort(key=lambda item: (item[0], item[1], item[2]))
    return [item[3] for item in scored[:limit]]


def _search_score(school: dict, query: str) -> int:
    name = school["name"].lower()
    aliases = [a.lower() for a in school.get("aliases") or []]
    abbr = school.get("abbr", "")
    pinyin = school.get("pinyin", "")
    if name == query:
        return 0
    if name.startswith(query):
        return 1
    if query in name:
        return 2
    if any(alias == query for alias in aliases):
        return 3
    if any(query in alias for alias in aliases):
        return 4
    if abbr and abbr == query:
        return 5
    if pinyin.startswith(query):
        return 6
    if query in pinyin:
        return 7
    if school.get("abbr", "").startswith(query):
        return 8
    if query in school.get("province", ""):
        return 9
    if query in school.get("name_en", "").lower():
        return 10
    return 20


# ---------------------------------------------------------------------------
# 素材（校徽 / 院徽）落盘 + 取色
# ---------------------------------------------------------------------------


def user_assets_dir_for(out_path: str) -> str:
    """约定：`<out 目录>/../user_assets/`。"""
    out_dir = os.path.dirname(os.path.abspath(out_path))
    return os.path.abspath(os.path.join(out_dir, os.pardir, "user_assets"))


def sniff_ext(raw: bytes) -> Optional[str]:
    if raw.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if raw.startswith(b"\xff\xd8\xff"):
        return "jpg"
    if raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
        return "webp"
    if raw.startswith(b"GIF87a") or raw.startswith(b"GIF89a"):
        return "gif"
    head = raw[:4096].lstrip()
    if head.startswith(b"<") and (b"<svg" in head.lower() or b"<?xml" in head):
        return "svg"
    return None


def decode_data_url(data_url: Any) -> Tuple[bytes, str]:
    """data:image/png;base64,... → (bytes, mime)。"""
    if not isinstance(data_url, str) or not data_url.startswith("data:"):
        raise ValueError("图片数据格式不对")
    header, _, payload = data_url.partition(",")
    if not payload:
        raise ValueError("图片数据为空")
    mime = header[5:].split(";")[0].strip().lower() or "application/octet-stream"
    try:
        raw = base64.b64decode(payload, validate=True)
    except (ValueError, binascii.Error):
        raise ValueError("图片 base64 解码失败")
    if not raw:
        raise ValueError("图片数据为空")
    if len(raw) > MAX_ASSET_BYTES:
        raise ValueError("图片超过 %d MB，请压缩后再上传" % (MAX_ASSET_BYTES // (1024 * 1024)))
    return raw, mime


def palette_from_bytes(raw: bytes, ext: str) -> Tuple[Optional[Dict[str, str]], str]:
    """返回 (palette 或 None, 提示语)。"""
    if ext == "svg":
        return None, "SVG 无法自动取色（图片仍会原样保存）"
    if _extract_palette is None:
        return None, "缺少 brand_from_logo.py，跳过自动取色"
    try:
        palette = _extract_palette(raw)
    except Exception as exc:  # noqa: BLE001 - 取色失败不阻断流程
        return None, "自动取色失败：%s" % exc
    return palette, ""


def ingest_asset(user_assets_dir: str, kind: str, slug: str, raw: bytes) -> Dict[str, Any]:
    """把一个上传图片写进 user_assets，返回信息（含取色）。"""
    ext = sniff_ext(raw)
    if ext is None:
        raise ValueError("只支持 PNG / SVG / JPG / WebP / GIF 图片")
    digest = hashlib.sha1(raw).hexdigest()[:8]
    safe_slug = slug if SLUG_RE.match(slug or "") else "school"
    os.makedirs(user_assets_dir, exist_ok=True)
    filename = "%s_%s_%s.%s" % (kind, safe_slug, digest, ext)
    path = os.path.join(user_assets_dir, filename)
    tmp = path + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(raw)
    os.replace(tmp, path)
    palette, note = palette_from_bytes(raw, ext)
    return {
        "path": os.path.abspath(path),
        "filename": filename,
        "ext": ext,
        "palette": palette,
        "note": note,
        "size": len(raw),
    }


# ---------------------------------------------------------------------------
# profile 组装
# ---------------------------------------------------------------------------


def normalize_purpose(value: Any) -> Optional[dict]:
    if value is None:
        return None
    if isinstance(value, int):
        if 1 <= value <= len(PURPOSES):
            return PURPOSES[value - 1]
        return None
    text = str(value).strip()
    if not text:
        return None
    for purpose in PURPOSES:
        if text == purpose["key"] or text == purpose["name"]:
            return purpose
    for index, purpose in enumerate(PURPOSES, 1):
        if text == str(index):
            return purpose
    for purpose in PURPOSES:
        if text and (text in purpose["name"] or purpose["name"] in text):
            return purpose
    return None


def normalize_role(value: Any, default: str = "研究生") -> str:
    if value is None:
        return default
    text = str(value).strip()
    if text in ROLES:
        return text
    mapping = {
        "本科": "本科生", "本科学生": "本科生",
        "硕士": "研究生", "硕士生": "研究生", "研究生": "研究生",
        "博士": "博士生", "博士研究生": "博士生",
        "老师": "教师", "教师": "教师",
        "其他": "其他",
    }
    return mapping.get(text, default)


def _as_int(value: Any) -> Optional[int]:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def compose_profile(
    catalog: Catalog,
    state: Dict[str, Any],
    out_path: str,
) -> Tuple[dict, List[str], List[dict], str]:
    """按 profile_schema.md 组装 profile.json。

    state 键（网页 / answers 共用）：
      school_name, school_slug, college, college_en, reporter, role, purpose,
      duration_min, date, skin, motion, humanize, audience, motto,
      footer_left, footer_right, lang, crest_data_url / crest_path,
      logo_data_url / logo_path, brand_colors

    返回 (profile, warnings, saved_assets, user_assets_dir)。
    """
    warnings: List[str] = []
    profile: Dict[str, Any] = {}

    school_name = str(state.get("school_name") or "").strip()
    school = catalog.resolve(school_name, str(state.get("school_slug") or ""))
    if school is None:
        if not school_name:
            warnings.append("没有填学校，生成器会用中性配色。")
        else:
            warnings.append(
                "没有在模板库里匹配到「%s」，生成器会用中性配色；school_slug 留空。" % school_name
            )
        school_slug = ""
        school_name = school_name or ""
        brand = {}
    else:
        school_slug = school["slug"]
        school_name = school["name"]
        brand = school.get("brand") or BUILTIN_BRANDS.get(school_slug) or {}

    purpose = normalize_purpose(state.get("purpose"))
    if purpose is None:
        purpose = PURPOSES[1]  # 组会文献汇报
        warnings.append("汇报目的没识别，按「组会文献汇报」处理。")

    reporter = str(state.get("reporter") or "").strip()
    if not reporter:
        warnings.append("没有填汇报人姓名，封面会留下划线占位。")

    role = normalize_role(state.get("role"), "研究生")

    duration = _as_int(state.get("duration_min"))
    if duration is None or duration <= 0:
        duration = int(purpose["default_min"])
    duration = max(1, min(300, duration))

    date_text = str(state.get("date") or "").strip() or default_date_text()

    skin = str(state.get("skin") or "auto").strip().lower()
    if skin not in SKIN_KEYS:
        warnings.append("皮肤「%s」不认识，已改为 auto。" % skin)
        skin = "auto"

    motion = str(state.get("motion") or "gentle").strip().lower()
    if motion not in MOTION_KEYS:
        motion = "gentle"

    humanize = state.get("humanize")
    if isinstance(humanize, str):
        humanize = humanize.strip().lower() not in ("0", "false", "no", "off", "否", "关")
    humanize = True if humanize is None else bool(humanize)

    audience = str(state.get("audience") or "").strip() or str(purpose["audience"])
    lang = str(state.get("lang") or "zh").strip() or "zh"

    user_assets_dir = user_assets_dir_for(out_path)
    saved_assets: List[dict] = []
    crest_override = ""
    logo_override = ""
    crest_palette: Optional[Dict[str, str]] = None
    logo_palette: Optional[Dict[str, str]] = None

    # 校徽
    try:
        crest_info = _resolve_uploaded_asset(
            catalog, state, "crest", school_slug, user_assets_dir, warnings
        )
        if crest_info:
            crest_override = crest_info["path"]
            crest_palette = crest_info.get("palette")
            saved_assets.append(crest_info)
    except ValueError as exc:
        warnings.append("校徽没有保存：%s" % exc)
    # 院徽 / 实验室 logo
    try:
        logo_info = _resolve_uploaded_asset(
            catalog, state, "logo", school_slug, user_assets_dir, warnings
        )
        if logo_info:
            logo_override = logo_info["path"]
            logo_palette = logo_info.get("palette")
            saved_assets.append(logo_info)
    except ValueError as exc:
        warnings.append("院徽没有保存：%s" % exc)

    brand_colors = _resolve_brand_colors(
        state.get("brand_colors"), crest_palette, logo_palette, warnings
    )

    profile = {
        "schema_version": SCHEMA_VERSION,
        "school": school_name,
        "school_slug": school_slug,
        "college": str(state.get("college") or "").strip(),
        "college_en": str(state.get("college_en") or "").strip(),
        "reporter": reporter,
        "role": role,
        "purpose": purpose["key"],
        "audience": audience,
        "duration_min": duration,
        "date": date_text,
        "skin": skin,
        "motion": motion,
        "humanize": humanize,
        "footer_left": str(state.get("footer_left") or "").strip(),
        "footer_right": str(state.get("footer_right") or "").strip(),
        "motto": str(state.get("motto") or brand.get("motto") or "").strip(),
        "crest_override": crest_override,
        "logo_override": logo_override,
        "brand_colors": brand_colors,
        "lang": lang,
        "created_at": now_local().isoformat(timespec="seconds"),
    }
    return profile, warnings, saved_assets, user_assets_dir


def _resolve_uploaded_asset(
    catalog: Catalog,
    state: Dict[str, Any],
    kind: str,
    slug: str,
    user_assets_dir: str,
    warnings: List[str],
) -> Optional[dict]:
    data_url = state.get("%s_data_url" % kind)
    path = state.get("%s_path" % kind)
    if data_url:
        raw, _mime = decode_data_url(data_url)
        return ingest_asset(user_assets_dir, kind, slug, raw)
    if path:
        candidate = os.path.abspath(os.path.expanduser(str(path)))
        if not os.path.isfile(candidate):
            warnings.append("%s 文件不存在，已跳过：%s" % ("校徽" if kind == "crest" else "院徽", candidate))
            return None
        with open(candidate, "rb") as fh:
            raw = fh.read()
        if len(raw) > MAX_ASSET_BYTES:
            warnings.append("%s 超过 %d MB，已跳过。" % (kind, MAX_ASSET_BYTES // (1024 * 1024)))
            return None
        return ingest_asset(user_assets_dir, kind, slug, raw)
    return None


def _resolve_brand_colors(
    explicit: Any,
    crest_palette: Optional[Dict[str, str]],
    logo_palette: Optional[Dict[str, str]],
    warnings: List[str],
) -> Optional[Dict[str, str]]:
    if isinstance(explicit, dict):
        primary = normalize_hex(explicit.get("primary"))
        if primary:
            colors = {"primary": primary}
            accent = normalize_hex(explicit.get("accent"))
            if accent:
                colors["accent"] = accent
            for key in ("deep", "soft"):
                value = normalize_hex(explicit.get(key))
                if value:
                    colors[key] = value
            return colors
        warnings.append("brand_colors 里没有合法 primary，已忽略。")
    if crest_palette:
        return dict(crest_palette)
    if logo_palette:
        return dict(logo_palette)
    return None


def write_profile(out_path: str, profile: dict) -> str:
    path = os.path.abspath(out_path)
    out_dir = os.path.dirname(path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(profile, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    os.replace(tmp, path)
    return path


# ---------------------------------------------------------------------------
# HTTP 服务
# ---------------------------------------------------------------------------


class WizardContext:
    def __init__(self, args, catalog: Catalog) -> None:
        self.args = args
        self.catalog = catalog
        self.out_path = args.out
        self.saved = False
        self.profile_path = ""
        self.profile: Dict[str, Any] = {}
        self.saved_assets: List[dict] = []
        self.warnings: List[str] = []
        self.user_assets_dir = user_assets_dir_for(args.out)
        self.last_activity = time.monotonic()
        self.stop_event = threading.Event()
        self.timeout_hit = False
        self.server: Optional[http.server.ThreadingHTTPServer] = None
        self.html = build_html(catalog).encode("utf-8")
        self._shutdown_timer: Optional[threading.Timer] = None
        self.lock = threading.Lock()

    def request_shutdown(self, delay: float = 0.0) -> None:
        if delay > 0:
            self._shutdown_timer = threading.Timer(delay, self.request_shutdown)
            self._shutdown_timer.daemon = True
            self._shutdown_timer.start()
            return
        self.stop_event.set()
        if self.server is not None:
            try:
                self.server.shutdown()
            except Exception:  # noqa: BLE001 - 关服务失败不算致命
                pass


class WizardHandler(http.server.BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "CollegeWizard/2.0"

    # 不让标准库往 stderr 打一堆英文日志；自己在关键点打印中文
    def log_message(self, fmt: str, *args: Any) -> None:  # noqa: D401
        return

    @property
    def ctx(self) -> WizardContext:
        return self.server.ctx  # type: ignore[attr-defined]

    # -- 响应工具 ----------------------------------------------------------
    def _send(self, code: int, body: bytes, content_type: str, cache: str = "no-store") -> None:
        try:
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", cache)
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            if self.command != "HEAD":
                self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, code: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send(code, body, "application/json; charset=utf-8")

    def _error(self, code: int, message: str) -> None:
        self._json(code, {"ok": False, "error": message})

    def _read_json(self) -> Optional[dict]:
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except (TypeError, ValueError):
            self._error(400, "请求长度不合法")
            return None
        if length <= 0:
            self._error(400, "请求体为空")
            return None
        if length > MAX_BODY_BYTES:
            self._error(413, "上传内容太大（上限 %d MB）" % (MAX_BODY_BYTES // (1024 * 1024)))
            return None
        try:
            raw = self.rfile.read(length)
            data = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, ValueError):
            self._error(400, "请求体不是合法 JSON")
            return None
        if not isinstance(data, dict):
            self._error(400, "请求体必须是 JSON 对象")
            return None
        return data

    # -- GET ---------------------------------------------------------------
    def do_GET(self) -> None:  # noqa: N802
        self.ctx.last_activity = time.monotonic()
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        try:
            if path in ("/", "/index.html"):
                self._send(200, self.ctx.html, "text/html; charset=utf-8")
            elif path == "/favicon.ico":
                self._send(204, b"", "image/x-icon")
            elif path == "/api/health":
                self._json(
                    200,
                    {
                        "ok": True,
                        "title": WIZARD_TITLE,
                        "saved": self.ctx.saved,
                        "school_count": len(self.ctx.catalog.schools),
                        "school_source": self.ctx.catalog.school_source,
                        "fallback": self.ctx.catalog.fallback,
                        "profile_path": self.ctx.profile_path,
                    },
                )
            elif path == "/api/schools":
                query = urllib.parse.parse_qs(parsed.query).get("q", [""])[0]
                limit = min(200, max(1, _as_int(urllib.parse.parse_qs(parsed.query).get("limit", ["60"])[0]) or 60))
                hits = self.ctx.catalog.search(query, limit=limit)
                self._json(
                    200,
                    {
                        "ok": True,
                        "q": query,
                        "count": len(hits),
                        "schools": [public_school_json(s) for s in hits],
                    },
                )
            elif path == "/api/brand":
                slug = urllib.parse.parse_qs(parsed.query).get("slug", [""])[0]
                school = self.ctx.catalog.by_slug.get(slug)
                if school is None:
                    self._error(404, "没有这所学校：%s" % slug)
                else:
                    self._json(
                        200,
                        {
                            "ok": True,
                            "slug": slug,
                            "name": school["name"],
                            "colors": school["colors"],
                            "has_brand": school["has_brand"],
                            "motto": school["motto"],
                            "skin_hint": school["skin_hint"],
                        },
                    )
            elif path.startswith("/api/crest/"):
                slug = urllib.parse.unquote(path[len("/api/crest/") :]).strip()
                item = self.ctx.catalog.crest_map.get(slug)
                if item is None:
                    self._error(404, "这所学校没有校徽文件：%s" % slug)
                    return
                full, mime = item
                try:
                    with open(full, "rb") as fh:
                        body = fh.read()
                except OSError:
                    self._error(404, "校徽文件读不到：%s" % full)
                    return
                self._send(200, body, mime, cache="public, max-age=3600")
            else:
                self._error(404, "没有这个地址：%s" % path)
        except Exception as exc:  # noqa: BLE001 - 服务不因单次请求崩掉
            self._error(500, "服务内部错误：%s" % exc)

    # -- POST --------------------------------------------------------------
    def do_POST(self) -> None:  # noqa: N802
        self.ctx.last_activity = time.monotonic()
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        try:
            if path == "/api/palette":
                data = self._read_json()
                if data is None:
                    return
                self._handle_palette(data)
            elif path == "/api/save":
                data = self._read_json()
                if data is None:
                    return
                self._handle_save(data)
            else:
                self._error(404, "没有这个接口：%s" % path)
        except Exception as exc:  # noqa: BLE001
            self._error(500, "服务内部错误：%s" % exc)

    def _handle_palette(self, data: dict) -> None:
        data_url = data.get("data_url")
        if not isinstance(data_url, str) or not data_url.startswith("data:"):
            self._error(400, "缺少图片数据（data_url）")
            return
        try:
            raw, _mime = decode_data_url(data_url)
            ext = sniff_ext(raw)
            if ext is None:
                self._error(400, "只支持 PNG / SVG / JPG / WebP / GIF 图片")
                return
            palette, note = palette_from_bytes(raw, ext)
        except ValueError as exc:
            self._error(400, str(exc))
            return
        if palette is None:
            self._json(200, {"ok": False, "error": note or "无法自动取色", "palette": None})
            return
        self._json(200, {"ok": True, "palette": palette, "note": note})

    def _handle_save(self, data: dict) -> None:
        ctx = self.ctx
        school_name = str(data.get("school") or "").strip()
        reporter = str(data.get("reporter") or "").strip()
        purpose = normalize_purpose(data.get("purpose"))
        errors: List[str] = []
        if not school_name:
            errors.append("请选择学校")
        if not reporter:
            errors.append("请填写汇报人姓名")
        if purpose is None:
            errors.append("请选择汇报目的")
        if errors:
            self._error(400, "；".join(errors))
            return

        state = {
            "school_name": school_name,
            "school_slug": data.get("school_slug") or "",
            "college": data.get("college") or "",
            "college_en": data.get("college_en") or "",
            "reporter": reporter,
            "role": data.get("role") or "研究生",
            "purpose": purpose["key"],
            "duration_min": data.get("duration_min"),
            "date": data.get("date") or "",
            "skin": data.get("skin") or "auto",
            "motion": data.get("motion") or "gentle",
            "humanize": data.get("humanize"),
            "audience": data.get("audience") or "",
            "motto": data.get("motto") or "",
            "footer_left": data.get("footer_left") or "",
            "footer_right": data.get("footer_right") or "",
            "lang": data.get("lang") or "zh",
            "brand_colors": data.get("brand_colors"),
        }
        # 图片：网页传 {name, data_url}
        for kind in ("crest", "logo"):
            item = data.get(kind)
            if isinstance(item, dict) and item.get("data_url"):
                state["%s_data_url" % kind] = item["data_url"]
            elif isinstance(item, str) and item:
                state["%s_path" % kind] = item

        try:
            profile, warnings, saved_assets, user_assets_dir = compose_profile(
                ctx.catalog, state, ctx.out_path
            )
            profile_path = write_profile(ctx.out_path, profile)
        except ValueError as exc:
            self._error(400, "生成 profile 失败：%s" % exc)
            return
        except OSError as exc:
            self._error(500, "写 profile.json 失败：%s" % exc)
            return

        with ctx.lock:
            ctx.saved = True
            ctx.profile_path = profile_path
            ctx.profile = profile
            ctx.saved_assets = saved_assets
            ctx.warnings = warnings

        log("收到提交：%s / %s / %s" % (profile["school"] or "（未匹配）", profile["reporter"] or "（未填）", purpose["name"]))
        for asset in saved_assets:
            label = "校徽" if asset["path"].find("/crest_") >= 0 else "院徽"
            log("已保存%s：%s" % (label, asset["path"]))
            if asset.get("palette"):
                palette = asset["palette"]
                log("  自动取色：primary %s / accent %s" % (palette.get("primary"), palette.get("accent")))
            elif asset.get("note"):
                log("  取色提示：%s" % asset["note"])
        for message in warnings:
            warn("提示：%s" % message)
        log("已写入：%s" % profile_path)
        log("服务 2 秒后退出。")

        self._json(
            200,
            {
                "ok": True,
                "profile_path": profile_path,
                "profile": profile,
                "user_assets_dir": user_assets_dir,
                "saved_assets": [
                    {"path": a["path"], "kind": "crest" if "/crest_" in a["path"].replace("\\", "/") else "logo",
                     "palette": a.get("palette"), "note": a.get("note")}
                    for a in saved_assets
                ],
                "warnings": warnings,
            },
        )
        ctx.request_shutdown(delay=2.0)


def public_school_json(school: dict) -> dict:
    return {
        "slug": school["slug"],
        "name": school["name"],
        "name_en": school["name_en"],
        "aliases": school["aliases"],
        "province": school["province"],
        "city": school["city"],
        "level": school["level"],
        "type": school["type"],
        "pinyin": school["pinyin"],
        "abbr": school["abbr"],
        "colors": school["colors"],
        "crest_url": "/api/crest/%s" % school["slug"] if school.get("has_crest") else "",
        "has_brand": school.get("has_brand", False),
    }


class WizardHTTPServer(http.server.ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, address, handler, ctx: WizardContext) -> None:
        self.ctx = ctx
        super().__init__(address, handler)


def bind_server(preferred_port: int, ctx: WizardContext) -> WizardHTTPServer:
    """优先用指定端口；被占用就顺延最多 10 个，最后退回系统随机端口。"""
    candidates: List[int] = []
    if preferred_port and preferred_port > 0:
        candidates.extend(range(preferred_port, preferred_port + 10))
    candidates.append(0)
    last_error: Optional[OSError] = None
    for port in candidates:
        try:
            server = WizardHTTPServer(("127.0.0.1", port), WizardHandler, ctx)
            if port != preferred_port and preferred_port:
                log("提示：端口 %d 被占用，已改用 %d。" % (preferred_port, server.server_address[1]))
            return server
        except OSError as exc:
            last_error = exc
            continue
    raise OSError("没有可用端口（最后错误：%s）" % last_error)


def open_browser(url: str) -> bool:
    """macOS 用 open，Windows 用 start，Linux 用 xdg-open；失败退回 webbrowser。"""
    try:
        if sys.platform == "darwin":
            subprocess.Popen(
                ["open", url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        elif os.name == "nt":
            subprocess.Popen(
                ["cmd", "/c", "start", "", url],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        else:
            subprocess.Popen(
                ["xdg-open", url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
        return True
    except Exception:  # noqa: BLE001
        try:
            return bool(webbrowser.open(url))
        except Exception:  # noqa: BLE001
            return False


def _watchdog(ctx: WizardContext, timeout: int) -> None:
    while not ctx.stop_event.wait(1.0):
        if time.monotonic() - ctx.last_activity >= timeout:
            ctx.timeout_hit = True
            log("\n等待超时（%d 秒没有新的页面请求），服务已退出，未生成 profile.json。" % timeout)
            ctx.stop_event.set()
            if ctx.server is not None:
                try:
                    ctx.server.shutdown()
                except Exception:  # noqa: BLE001
                    pass
            return


def run_http(args, catalog: Catalog) -> int:
    ctx = WizardContext(args, catalog)
    try:
        server = bind_server(args.port or 0, ctx)
    except OSError as exc:
        warn("启动本地服务失败：%s" % exc)
        return 2
    ctx.server = server
    port = server.server_address[1]
    url = "http://127.0.0.1:%d/" % port

    log("")
    log("=" * 62)
    log("  %s" % WIZARD_TITLE)
    log("=" * 62)
    log("技能目录：%s" % catalog.skill_dir)
    if catalog.fallback:
        log("学校数据：内置 %d 所样例（%s）" % (len(catalog.schools), catalog.fallback_reason))
        log("  → 把完整 schools.json 放到 assets/schools/ 后重启，即可搜全国院校。")
    else:
        log("学校数据：%d 所（%s）" % (len(catalog.schools), catalog.school_source_path))
    log(
        "品牌包：%d 所匹配%s（文件里 %d 条）"
        % (
            getattr(catalog, "matched_brand_count", 0),
            "" if getattr(catalog, "matched_brand_count", 0) else "（缺失，预览用中性色）",
            len(catalog.brands),
        )
    )
    log("输出文件：%s" % os.path.abspath(args.out))
    log("素材目录：%s" % ctx.user_assets_dir)
    log("")
    if args.no_open:
        log("浏览器未自动打开（--no-open）。请手动打开：%s" % url)
    else:
        opened = open_browser(url)
        if opened:
            log("已在浏览器打开：%s" % url)
        else:
            log("自动打开浏览器失败，请手动打开：%s" % url)
    log("")
    log("等待页面提交…（%d 秒无操作自动退出，Ctrl+C 立即退出）" % args.timeout)

    watchdog = threading.Thread(target=_watchdog, args=(ctx, args.timeout), daemon=True)
    watchdog.start()

    try:
        server.serve_forever(poll_interval=0.2)
    except KeyboardInterrupt:
        log("\n收到 Ctrl+C，服务退出。")
        ctx.stop_event.set()
    finally:
        try:
            server.server_close()
        except Exception:  # noqa: BLE001
            pass

    if ctx.saved:
        log("")
        log("完成：profile.json 已生成 → %s" % ctx.profile_path)
        log("服务已退出。")
        return 0
    if ctx.timeout_hit:
        return 2
    return 0


# ---------------------------------------------------------------------------
# 纯文本模式（agent 用）
# ---------------------------------------------------------------------------


def _answers_from_file(path: str) -> dict:
    full = os.path.abspath(os.path.expanduser(path))
    if not os.path.isfile(full):
        raise ValueError("answers 文件不存在：%s" % full)
    with open(full, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError("answers 文件必须是 JSON 对象")
    return data


def _plain_question_lines(catalog: Catalog) -> List[str]:
    purpose_lines = [
        "   %d) %s —— %s（默认 %d 分钟 / %d 页）"
        % (i, p["name"], p["desc"], p["default_min"], p["default_pages"])
        for i, p in enumerate(PURPOSES, 1)
    ]
    source = (
        "内置 %d 所样例（%s）" % (len(catalog.schools), catalog.fallback_reason)
        if catalog.fallback
        else "%d 所（%s）" % (len(catalog.schools), catalog.school_source_path)
    )
    return [
        "学校数据：%s" % source,
        "",
        "请按顺序回答 8 个问题（可以直接把这段发给用户）：",
        "",
        "1. 学校：中文名 / 拼音 / 首字母 / 别名都行，例如 示例大学、dongbeidaxue、dbdx",
        "2. 学院 / 系名（可留空）：例如 示例学院",
        "3. 汇报人姓名",
        "4. 身份：本科生 / 研究生 / 博士生 / 教师 / 其他",
        "5. 汇报目的（填序号或名字）：",
    ] + purpose_lines + [
        "6. 计划时长（分钟；留空按目的默认值）",
        "7. 封面日期（留空用当前月份，例如 2026 年 11 月）",
        "8. 皮肤（auto / academic / campus / modern / bold，默认 auto）、"
        "动画（none / gentle / lively，默认 gentle）、说人话（默认开）",
        "",
        "可选：校徽 / 院徽图片路径（本机路径，PNG / SVG / JPG）。",
    ]


def _print_answers_summary(answers: dict) -> None:
    purpose = normalize_purpose(answers.get("purpose"))
    log("— 回答 —")
    log("  学校：%s" % (answers.get("school") or "（未填）"))
    log("  学院：%s" % (answers.get("college") or "（未填）"))
    log("  姓名：%s" % (answers.get("reporter") or "（未填）"))
    log("  身份：%s" % normalize_role(answers.get("role")))
    log("  目的：%s" % (purpose["name"] if purpose else "（未填）"))
    duration = _as_int(answers.get("duration_min"))
    if duration is None or duration <= 0:
        duration = purpose["default_min"] if purpose else "默认"
    log("  时长：%s 分钟" % duration)
    log("  日期：%s" % (answers.get("date") or default_date_text()))
    humanize = answers.get("humanize", True)
    if isinstance(humanize, str):
        humanize = humanize.strip().lower() not in ("0", "false", "no", "off", "否", "关")
    log("  皮肤：%s  动画：%s  说人话：%s"
        % (answers.get("skin") or "auto", answers.get("motion") or "gentle", "开" if humanize else "关"))
    if answers.get("crest"):
        log("  校徽：%s" % answers["crest"])
    if answers.get("logo"):
        log("  院徽：%s" % answers["logo"])


def _state_from_answers(answers: dict) -> Dict[str, Any]:
    return {
        "school_name": answers.get("school") or "",
        "school_slug": answers.get("school_slug") or "",
        "college": answers.get("college") or "",
        "college_en": answers.get("college_en") or "",
        "reporter": answers.get("reporter") or "",
        "role": answers.get("role"),
        "purpose": answers.get("purpose"),
        "duration_min": answers.get("duration_min"),
        "date": answers.get("date") or "",
        "skin": answers.get("skin") or "auto",
        "motion": answers.get("motion") or "gentle",
        "humanize": answers.get("humanize", True),
        "audience": answers.get("audience") or "",
        "motto": answers.get("motto") or "",
        "footer_left": answers.get("footer_left") or "",
        "footer_right": answers.get("footer_right") or "",
        "lang": answers.get("lang") or "zh",
        "crest_path": answers.get("crest") or "",
        "logo_path": answers.get("logo") or "",
        "brand_colors": answers.get("brand_colors"),
    }


def _generate_from_state(args, catalog: Catalog, state: Dict[str, Any], label: str) -> int:
    try:
        profile, warnings, saved_assets, user_assets_dir = compose_profile(catalog, state, args.out)
        profile_path = write_profile(args.out, profile)
    except ValueError as exc:
        warn("生成失败：%s" % exc)
        return 2
    except OSError as exc:
        warn("写文件失败：%s" % exc)
        return 2

    log("")
    log("profile.json 已写入：%s" % profile_path)
    log("  学校：%s%s" % (profile["school"] or "（未匹配）", "（%s）" % profile["school_slug"] if profile["school_slug"] else ""))
    if profile["college"]:
        log("  学院：%s" % profile["college"])
    log("  姓名 / 身份：%s / %s" % (profile["reporter"] or "（空）", profile["role"]))
    log("  目的：%s（%s 分钟，听众：%s）" % (PURPOSE_BY_KEY[profile["purpose"]]["name"], profile["duration_min"], profile["audience"]))
    log("  日期：%s" % profile["date"])
    log("  皮肤：%s  动画：%s  说人话：%s"
        % (profile["skin"], profile["motion"], "开" if profile["humanize"] else "关"))
    if profile["brand_colors"]:
        log("  品牌配色：primary %s / accent %s"
            % (profile["brand_colors"].get("primary"), profile["brand_colors"].get("accent")))
    if saved_assets:
        log("  素材目录：%s" % user_assets_dir)
        for asset in saved_assets:
            log("    已保存：%s" % asset["path"])
            if asset.get("palette"):
                log("    自动取色：%s" % asset["palette"]["primary"])
            elif asset.get("note"):
                log("    %s" % asset["note"])
    for message in warnings:
        warn("提示：%s" % message)
    return 0


def run_plain(args, catalog: Catalog) -> int:
    log("=" * 66)
    log("  %s（纯文本模式）" % WIZARD_TITLE)
    log("=" * 66)
    for line in _plain_question_lines(catalog):
        log(line)

    if args.answers:
        try:
            answers = _answers_from_file(args.answers)
        except (ValueError, OSError) as exc:
            warn("读取 answers 失败：%s" % exc)
            return 2
        log("")
        log("已读取答案文件：%s" % os.path.abspath(args.answers))
        _print_answers_summary(answers)
        return _generate_from_state(args, catalog, _state_from_answers(answers), "answers")

    log("")
    if not sys.stdin.isatty():
        log("当前不是交互终端，不逐项提问。")
        log("请把上面 8 个问题发给用户，拿到答案后写成 answers.json，再执行：")
        log("  python3 profile_wizard.py --out %s --answers answers.json" % args.out)
        return 0

    log("逐项回答（直接回车用默认值）：")
    answers: Dict[str, Any] = {}
    try:
        answers["school"] = input("1. 学校：").strip()
        answers["college"] = input("2. 学院 / 系名（可留空）：").strip()
        answers["reporter"] = input("3. 汇报人姓名：").strip()
        answers["role"] = input("4. 身份（本科生/研究生/博士生/教师/其他，默认研究生）：").strip() or "研究生"
        purpose_text = input("5. 汇报目的（序号或名字，默认 2 组会文献汇报）：").strip() or "2"
        answers["purpose"] = purpose_text
        purpose = normalize_purpose(purpose_text) or PURPOSE_BY_KEY["group_paper"]
        answers["duration_min"] = input("6. 计划时长（分钟，默认 %d）：" % purpose["default_min"]).strip() or str(purpose["default_min"])
        answers["date"] = input("7. 封面日期（默认 %s）：" % default_date_text()).strip() or default_date_text()
        answers["skin"] = input("8.1 皮肤（auto/academic/campus/modern/bold，默认 auto）：").strip() or "auto"
        answers["motion"] = input("8.2 动画（none/gentle/lively，默认 gentle）：").strip() or "gentle"
        humanize_text = input("8.3 说人话检查（开/关，默认开）：").strip()
        answers["humanize"] = humanize_text not in ("关", "否", "no", "false", "0")
        answers["crest"] = input("可选：校徽图片路径（直接回车跳过）：").strip()
        answers["logo"] = input("可选：院徽 / 实验室 logo 路径（直接回车跳过）：").strip()
    except EOFError:
        warn("输入提前结束。")
        return 2

    log("")
    _print_answers_summary(answers)
    return _generate_from_state(args, catalog, _state_from_answers(answers), "plain")


# ---------------------------------------------------------------------------
# 页面（全部内嵌：无 CDN、离线可用，手机宽度简单响应式）
# ---------------------------------------------------------------------------

PAGE_TEMPLATE = r"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>学院汇报 · 模板匹配向导</title>
<style>
:root{
  --p:#1F3A5F; --a:#3A6EA5; --d:#172B47; --s:#EDF1F7;
  --ink:#1b2330; --bg:#f4f6fa; --card:#ffffff; --line:#e2e8f1; --muted:#6b7789;
  --radius:14px; --shadow:0 8px 28px rgba(23,43,71,.07);
  --font:-apple-system,BlinkMacSystemFont,"PingFang SC","Hiragino Sans GB","Microsoft YaHei","Noto Sans CJK SC",sans-serif;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:var(--bg);color:var(--ink);font-family:var(--font);font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased}
button{font-family:inherit}
.wrap{max-width:1180px;margin:0 auto;padding:22px 20px 40px}
/* 顶部 */
.hero{display:flex;align-items:center;gap:14px;background:linear-gradient(135deg,var(--d),var(--p));color:#fff;border-radius:18px;padding:20px 24px;box-shadow:var(--shadow);transition:background .35s ease}
.hero-mark{flex:0 0 46px;height:46px;border-radius:12px;background:rgba(255,255,255,.16);display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:700}
.hero h1{margin:0;font-size:21px;letter-spacing:.5px}
.hero p{margin:4px 0 0;font-size:12.5px;opacity:.82}
.hero .hero-tag{margin-left:auto;font-size:11.5px;background:rgba(255,255,255,.14);padding:5px 11px;border-radius:999px;white-space:nowrap}
/* 布局 */
.layout{display:grid;grid-template-columns:minmax(0,1fr) 372px;gap:20px;margin-top:20px;align-items:start}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow)}
.main{min-width:0}
.side{position:sticky;top:18px}
/* 步骤 */
.steps{display:flex;gap:8px;margin:0 0 14px;padding:0;list-style:none}
.steps li{display:flex;align-items:center;gap:7px;font-size:13px;color:var(--muted);background:#fff;border:1px solid var(--line);border-radius:999px;padding:7px 14px 7px 8px;transition:.2s}
.steps li span{width:22px;height:22px;border-radius:50%;background:#eef1f6;color:#7a8698;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700}
.steps li.active{color:var(--p);border-color:var(--p);box-shadow:0 3px 10px rgba(31,58,95,.08)}
.steps li.active span{background:var(--p);color:#fff}
.steps li.done span{background:var(--s);color:var(--p)}
.steps li.done{color:var(--p)}
/* 面板 */
.panel{padding:24px 26px 26px}
.panel h2{margin:0 0 4px;font-size:17.5px;letter-spacing:.4px}
.panel .sub{margin:0 0 16px;color:var(--muted);font-size:12.5px}
.hidden{display:none !important}
/* 通用控件 */
.field{display:block;margin:0 0 16px}
.field>span{display:block;font-size:13px;color:#42506a;margin-bottom:6px;font-weight:600}
.field>span .req{color:#d14343}
.field>span .small{font-weight:400}
input[type=text],input[type=search],input:not([type]){
  width:100%;padding:11px 13px;border:1px solid var(--line);border-radius:10px;font-size:14.5px;font-family:inherit;background:#fbfcfe;color:var(--ink);outline:none;transition:.18s}
input:focus{border-color:var(--p);background:#fff;box-shadow:0 0 0 3px rgba(31,58,95,.09)}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.searchbox{position:relative;margin-bottom:10px}
.searchbox input{padding-right:92px}
.searchbox .searchcount{position:absolute;right:12px;top:50%;transform:translateY(-50%);font-size:12px;color:var(--muted)}
.school-list{max-height:342px;overflow:auto;border:1px solid var(--line);border-radius:12px;background:#fbfcfe;margin-bottom:16px;padding:4px}
.school{display:flex;align-items:center;gap:10px;width:100%;padding:9px 10px;border:0;background:transparent;border-radius:9px;cursor:pointer;text-align:left;transition:.14s}
.school:hover{background:#f0f4fa}
.school.sel{background:var(--s);box-shadow:inset 0 0 0 1.5px var(--p)}
.crest-mini{position:relative;flex:0 0 30px;height:30px;border-radius:8px;background:var(--p);color:#fff;overflow:hidden;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700}
.crest-mini img{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#fff;padding:2px}
.school-main{min-width:0;flex:1}
.school-name{display:block;font-size:14.5px;font-weight:600;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.school-meta{display:block;font-size:11.5px;color:var(--muted)}
.badges{display:flex;gap:4px;flex:0 0 auto}
.badges em{font-style:normal;font-size:10.5px;padding:2px 6px;border-radius:6px;background:#eef2f8;color:#5b6b85;white-space:nowrap}
.badges em.b985{background:#e9f1fd;color:#1b5ea8}
.badges em.b211{background:#eef7ef;color:#2c7a44}
.badges em.bsy{background:#fdf3e7;color:#a3661a}
.empty{padding:26px;text-align:center;color:var(--muted);font-size:13px}
.note{font-size:12px;color:var(--muted);background:#f8fafc;border:1px dashed var(--line);border-radius:10px;padding:9px 12px;margin:0 0 14px}
.note.warn{background:#fff8eb;border-color:#f0dfbd;color:#8a6414}
/* chips */
.chips{display:flex;flex-wrap:wrap;gap:8px}
.chip{border:1px solid var(--line);background:#fbfcfe;color:#44506a;border-radius:999px;padding:8px 15px;font-size:13.5px;cursor:pointer;transition:.15s}
.chip:hover{border-color:var(--p)}
.chip.sel{background:var(--p);border-color:var(--p);color:#fff}
/* 目的卡片 */
.purpose-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(206px,1fr));gap:11px}
.purpose{display:flex;flex-direction:column;gap:5px;text-align:left;border:1px solid var(--line);background:#fbfcfe;border-radius:12px;padding:12px 13px;cursor:pointer;transition:.16s}
.purpose:hover{border-color:var(--p);transform:translateY(-1px)}
.purpose.sel{border-color:var(--p);box-shadow:0 0 0 3px rgba(31,58,95,.09);background:#fff}
.purpose-name{font-size:14.5px;font-weight:700;color:var(--ink)}
.purpose-desc{font-size:12px;color:var(--muted);line-height:1.5}
.purpose-meta{font-size:11.5px;color:var(--p);font-weight:600;margin-top:2px}
/* 滑块 */
input[type=range]{width:100%;accent-color:var(--p);margin-top:6px}
.dur-val{color:var(--p);font-weight:700}
/* 皮肤卡片 */
.skin-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(178px,1fr));gap:11px}
.skin{display:flex;flex-direction:column;gap:7px;text-align:left;border:1px solid var(--line);background:#fbfcfe;border-radius:12px;padding:10px;cursor:pointer;transition:.16s}
.skin:hover{border-color:var(--p);transform:translateY(-1px)}
.skin.sel{border-color:var(--p);box-shadow:0 0 0 3px rgba(31,58,95,.09);background:#fff}
.wf{position:relative;display:block;height:62px;border-radius:8px;overflow:hidden;background:#fff;border:1px solid var(--line)}
.wf i{position:absolute;display:block;border-radius:1px}
.wf .l1{background:var(--p)}
.wf .l2{background:#2b3441}
.wf .l3{background:#c5cdd9}
.wf .l4{background:var(--a)}
.wf .l5{background:var(--s)}
.wf .l6{background:var(--d)}
.wf-auto i{background:var(--p)}
.wf-auto::after{content:"按目的自动选";position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:11.5px;color:var(--p);font-weight:700;background:linear-gradient(135deg,var(--s),#fff)}
.wf-academic .l1{left:7%;top:13%;width:26%;height:9%}
.wf-academic .l2{left:7%;top:34%;width:58%;height:8%}
.wf-academic .l3{left:7%;top:52%;width:76%;height:1.5px}
.wf-academic .l4{left:7%;top:62%;width:52%;height:4%;background:#c5cdd9}
.wf-academic .l5{left:7%;top:70%;width:44%;height:4%;background:#c5cdd9}
.wf-academic .l6{left:7%;bottom:8%;width:30%;height:4%;background:#c5cdd9}
.wf-campus .l1{left:0;top:0;bottom:0;width:40%}
.wf-campus .l2{left:7%;top:34%;width:26%;height:9%;background:#fff}
.wf-campus .l3{left:7%;top:48%;width:20%;height:6%;background:#fff;opacity:.8}
.wf-campus .l4{left:48%;top:22%;width:42%;height:7%;background:#2b3441}
.wf-campus .l5{left:48%;top:38%;width:36%;height:4%;background:#c5cdd9}
.wf-campus .l6{left:48%;top:47%;width:40%;height:4%;background:#c5cdd9}
.wf-modern .l1{left:7%;top:14%;width:16%;height:4%;background:#9aa6b8}
.wf-modern .l2{left:7%;top:26%;width:62%;height:9%;background:#2b3441}
.wf-modern .l3{left:7%;top:48%;width:80%;height:1.5px;background:#dfe5ee}
.wf-modern .l4{left:7%;top:58%;width:56%;height:4%;background:#c5cdd9}
.wf-modern .l5{left:7%;top:67%;width:48%;height:4%;background:#c5cdd9}
.wf-modern .l6{left:7%;bottom:10%;width:22%;height:3%;background:var(--p)}
.wf-bold{background:var(--d)}
.wf-bold .l1{right:0;top:0;width:46%;height:64%;background:var(--p);opacity:.9}
.wf-bold .l2{left:8%;top:26%;width:44%;height:9%;background:#fff}
.wf-bold .l3{left:8%;top:44%;width:32%;height:6%;background:#fff;opacity:.75}
.wf-bold .l4{left:8%;bottom:12%;width:56%;height:3%;background:var(--a)}
.wf-bold .l5{right:8%;bottom:12%;width:26%;height:14%;background:var(--a);opacity:.85}
.skin-name{font-size:13.5px;font-weight:700}
.skin-desc{font-size:11.5px;color:var(--muted);line-height:1.45}
/* 开关 */
.toggle{display:flex;align-items:center;gap:10px;cursor:pointer;margin:2px 0 18px}
.toggle input{position:absolute;opacity:0;width:0;height:0}
.toggle .track{flex:0 0 44px;height:25px;border-radius:999px;background:#cfd8e5;position:relative;transition:.18s}
.toggle .track::after{content:"";position:absolute;left:3px;top:3px;width:19px;height:19px;border-radius:50%;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.18);transition:.18s}
.toggle input:checked+.track{background:var(--p)}
.toggle input:checked+.track::after{transform:translateX(19px)}
.toggle .txt{font-size:13.5px;font-weight:600}
.toggle .txt small{display:block;font-weight:400;color:var(--muted);font-size:11.5px}
/* 上传 */
.uploads{display:grid;grid-template-columns:1fr 1fr;gap:13px;margin-bottom:8px}
.upload{border:1px dashed #cfd8e5;border-radius:12px;padding:13px;background:#fbfcfe;min-height:132px}
.upload .up-title{font-size:13.5px;font-weight:700}
.upload .up-sub{font-size:11.5px;color:var(--muted);margin:2px 0 9px}
.upload .up-preview{display:flex;align-items:center;gap:8px;min-height:34px}
.upload .up-preview img{max-height:34px;max-width:104px;border-radius:4px;background:#fff;border:1px solid var(--line);padding:2px}
.upload .up-preview .name{font-size:11.5px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.upload .palette{display:flex;gap:5px;margin-top:7px;flex-wrap:wrap}
.upload .palette span{width:26px;height:16px;border-radius:4px;border:1px solid rgba(0,0,0,.08)}
.upload .up-note{font-size:11px;color:var(--muted);margin-top:7px;min-height:14px}
.upload .up-note.err{color:#c14a3f}
.btn{font-family:inherit;border-radius:10px;font-size:14px;cursor:pointer;transition:.16s;border:1px solid transparent;padding:11px 20px;font-weight:600}
.btn.primary{background:var(--p);color:#fff;box-shadow:0 4px 14px rgba(31,58,95,.18)}
.btn.primary:hover{filter:brightness(1.07)}
.btn.primary:disabled{opacity:.65;cursor:default}
.btn.ghost{background:#fff;border-color:var(--line);color:#44506a}
.btn.ghost:hover{border-color:var(--p);color:var(--p)}
.btn.small{padding:7px 12px;font-size:12.5px;border-radius:8px}
.btn.danger{background:#fff;border-color:#f0d3d0;color:#b3453c}
.btn.danger:hover{border-color:#d14343}
.actions{display:flex;align-items:center;gap:10px;margin-top:16px;flex-wrap:wrap}
.actions .hint{font-size:12.5px;color:var(--muted);margin-left:auto}
.actions .hint.err{color:#c14a3f;font-weight:600}
/* 侧边预览 */
.preview{padding:16px}
.preview h3{margin:0 0 10px;font-size:13.5px;color:#42506a;letter-spacing:.5px}
.pv-slide{position:relative;background:#fff;border:1px solid var(--line);border-radius:10px;aspect-ratio:16/9;padding:12px 14px;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 2px 10px rgba(23,43,71,.05)}
.pv-slide::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:var(--p)}
.pv-top{display:flex;align-items:center;gap:7px;min-height:20px}
.pv-crest-fallback{flex:0 0 auto;width:34px;height:15px;border-radius:4px;background:var(--p);color:#fff;font-size:9.5px;display:flex;align-items:center;justify-content:center;font-weight:700;padding:0 3px;overflow:hidden}
.pv-crest{max-height:17px;max-width:88px;object-fit:contain}
.pv-school{font-size:11px;font-weight:700;color:var(--p);line-height:1.2}
.pv-college{font-size:8.5px;color:var(--muted);line-height:1.2}
.pv-title{margin-top:auto;font-size:15px;font-weight:800;color:#1b2330;letter-spacing:.3px}
.pv-title::after{content:"";display:block;width:34px;height:2.5px;border-radius:2px;background:var(--a);margin-top:5px}
.pv-sub{font-size:10px;color:#55617a;margin-top:6px}
.pv-lines{margin-top:9px;display:flex;flex-direction:column;gap:4px}
.pv-lines i{display:block;height:3px;border-radius:2px;background:#e4e9f1}
.pv-lines i:nth-child(1){width:74%}
.pv-lines i:nth-child(2){width:62%}
.pv-lines i:nth-child(3){width:48%}
.pv-foot{position:absolute;left:14px;right:14px;bottom:7px;display:flex;justify-content:space-between;gap:8px;font-size:7.5px;color:#8d99ab;border-top:1px solid #eef1f6;padding-top:4px}
.pv-foot span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pv-colors{display:flex;gap:6px;margin-top:11px}
.pv-colors span{flex:1;height:16px;border-radius:5px;position:relative;border:1px solid rgba(0,0,0,.06)}
.pv-colors b{position:absolute;left:0;right:0;top:18px;font-size:8.5px;font-weight:500;color:var(--muted);text-align:center}
.pv-colors{margin-bottom:22px}
.pv-meta{display:grid;grid-template-columns:1fr 1fr;gap:8px 10px;margin:0;font-size:11.5px}
.pv-meta div{min-width:0}
.pv-meta dt{color:var(--muted);font-size:10.5px;margin:0}
.pv-meta dd{margin:1px 0 0;font-weight:600;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pv-warn{margin-top:12px;font-size:11.5px;color:#8a6414;background:#fff8eb;border:1px solid #f0dfbd;border-radius:8px;padding:7px 10px;display:none}
/* 成功页 */
.done{padding:36px 30px;text-align:center}
.done .done-icon{width:58px;height:58px;border-radius:50%;background:var(--p);color:#fff;font-size:28px;display:flex;align-items:center;justify-content:center;margin:0 auto 14px}
.done h2{margin:0 0 6px;font-size:20px}
.done .muted{color:var(--muted);font-size:13px}
.summary{width:100%;max-width:560px;margin:20px auto 0;border-collapse:collapse;text-align:left;font-size:13.5px}
.summary th{width:34%;color:var(--muted);font-weight:500;padding:8px 10px;border-top:1px solid var(--line);vertical-align:top}
.summary td{padding:8px 10px;border-top:1px solid var(--line);font-weight:600;word-break:break-all}
.pathbox{display:flex;align-items:center;gap:10px;justify-content:center;margin-top:18px;flex-wrap:wrap}
.pathbox span{font-size:12px;color:var(--muted)}
.pathbox code{background:#f2f5fa;border:1px solid var(--line);border-radius:8px;padding:7px 11px;font-size:12.5px;color:var(--p);word-break:break-all}
.done .warn-list{margin:16px auto 0;max-width:560px;text-align:left;font-size:12.5px;color:#8a6414;background:#fff8eb;border:1px solid #f0dfbd;border-radius:10px;padding:10px 14px}
.done .warn-list li{margin:3px 0}
footer{margin-top:22px;text-align:center;font-size:11.5px;color:#93a0b2}
/* 响应式 */
@media (max-width: 980px){
  .layout{grid-template-columns:1fr}
  .side{position:static;order:2}
  .main{order:1}
  .hero .hero-tag{display:none}
  .panel{padding:18px 16px 20px}
  .uploads{grid-template-columns:1fr}
  .grid2{grid-template-columns:1fr}
  .wrap{padding:14px 12px 30px}
  .hero{padding:16px 16px;border-radius:14px}
  .hero h1{font-size:17px}
}
@media (max-width: 560px){
  body{font-size:14px}
  .steps li{flex:1;justify-content:center;padding:7px 6px;font-size:12px}
  .steps li span{width:19px;height:19px;font-size:11px}
  .purpose-grid{grid-template-columns:1fr}
  .skin-grid{grid-template-columns:1fr 1fr}
  .school-list{max-height:300px}
  .badges{display:none}
  .pv-title{font-size:13px}
}
</style>
</head>
<body>
<div class="wrap">
  <header class="hero" id="hero">
    <div class="hero-mark">院</div>
    <div>
      <h1>学院汇报 · 模板匹配向导</h1>
      <p>三步生成汇报档案 · 全部在本机运行，不上传任何数据</p>
    </div>
    <span class="hero-tag" id="heroTag">离线向导</span>
  </header>

  <div class="layout">
    <main class="main">
      <ol class="steps" id="steps">
        <li data-step="1" class="active"><span>1</span>选学校</li>
        <li data-step="2"><span>2</span>填信息</li>
        <li data-step="3"><span>3</span>选风格</li>
      </ol>

      <!-- 第 1 步 -->
      <section class="card panel" id="panel1">
        <h2>第 1 步 · 选学校</h2>
        <p class="sub" id="schoolHint">搜索学校，选中后右侧会实时预览这所学校的模板配色。</p>
        <div id="fallbackNote" class="note warn hidden"></div>
        <div class="searchbox">
          <input id="q" type="search" autocomplete="off" placeholder="搜学校：中文名 / 拼音 / 首字母 / 别名，如 示例大学、dongbei、dbdx">
          <span class="searchcount" id="count"></span>
        </div>
        <div class="school-list" id="list" role="listbox" aria-label="学校列表"></div>
        <div class="grid2">
          <label class="field"><span>学院 / 系名（选填，显示在封面和页脚）</span>
            <input id="college" type="text" placeholder="如：示例学院" autocomplete="off"></label>
          <label class="field"><span>学院英文名（选填，页脚右侧用）</span>
            <input id="collegeEn" type="text" placeholder="如：Sydney Smart Technology College" autocomplete="off"></label>
        </div>
      </section>

      <!-- 第 2 步 -->
      <section class="card panel hidden" id="panel2">
        <h2>第 2 步 · 填汇报信息</h2>
        <p class="sub">姓名和目的必填；时长按目的给了默认值，可以拖。</p>
        <div class="grid2">
          <label class="field"><span>汇报人姓名 <b class="req">*</b></span>
            <input id="reporter" type="text" placeholder="如：张三" autocomplete="off"></label>
          <label class="field"><span>封面日期</span>
            <input id="date" type="text" placeholder="如：2026 年 11 月" autocomplete="off"></label>
        </div>
        <div class="field"><span>身份 <b class="req">*</b></span>
          <div class="chips" id="roles"></div>
        </div>
        <div class="field"><span>汇报目的 <b class="req">*</b></span>
          <div class="purpose-grid" id="purposes"></div>
        </div>
        <div class="field">
          <span>计划时长：<b class="dur-val" id="durVal">22</b> 分钟
            <span class="small" id="durHint" style="color:#8a97a9"></span></span>
          <input type="range" id="duration" min="5" max="60" step="1" value="22">
        </div>
      </section>

      <!-- 第 3 步 -->
      <section class="card panel hidden" id="panel3">
        <h2>第 3 步 · 选风格与素材</h2>
        <p class="sub">皮肤决定整体视觉；动画档位决定转场与要点出现方式；校徽院徽可选。</p>
        <div class="field"><span>皮肤</span>
          <div class="skin-grid" id="skins"></div>
        </div>
        <div class="field"><span>动画档位</span>
          <div class="chips" id="motions"></div>
        </div>
        <label class="toggle">
          <input type="checkbox" id="humanize" checked>
          <span class="track"></span>
          <span class="txt">「说人话」检查<small>正文与讲稿去掉 AI 味、空话套话（推荐开）</small></span>
        </label>
        <div class="uploads">
          <div class="upload">
            <div class="up-title">校徽<br>（可选）</div>
            <div class="up-sub">PNG / SVG / JPG，上传后自动取色</div>
            <input type="file" id="fileCrest" accept="image/png,image/svg+xml,image/jpeg,image/webp,image/gif" hidden>
            <button class="btn ghost small" id="pickCrest" type="button">选择文件</button>
            <div class="up-preview" id="crestPreview"></div>
            <div class="palette" id="crestPalette"></div>
            <div class="up-note" id="crestNote"></div>
          </div>
          <div class="upload">
            <div class="up-title">院徽 / 实验室 logo<br>（可选）</div>
            <div class="up-sub">放封面右下与页脚右侧</div>
            <input type="file" id="fileLogo" accept="image/png,image/svg+xml,image/jpeg,image/webp,image/gif" hidden>
            <button class="btn ghost small" id="pickLogo" type="button">选择文件</button>
            <div class="up-preview" id="logoPreview"></div>
            <div class="palette" id="logoPalette"></div>
            <div class="up-note" id="logoNote"></div>
          </div>
        </div>
        <p class="note">隐私：图片只写入本机 <code>user_assets/</code> 目录，向导不联网、不上传。SVG 能保存但无法自动取色。</p>
      </section>

      <!-- 完成 -->
      <section class="card panel done hidden" id="panelDone">
        <div class="done-icon">✓</div>
        <h2>汇报档案已生成</h2>
        <p class="muted">可以直接关掉这个页面了；本地服务 2 秒后自动退出。</p>
        <table class="summary" id="doneSummary"></table>
        <div class="pathbox"><span>profile.json</span><code id="donePath"></code></div>
        <p class="muted" style="margin-top:10px">素材目录：<code id="doneAssets" style="font-size:11.5px">—</code></p>
        <div class="warn-list hidden" id="doneWarn"></div>
      </section>

      <div class="actions" id="actions">
        <button class="btn ghost hidden" id="btnBack" type="button">← 上一步</button>
        <button class="btn primary" id="btnNext" type="button">下一步 →</button>
        <button class="btn primary hidden" id="btnSubmit" type="button">生成汇报档案</button>
        <span class="hint" id="navHint"></span>
      </div>
    </main>

    <aside class="side">
      <div class="card preview">
        <h3>实时预览</h3>
        <div class="pv-slide">
          <div class="pv-top">
            <img class="pv-crest hidden" id="pvCrest" alt="">
            <span class="pv-crest-fallback" id="pvFallback">院</span>
            <div>
              <div class="pv-school" id="pvSchool">未选择学校</div>
              <div class="pv-college" id="pvCollege">学院 / 系名</div>
            </div>
          </div>
          <div class="pv-title" id="pvTitle">汇报题目</div>
          <div class="pv-sub" id="pvSub">汇报人 · 日期</div>
          <div class="pv-lines"><i></i><i></i><i></i></div>
          <div class="pv-foot"><span id="pvMotto">校训</span><span id="pvDate">日期</span></div>
        </div>
        <div class="pv-colors" id="pvColors"></div>
        <dl class="pv-meta">
          <div><dt>学校</dt><dd id="mSchool">—</dd></div>
          <div><dt>姓名 / 身份</dt><dd id="mPerson">—</dd></div>
          <div><dt>目的</dt><dd id="mPurpose">—</dd></div>
          <div><dt>时长</dt><dd id="mDur">—</dd></div>
          <div><dt>皮肤</dt><dd id="mSkin">—</dd></div>
          <div><dt>动画</dt><dd id="mMotion">—</dd></div>
        </dl>
        <div class="pv-warn" id="pvWarn"></div>
      </div>
    </aside>
  </div>
  <footer>学院汇报 PPT 技能 v2 · 本地向导 · 生成的文件只保存在这台电脑上</footer>
</div>

<script>
const DATA = __DATA_JSON__;
const $ = (id) => document.getElementById(id);
const state = {
  step: 1,
  school: null,
  college: "", collegeEn: "",
  reporter: "", role: "研究生", purpose: null,
  duration: null, durationTouched: false,
  date: DATA.meta.default_date,
  skin: "auto", motion: "gentle", humanize: true,
  crest: null, logo: null, crestPalette: null, logoPalette: null,
  submitting: false,
};

/* ---------- 小工具 ---------- */
function esc(s){ return String(s == null ? "" : s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function hexOk(v){ return /^#[0-9a-fA-F]{6}$/.test(v || ""); }
function setTheme(colors){
  const root = document.documentElement.style;
  root.setProperty("--p", colors.primary);
  root.setProperty("--a", colors.accent);
  root.setProperty("--d", colors.deep || colors.primary);
  root.setProperty("--s", colors.soft || "#EDF1F7");
}
function navHint(text, isErr){
  const el = $("navHint");
  el.textContent = text || "";
  el.className = "hint" + (isErr ? " err" : "");
}
function showWarn(text){
  const el = $("pvWarn");
  if (text){ el.textContent = text; el.style.display = "block"; }
  else { el.textContent = ""; el.style.display = "none"; }
}
function resolvedSkin(){
  if (state.skin !== "auto") return state.skin;
  return state.purpose ? state.purpose.skin : "academic";
}

/* ---------- 学校列表 ---------- */
function schoolScore(s, q){
  if (!q) return 0;
  const n = s.name.toLowerCase();
  if (n === q) return 0;
  if (n.startsWith(q)) return 1;
  if (n.includes(q)) return 2;
  const aliases = (s.aliases || []).map(a => a.toLowerCase());
  if (aliases.some(a => a === q)) return 3;
  if (aliases.some(a => a.includes(q))) return 4;
  if (s.abbr === q) return 5;
  if (s.pinyin.startsWith(q)) return 6;
  if (s.pinyin.includes(q)) return 7;
  if ((s.province || "").includes(q)) return 8;
  if ((s.name_en || "").toLowerCase().includes(q)) return 9;
  return 20;
}
function filterSchools(q){
  q = (q || "").trim().toLowerCase();
  if (!q) return DATA.schools.slice(0, 60);
  const tokens = q.split(/\s+/).filter(Boolean);
  const hits = [];
  for (const s of DATA.schools){
    const hay = [s.name, s.name_en, s.province, s.city, s.pinyin, s.abbr].concat(s.aliases || []).join(" ").toLowerCase();
    if (tokens.some(t => !hay.includes(t))) continue;
    hits.push({ s, score: schoolScore(s, q) });
  }
  hits.sort((a, b) => a.score - b.score);
  return hits.slice(0, 80).map(x => x.s);
}
function renderList(){
  const q = $("q").value;
  const hits = filterSchools(q);
  const listEl = $("list");
  if (!hits.length){
    listEl.innerHTML = '<div class="empty">没有匹配的学校。试试中文名、拼音（dongbei）或首字母（dbdx）。</div>';
    $("count").textContent = "0 所";
    return;
  }
  listEl.innerHTML = hits.map(s => {
    const sel = state.school && state.school.slug === s.slug ? " sel" : "";
    const badges = (s.level || []).slice(0, 3).map(l => {
      const cls = l === "985" ? "b985" : (l === "211" ? "b211" : "bsy");
      return '<em class="' + cls + '">' + esc(l) + '</em>';
    }).join("");
    const crest = s.crest_url
      ? '<img src="' + s.crest_url + '" alt="" loading="lazy" onerror="this.remove()">'
      : "";
    const meta = [s.province, (s.city && s.city !== s.province) ? s.city : "", s.type].filter(Boolean).join(" · ");
    return '<button class="school' + sel + '" type="button" data-slug="' + s.slug + '" role="option">'
      + '<span class="crest-mini" style="background:' + s.colors.primary + '"><b>' + esc(s.name.slice(0,1)) + '</b>' + crest + '</span>'
      + '<span class="school-main"><span class="school-name">' + esc(s.name) + '</span>'
      + '<span class="school-meta">' + esc(meta) + '</span></span>'
      + '<span class="badges">' + badges + '</span></button>';
  }).join("");
  const total = DATA.meta.school_count;
  $("count").textContent = (q ? hits.length : Math.min(60, total)) + " / " + total + " 所";
}
let collegeAuto = false;
function selectSchool(slug){
  const s = DATA.schools.find(x => x.slug === slug);
  if (!s) return;
  state.school = s;
  setTheme(s.colors);
  // 用户没手填过学院时，跟着学校带出默认学院（换校时也更新/清空）
  if (!state.college || collegeAuto){
    const next = s.college_default || "";
    $("college").value = next;
    state.college = next;
    collegeAuto = !!next;
  }
  state.college = $("college").value.trim();
  renderList();
  updatePreview();
  navHint("");
}

/* ---------- 目的 / 身份 / 皮肤 / 动画 ---------- */
function renderPurposes(){
  $("purposes").innerHTML = DATA.purposes.map(p =>
    '<button class="purpose' + (state.purpose && state.purpose.key === p.key ? " sel" : "") + '" type="button" data-key="' + p.key + '">'
    + '<span class="purpose-name">' + esc(p.name) + '</span>'
    + '<span class="purpose-desc">' + esc(p.desc) + '</span>'
    + '<span class="purpose-meta">默认 ' + p.default_min + ' 分钟 · ' + p.default_pages + ' 页</span></button>'
  ).join("");
}
function selectPurpose(key){
  const p = DATA.purposes.find(x => x.key === key);
  if (!p) return;
  state.purpose = p;
  if (!state.durationTouched || state.duration == null){
    state.duration = p.default_min;
    $("duration").value = p.default_min;
    $("durVal").textContent = p.default_min;
  }
  $("durHint").textContent = "（" + p.name + "默认 " + p.default_min + " 分钟 / " + p.default_pages + " 页）";
  renderPurposes();
  updatePreview();
  navHint("");
}
function renderRoles(){
  $("roles").innerHTML = DATA.roles.map(r =>
    '<button class="chip' + (state.role === r ? " sel" : "") + '" type="button" data-role="' + esc(r) + '">' + esc(r) + '</button>'
  ).join("");
}
function wfInner(key){
  if (key === "auto") return "";
  return '<i class="l1"></i><i class="l2"></i><i class="l3"></i><i class="l4"></i><i class="l5"></i><i class="l6"></i>';
}
function renderSkins(){
  $("skins").innerHTML = DATA.skins.map(sk =>
    '<button class="skin' + (state.skin === sk.key ? " sel" : "") + '" type="button" data-key="' + sk.key + '">'
    + '<span class="wf wf-' + sk.key + '">' + wfInner(sk.key) + '</span>'
    + '<span class="skin-name">' + esc(sk.name) + '</span>'
    + '<span class="skin-desc">' + esc(sk.desc) + '</span></button>'
  ).join("");
}
function renderMotions(){
  $("motions").innerHTML = DATA.motions.map(m =>
    '<button class="chip' + (state.motion === m.key ? " sel" : "") + '" type="button" data-key="' + m.key + '" title="' + esc(m.desc) + '">' + esc(m.name) + '</button>'
  ).join("");
}

/* ---------- 预览 ---------- */
function updatePreview(){
  const s = state.school;
  const colors = s ? s.colors : DATA.meta.neutral_colors;
  const crestImg = $("pvCrest"), fallback = $("pvFallback");
  if (s && s.crest_url){
    crestImg.src = s.crest_url;
    crestImg.classList.remove("hidden");
    fallback.classList.add("hidden");
  } else {
    crestImg.classList.add("hidden");
    crestImg.removeAttribute("src");
    fallback.classList.remove("hidden");
    fallback.textContent = s ? s.name.slice(0, 1) : "院";
    fallback.style.background = colors.primary;
  }
  $("pvSchool").textContent = s ? s.name : "未选择学校";
  $("pvCollege").textContent = state.college || (s && s.college_default) || "学院 / 系名";
  $("pvSub").textContent = (state.reporter || "汇报人") + " · " + (state.date || DATA.meta.default_date);
  $("pvMotto").textContent = (s && s.motto) || "校训";
  $("pvDate").textContent = state.date || DATA.meta.default_date;

  $("pvColors").innerHTML = ["primary","accent","deep","soft"].map(k =>
    '<span style="background:' + colors[k] + '" title="' + k + ' ' + colors[k] + '"><b>' + colors[k] + '</b></span>'
  ).join("");

  $("mSchool").textContent = s ? s.name : "—";
  $("mPerson").textContent = (state.reporter || "—") + " / " + state.role;
  $("mPurpose").textContent = state.purpose ? state.purpose.name : "—";
  $("mDur").textContent = (state.duration || "—") + " 分钟";
  const skinText = state.skin === "auto" ? "auto → " + resolvedSkin() : state.skin;
  $("mSkin").textContent = skinText;
  $("mMotion").textContent = state.motion;

  const why = [];
  if (s && !s.has_brand) why.push("这所学校没有品牌包，预览用的是中性色。");
  if (state.crestPalette) why.push("已从校徽取色，提交后会写入 brand_colors。");
  showWarn(why.join(" "));
}

/* ---------- 上传 ---------- */
function renderUpload(kind){
  const item = kind === "crest" ? state.crest : state.logo;
  const palette = kind === "crest" ? state.crestPalette : state.logoPalette;
  const preview = $(kind === "crest" ? "crestPreview" : "logoPreview");
  const paletteEl = $(kind === "crest" ? "crestPalette" : "logoPalette");
  if (!item){
    preview.innerHTML = "";
    paletteEl.innerHTML = "";
    return;
  }
  preview.innerHTML = '<img src="' + item.dataUrl + '" alt=""><span class="name">' + esc(item.name) + '</span>'
    + '<button class="btn danger small" type="button" data-clear="' + kind + '">移除</button>';
  paletteEl.innerHTML = palette ? ["primary","accent","deep","soft"].map(k =>
    '<span style="background:' + palette[k] + '" title="' + k + ' ' + palette[k] + '"></span>'
  ).join("") : "";
  preview.querySelector("[data-clear]").addEventListener("click", () => {
    if (kind === "crest"){ state.crest = null; state.crestPalette = null; }
    else { state.logo = null; state.logoPalette = null; }
    $(kind === "crest" ? "fileCrest" : "fileLogo").value = "";
    $(kind === "crest" ? "crestNote" : "logoNote").textContent = "";
    renderUpload(kind);
    updatePreview();
  });
}
function note(kind, text, isErr){
  const el = $(kind === "crest" ? "crestNote" : "logoNote");
  el.textContent = text || "";
  el.className = "up-note" + (isErr ? " err" : "");
}
function bindUpload(kind){
  const fileInput = $(kind === "crest" ? "fileCrest" : "fileLogo");
  const pick = $(kind === "crest" ? "pickCrest" : "pickLogo");
  pick.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    const file = fileInput.files && fileInput.files[0];
    if (!file) return;
    if (file.size > 4 * 1024 * 1024){ note(kind, "文件超过 4 MB，请压缩后再上传", true); fileInput.value = ""; return; }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result);
      if (kind === "crest"){ state.crest = { name: file.name, dataUrl }; state.crestPalette = null; }
      else { state.logo = { name: file.name, dataUrl }; state.logoPalette = null; }
      renderUpload(kind);
      note(kind, "已载入，正在自动取色…");
      fetch("/api/palette", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ data_url: dataUrl, kind: kind })
      }).then(r => r.json()).then(j => {
        if (j && j.ok && j.palette){
          if (kind === "crest") state.crestPalette = j.palette; else state.logoPalette = j.palette;
          note(kind, "自动取色完成：" + j.palette.primary + " / " + j.palette.accent);
        } else {
          note(kind, (j && j.error) || "无法自动取色（文件仍会保存）", true);
        }
        renderUpload(kind);
        updatePreview();
      }).catch(() => {
        note(kind, "取色服务没响应，提交时仍会保存图片", true);
        renderUpload(kind);
      });
    };
    reader.readAsDataURL(file);
  });
}

/* ---------- 步骤 ---------- */
function goStep(step){
  if (step > 1 && !state.school){
    navHint("请先在第 1 步选择学校", true);
    $("q").focus();
    return;
  }
  if (step > 2){
    if (!state.reporter.trim()){ navHint("请填写汇报人姓名", true); $("reporter").focus(); return; }
    if (!state.purpose){ navHint("请选择汇报目的", true); return; }
    if (state.duration == null) state.duration = state.purpose.default_min;
  }
  state.step = step;
  navHint("");
  ["panel1","panel2","panel3","panelDone"].forEach((id, i) => {
    $(id).classList.toggle("hidden", i !== step - 1);
  });
  document.querySelectorAll("#steps li").forEach(li => {
    const n = Number(li.dataset.step);
    li.classList.toggle("active", n === step);
    li.classList.toggle("done", n < step);
  });
  const last = step === 3;
  $("btnBack").classList.toggle("hidden", step <= 1 || step > 3);
  $("btnNext").classList.toggle("hidden", last || step > 3);
  $("btnSubmit").classList.toggle("hidden", !last || step > 3);
  $("steps").classList.toggle("hidden", step > 3);
  $("actions").classList.toggle("hidden", step > 3);
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ---------- 提交 ---------- */
function validateAll(){
  if (!state.school) return "请先选择学校";
  if (!state.reporter.trim()) return "请填写汇报人姓名";
  if (!state.purpose) return "请选择汇报目的";
  return "";
}
function submit(){
  if (state.submitting) return;
  const err = validateAll();
  if (err){ navHint(err, true); return; }
  state.submitting = true;
  const btn = $("btnSubmit");
  btn.disabled = true;
  btn.textContent = "正在生成…";
  navHint("");
  const payload = {
    school: state.school.name,
    school_slug: state.school.slug,
    college: $("college").value.trim(),
    college_en: $("collegeEn").value.trim(),
    reporter: state.reporter.trim(),
    role: state.role,
    purpose: state.purpose.key,
    duration_min: state.duration,
    date: state.date.trim(),
    skin: state.skin,
    motion: state.motion,
    humanize: state.humanize,
    audience: state.purpose.audience,
    crest: state.crest ? { name: state.crest.name, data_url: state.crest.dataUrl } : null,
    logo: state.logo ? { name: state.logo.name, data_url: state.logo.dataUrl } : null,
  };
  fetch("/api/save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  }).then(r => r.json().then(j => ({ status: r.status, body: j }))).then(({ status, body }) => {
    if (!body || !body.ok) throw new Error((body && body.error) || "生成失败（HTTP " + status + "）");
    showDone(body);
  }).catch(e => {
    state.submitting = false;
    btn.disabled = false;
    btn.textContent = "生成汇报档案";
    navHint(e.message || "生成失败，请重试", true);
  });
}
const PURPOSE_NAMES = {};
DATA.purposes.forEach(p => PURPOSE_NAMES[p.key] = p.name);
function showDone(res){
  state.submitting = true;
  const profile = res.profile || {};
  const skinText = profile.skin === "auto" ? "auto → " + resolvedSkin() : profile.skin;
  const motionNames = { none: "无动画", gentle: "轻淡入", lively: "活泼" };
  const crestAsset = (res.saved_assets || []).find(a => a.kind === "crest");
  const logoAsset = (res.saved_assets || []).find(a => a.kind === "logo");
  const rows = [
    ["学校 / 学院", (profile.school || "—") + (profile.school_slug ? "（" + profile.school_slug + "）" : "") + (profile.college ? " · " + profile.college : "")],
    ["汇报人 / 身份", (profile.reporter || "（空）") + " / " + (profile.role || "—")],
    ["汇报目的", PURPOSE_NAMES[profile.purpose] || profile.purpose || "—"],
    ["计划时长 / 听众", (profile.duration_min || "—") + " 分钟 / " + (profile.audience || "—")],
    ["封面日期", profile.date || "—"],
    ["皮肤 / 动画", skinText + " / " + (motionNames[profile.motion] || profile.motion)],
    ["说人话检查", profile.humanize ? "开" : "关"],
    ["品牌配色", profile.brand_colors
      ? (profile.brand_colors.primary || "") + " / " + (profile.brand_colors.accent || "")
      : "沿用模板库（未从图片取色）"],
    ["校徽", crestAsset ? crestAsset.path : "未上传（用模板库校徽）"],
    ["院徽 / logo", logoAsset ? logoAsset.path : "未上传"],
  ];
  $("doneSummary").innerHTML = rows.map(r => "<tr><th>" + esc(r[0]) + "</th><td>" + esc(r[1]) + "</td></tr>").join("");
  $("donePath").textContent = res.profile_path || "";
  $("doneAssets").textContent = res.user_assets_dir || "—";
  const warnEl = $("doneWarn");
  const warnings = res.warnings || [];
  if (warnings.length){
    warnEl.innerHTML = "<b>提示</b><ul>" + warnings.map(w => "<li>" + esc(w) + "</li>").join("") + "</ul>";
    warnEl.classList.remove("hidden");
  } else {
    warnEl.classList.add("hidden");
  }
  $("panel1").classList.add("hidden");
  $("panel2").classList.add("hidden");
  $("panel3").classList.add("hidden");
  $("panelDone").classList.remove("hidden");
  $("steps").classList.add("hidden");
  $("actions").classList.add("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ---------- 初始化 ---------- */
(function init(){
  const meta = DATA.meta;
  $("heroTag").textContent = "共 " + meta.school_count + " 所院校";
  $("schoolHint").textContent = "共 " + meta.school_count + " 所院校，按 985 / 211 / 双一流 排序；输入中文名、拼音或首字母都能搜到。";
  if (meta.fallback){
    const note = $("fallbackNote");
    note.innerHTML = '<b>提示</b>：' + esc(meta.fallback_reason) + '，现在用内置 15 所样例。把完整 <code>schools.json</code> 放到 <code>assets/schools/</code> 后重启向导，即可搜索全国院校。';
    note.classList.remove("hidden");
  }
  if (window.innerWidth < 560) $("q").placeholder = "搜学校：中文名 / 拼音 / 首字母";
  $("date").value = meta.default_date;
  renderList();
  renderRoles();
  renderPurposes();
  renderSkins();
  renderMotions();
  updatePreview();
  setTheme(meta.neutral_colors);

  $("q").addEventListener("input", renderList);
  $("q").addEventListener("keydown", (e) => {
    if (e.key === "Enter"){
      e.preventDefault();
      const first = $("list").querySelector(".school");
      if (first) selectSchool(first.dataset.slug);
    }
  });
  $("list").addEventListener("click", (e) => {
    const row = e.target.closest(".school");
    if (row) selectSchool(row.dataset.slug);
  });
  $("college").addEventListener("input", () => { collegeAuto = false; state.college = $("college").value.trim(); updatePreview(); });
  $("collegeEn").addEventListener("input", () => { state.collegeEn = $("collegeEn").value.trim(); });
  $("reporter").addEventListener("input", () => { state.reporter = $("reporter").value; updatePreview(); });
  $("date").addEventListener("input", () => { state.date = $("date").value; updatePreview(); });
  $("roles").addEventListener("click", (e) => {
    const chip = e.target.closest("[data-role]");
    if (!chip) return;
    state.role = chip.dataset.role;
    renderRoles();
    updatePreview();
  });
  $("purposes").addEventListener("click", (e) => {
    const card = e.target.closest(".purpose");
    if (card) selectPurpose(card.dataset.key);
  });
  $("duration").addEventListener("input", () => {
    state.duration = Number($("duration").value);
    state.durationTouched = true;
    $("durVal").textContent = state.duration;
    updatePreview();
  });
  $("skins").addEventListener("click", (e) => {
    const card = e.target.closest(".skin");
    if (!card) return;
    state.skin = card.dataset.key;
    renderSkins();
    updatePreview();
  });
  $("motions").addEventListener("click", (e) => {
    const chip = e.target.closest("[data-key]");
    if (!chip) return;
    state.motion = chip.dataset.key;
    renderMotions();
    updatePreview();
  });
  $("humanize").addEventListener("change", () => { state.humanize = $("humanize").checked; });
  bindUpload("crest");
  bindUpload("logo");
  $("btnNext").addEventListener("click", () => goStep(state.step + 1));
  $("btnBack").addEventListener("click", () => goStep(state.step - 1));
  $("btnSubmit").addEventListener("click", submit);
})();
</script>
</body>
</html>
"""


def build_html(catalog: Catalog) -> str:
    data = catalog.page_data
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    return PAGE_TEMPLATE.replace("__DATA_JSON__", payload)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def default_skill_dir() -> str:
    """脚本在 <skill>/scripts/ 下时自动推出技能根。"""
    parent = os.path.dirname(SCRIPT_DIR)
    if os.path.isdir(os.path.join(parent, "references")):
        return parent
    return SCRIPT_DIR


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="profile_wizard.py",
        description="学院汇报 · 模板匹配向导：本地网页收集学校/姓名/身份/目的/时长/皮肤/校徽，输出 profile.json",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python3 profile_wizard.py --out profile.json\n"
            "  python3 profile_wizard.py --no-open --port 8765 --out work/wizard/profile.json\n"
            "  python3 profile_wizard.py --plain --out profile.json\n"
            "  python3 profile_wizard.py --plain --answers answers.json --out profile.json\n"
        ),
    )
    parser.add_argument("--skill-dir", default=None, help="技能根目录（默认自动探测脚本上一级）")
    parser.add_argument("--out", required=True, help="输出 profile.json 路径")
    parser.add_argument("--port", type=int, default=0, help="本地端口（默认 0 = 随机；被占用自动顺延）")
    parser.add_argument("--no-open", action="store_true", help="不自动打开浏览器，只打印 URL")
    parser.add_argument("--timeout", type=int, default=900, help="等待超时秒数（默认 900）")
    parser.add_argument("--plain", action="store_true", help="纯文本模式：不开浏览器，问题逐条打印")
    parser.add_argument("--answers", default=None, help="answers.json 路径；给了就直接生成 profile")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    skill_dir = os.path.abspath(os.path.expanduser(args.skill_dir or default_skill_dir()))
    if not os.path.isdir(skill_dir):
        warn("技能目录不存在：%s" % skill_dir)
        return 2
    catalog = Catalog(skill_dir)

    if args.plain or args.answers:
        return run_plain(args, catalog)
    return run_http(args, catalog)


if __name__ == "__main__":
    raise SystemExit(main())

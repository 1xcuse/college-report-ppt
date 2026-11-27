#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""verify_claims.py — content.json 内容回查（生成 PPT 前的强制关卡）。

用法：
    python3 verify_claims.py --content content.json --paper work/xxx/paper/paper.md \
        [--figures work/xxx/figures/figures.json] [--json report.json] [--strict] \
        [--source-doc metrics_source.md]

检查五件事：
  1) 数字 / 单位 / 百分比 / 英文术语 / 专有名词能否在论文原文找到
     （全半角、空格、连字符归一化；数字容差 0；「论文没有写」句式豁免）；
  2) paper_card 必填字段齐不齐（题名/作者/摘要概述/期刊/DOI/分区来源说明）；
  3) 文案长度是否超过 format_spec 上限（标题 22、k 12、t 62、notes 140、图注 40、表格 28…）；
  4) 图注编号与 figures.json 是否对得上（文件存在、编号不冲突、子图对应）；
  5) 标题是否空泛（「研究背景」「方法」这类空标题判为提示）。

输出：按页面分组的表（页号/字段/问题/候选证据原文片段/建议）；`--strict` 有 error 时 exit 1。
JSON 报告：`--json report.json`。

注意：`paper.ranks / paper.metrics / paper.volume / paper.doi` 这类「期刊外部数据」
不是论文正文能核验的，需要 `--source-doc` 提供现查来源（LetPub/期刊官网的本地记录），
或至少写 `paper.source_note`；否则只给提示，不判 error。
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from paper_extract import normalize_text, normalize_key, clean_inline  # noqa: E402
except Exception:  # pragma: no cover
    def normalize_text(s):
        return re.sub(r"\s+", " ", str(s or "")).strip()

    def normalize_key(s):
        return re.sub(r"[\s\-_/]+", "", normalize_text(s).lower())

    def clean_inline(s):
        return normalize_text(s)


# ---------------------------------------------------------------------------
# format_spec 上限（与 references/format_spec.md 保持一致）
# ---------------------------------------------------------------------------

LIMITS = {
    "title": 22,
    "eyebrow": 16,
    "lead": 48,
    "k": 12,
    "t": 62,
    "notes": 140,
    "table_cell": 28,
    "caption": 40,
    "paper_summary": 48,
    "keynum_num": 6,
    "keynum_unit": 6,
}

# 标题空泛白名单（写了等于没写）
VAGUE_TITLES = {
    "研究背景", "背景", "研究内容", "内容", "方法", "研究方法", "研究思路", "思路",
    "模型", "系统模型", "问题建模", "问题描述", "实验", "实验部分", "实验结果",
    "实验设置", "结果", "结论", "总结", "总结与展望", "引言", "介绍", "论文简介",
    "相关工作", "主要工作", "技术路线", "创新点", "贡献", "展望", "论文方法",
    "论文结果", "系统设计", "算法设计", "数据分析", "评估", "讨论",
}

# 期刊外部指标（影响因子/分区）不是论文正文能核验的；keynums 里允许出现，
# 但需要 --source-doc 来源，否则只给提示
EXTERNAL_CLAIM_KW = ["影响因子", "分区", "JCR", "CCF", "中科院", "CiteScore", "SJR",
                     "H-index", "H 指数", "IF ", " Impact Factor"]

EXEMPT_PHRASES = [
    "论文没有写", "论文没写", "论文未写", "论文没有给出", "论文没给", "论文未给出",
    "原文没有写", "原文没有给出", "原文未给出", "原文没有", "论文未报告", "未报告",
    "没有给出", "未给出", "没有提供", "未提供", "没有说明", "未说明",
]

# 英文虚词 / 常见词（不做原文回查）
EN_STOP = {
    "a", "an", "the", "of", "and", "or", "for", "with", "to", "in", "on", "by", "as", "at",
    "from", "via", "per", "vs", "versus", "is", "are", "was", "were", "be", "been", "being",
    "this", "that", "these", "those", "we", "our", "it", "its", "they", "their", "than",
    "then", "so", "such", "both", "all", "each", "more", "less", "high", "low", "new",
    "can", "may", "not", "no", "up", "down", "out", "under", "over", "between", "within",
    "into", "two", "three", "four", "one", "zero", "first", "second", "third", "also",
    "based", "using", "used", "use", "when", "where", "which", "what", "how", "why",
    "if", "but", "while", "about", "after", "before", "during", "only", "very", "much",
    "many", "some", "any", "every", "other", "another", "same", "different", "due",
    "e", "g", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x", "et", "al",
    "fig", "figure", "figs", "table", "tab", "eq", "eqs", "section", "sec", "ref", "refs",
    "ieee", "vol", "no", "pp", "doi",
}

# 中文数字单位 → 倍率
CN_UNIT_SCALE = {"万": 10000.0, "亿": 100000000.0, "千": 1000.0}

UNIT_ALIASES = {
    "％": "%", "‰": "‰",
    "米": "m", "千米": "km", "公里": "km", "厘米": "cm", "毫米": "mm",
    "秒": "s", "毫秒": "ms", "分钟": "min", "小时": "h",
    "分贝": "dB", "分贝毫瓦": "dBm",
    "赫兹": "Hz", "千赫兹": "kHz", "兆赫兹": "MHz", "吉赫兹": "GHz",
    "度": "°", "度每秒": "°/s",
    "倍": "倍", "次": "次", "个": "个", "条": "条", "架": "架", "台": "台",
    "维": "维", "层": "层", "轮": "轮", "回合": "回合", "时隙": "slot",
}

UNIT_PATTERN = (
    "%|‰|dBm|dBW|dBi|dB|bps/Hz|bit/s|bps|km/h|m/s|ms|min|h|km|cm|mm|m|s|"
    "GHz|MHz|kHz|Hz|THz|万|亿|千|倍|维|层|条|个|架|台|次|度|轮|回合|时隙|°"
)

NUM_REGEX = re.compile(
    r"(?<![A-Za-z0-9.])("
    r"\d{1,3}(?:,\d{3})+(?:\.\d+)?"
    r"|\d+(?:\.\d+)?"
    r")(?![A-Za-z0-9])(?:\s*(" + UNIT_PATTERN + r"))?"
)

FORMULA_NO_RE = re.compile(r"[（(]\s*(\d{1,2}[a-z]?)\s*[)）]")
# 只去掉「1.1　标题」这种编号，避免把表格里的 "80,000"、"40" 当成编号削掉
SECTION_NUM_RE = re.compile(
    r"^\s*(?:(?:\d+(?:\.\d+)+)\s+|\d{1,2}[.、．:：]\s+|\d{1,2}[)）]\s*"
    r"|[一二三四五六七八九十]+[、.．]\s*)"
)
FIG_NUM_PREFIX_RE = re.compile(
    r"^\s*(?:图|Fig\.?|Figure|表|Table)\s*\d+(?:\.\d+)?\s*(?:[（(][a-zA-Z][)）])?\s*[.．、:：]?\s*"
)


# ---------------------------------------------------------------------------
# 数字 / 英文术语抽取与归一化
# ---------------------------------------------------------------------------

def canon_number(s: str) -> str:
    s = s.replace(",", "").strip()
    if not s:
        return ""
    if "." in s:
        s = s.rstrip("0").rstrip(".")
        if s.startswith("."):
            s = "0" + s
        return s
    s = s.lstrip("0")
    return s if s else "0"


def norm_unit(u: str):
    if not u:
        return ""
    u = u.strip()
    return UNIT_ALIASES.get(u, u)


def extract_numbers(text):
    """返回 [(value:str, unit:str, raw:str)]，已做全半角/单位归一化。"""
    out = []
    for m in NUM_REGEX.finditer(text or ""):
        raw_num, raw_unit = m.group(1), m.group(2)
        value = canon_number(raw_num)
        unit = norm_unit(raw_unit)
        if not value:
            continue
        if unit in CN_UNIT_SCALE:
            try:
                value = canon_number(str(float(value) * CN_UNIT_SCALE[unit]))
            except Exception:
                pass
            unit = ""
        out.append((value, unit, m.group(0).strip()))
    return out


def extract_en_tokens(text):
    """英文术语 / 专有名词（去掉 HTML 标签与纯数字）。"""
    if not text:
        return []
    t = re.sub(r"<[^>]+>", " ", text)
    toks = []
    for m in re.finditer(r"[A-Za-z][A-Za-z0-9]*(?:[-–_][A-Za-z0-9]+)*", t):
        w = m.group(0)
        low = w.lower()
        if len(w) < 2:
            continue
        if low in EN_STOP:
            continue
        if re.fullmatch(r"[ivxIVX]+", w):  # 罗马数字章节号
            continue
        toks.append(w)
    return toks


def split_sentences(text):
    return [s for s in re.split(r"(?<=[。；;！!？?])", text or "") if s.strip()]


def has_exemption(text):
    return any(p in (text or "") for p in EXEMPT_PHRASES)


def strip_numbering(text, allow_pure_numbering=False):
    t = FIG_NUM_PREFIX_RE.sub("", text or "")
    t = SECTION_NUM_RE.sub("", t)
    t = FORMULA_NO_RE.sub(" ", t)  # (6a) 这类公式编号
    # eyebrow 常见写法是只有编号 "2.1"：编号是结构，不是数据。
    # 但表格单元格 / keynums 里的纯数字是数据，不能豁免。
    if allow_pure_numbering and re.fullmatch(r"[\d\s.．、:：()（）-]+", t or ""):
        return ""
    return t


# ---------------------------------------------------------------------------
# 论文正文索引
# ---------------------------------------------------------------------------

class PaperIndex:
    def __init__(self, md_path):
        self.md_path = md_path
        with open(md_path, "r", encoding="utf-8", errors="replace") as f:
            raw = f.read()
        self.raw = raw
        # 按 <!-- page:N --> 切页
        parts = re.split(r"<!--\s*page\s*:\s*(\d+)\s*-->", raw)
        self.pages = []  # [(page_no, text)]
        if len(parts) > 1:
            preface = parts[0]
            if preface.strip():
                self.pages.append((1, preface))
            for i in range(1, len(parts), 2):
                pno = int(parts[i])
                text = parts[i + 1] if i + 1 < len(parts) else ""
                self.pages.append((pno, text))
        else:
            self.pages = [(1, raw)]
        self.page_text = {}
        for pno, text in self.pages:
            self.page_text[pno] = self.page_text.get(pno, "") + "\n" + text
        self.full = "\n".join(t for _, t in self.pages)
        # 去掉 markdown 修饰，保留可匹配文本
        plain = re.sub(r"<!--.*?-->", " ", self.full, flags=re.S)
        plain = re.sub(r"<[^>]+>", " ", plain)
        plain = re.sub(r"^#{1,6}\s*", "", plain, flags=re.M)
        plain = plain.replace("**", " ").replace("__", " ")
        plain = re.sub(r"^\s*[>|]\s?", " ", plain, flags=re.M)
        self.plain = plain
        self.norm = normalize_text(plain)
        self.lower = self.norm.lower()
        self.key = normalize_key(plain)
        self._num_index = None
        self._token_index = None

    # --- 数字索引 ---
    @property
    def num_index(self):
        if self._num_index is None:
            idx = {}
            for pno, text in self.pages:
                t = normalize_text(text)
                for value, unit, raw in extract_numbers(t):
                    idx.setdefault(value, []).append((pno, unit, raw))
            self._num_index = idx
        return self._num_index

    def find_number(self, value, unit):
        """返回 (是否找到, 论文单位, [(page, snippet)])。"""
        hits = self.num_index.get(value, [])
        if not hits:
            # 允许 "5 万" 归一化后与原值 50000 之外的写法：按数值近似不再放宽（容差 0）
            return False, "", []
        same_unit = [h for h in hits if norm_unit(h[1]) == norm_unit(unit)]
        pick = same_unit or [h for h in hits if not h[1] or not unit]
        if not pick:
            pick = hits[:1]
        ev = []
        for pno, u, raw in pick[:2]:
            ev.append({"page": pno, "snippet": self.snippet("%s" % raw, pno)})
        paper_unit = pick[0][1]
        return True, paper_unit, ev

    def snippet(self, needle, page=None, span=110):
        if page is not None and page in self.page_text:
            hay = self.page_text[page]
        else:
            hay = self.full
        i = hay.find(needle)
        if i < 0:
            # 用归一化文本再找
            j = normalize_text(hay).find(needle)
            if j < 0:
                return ""
            i = max(0, j - span // 3)
        start = max(0, i - span // 2)
        end = min(len(hay), i + len(needle) + span // 2)
        s = re.sub(r"\s+", " ", hay[start:end]).strip()
        return s

    # --- 英文术语索引 ---
    @property
    def token_index(self):
        if self._token_index is None:
            idx = {}
            for pno, text in self.pages:
                t = normalize_text(text)
                for w in re.finditer(r"[A-Za-z][A-Za-z0-9]*(?:[-–][A-Za-z0-9]+)*", t):
                    idx.setdefault(w.group(0).lower(), pno)
            self._token_index = idx
        return self._token_index

    def find_token(self, token):
        """返回 (是否找到, page, 证据片段)。"""
        tok = token.lower()
        # 下标写法差异：P_max / Pmax / P max / P-max 都算同一个术语
        base = re.sub(r"[_\-–]", " ", tok)
        compact = re.sub(r"[\s_\-–]+", "", tok)
        variants = {tok, base, compact, tok.replace("_", "")}
        for v in variants:
            if not v:
                continue
            if re.search(r"(?<![a-z0-9])" + re.escape(v) + r"(?![a-z0-9])", self.lower):
                page = self._find_token_page(v)
                return True, page, self.snippet(token)
        # 复数 / 去复数
        for v in list(variants):
            for alt in (v + "s", v[:-1] if v.endswith("s") else ""):
                if alt and re.search(r"(?<![a-z0-9])" + re.escape(alt) + r"(?![a-z0-9])", self.lower):
                    return True, self._find_token_page(alt), self.snippet(token)
        # 无空格全文中兜底（只对较长术语）
        if len(compact) >= 5 and compact in self.key:
            page = self.token_index.get(tok)
            return True, page, self.snippet(token)
        return False, None, ""

    def _find_token_page(self, v):
        page = self.token_index.get(v)
        if page is not None:
            return page
        m = re.search(r"<!--\s*page\s*:\s*(\d+)\s*-->", self.raw)
        return int(m.group(1)) if m else None

    def similar_numbers(self, value, limit=3):
        """给错误数字找相近候选（整数部分相同或数值接近）。"""
        try:
            v = float(value)
        except ValueError:
            return []
        cands = []
        for val, hits in self.num_index.items():
            try:
                fv = float(val)
            except ValueError:
                continue
            if fv == v:
                continue
            if val.split(".")[0] == value.split(".")[0] or (v and abs(fv - v) / max(abs(v), 1e-9) < 0.15):
                cands.append((abs(fv - v), val, hits[0][0]))
        cands.sort()
        out = []
        for _, val, pno in cands[:limit]:
            out.append({"value": val, "page": pno, "snippet": self.snippet(val, pno)})
        return out

    def similar_tokens(self, token, limit=2):
        low = token.lower()
        out = []
        for tok, pno in self.token_index.items():
            if len(tok) < 4:
                continue
            if tok.startswith(low[:4]) or (len(low) >= 5 and low[:4] in tok):
                out.append({"token": tok, "page": pno, "snippet": self.snippet(tok, pno)})
                if len(out) >= limit:
                    break
        return out


# ---------------------------------------------------------------------------
# content.json 页面编号 / 字段遍历
# ---------------------------------------------------------------------------

def resolve_page_numbers(content):
    pages = content.get("pages") or []
    paper = content.get("paper") or {}
    sections = content.get("sections") or []
    page_types = [p.get("type") for p in pages]
    has_paper = bool(paper.get("title_cn") or paper.get("title_en"))
    cover_explicit = bool(pages) and pages[0].get("type") == "cover"
    body_count = sum(1 for t in page_types if t not in ("cover", "paper_card", "toc"))
    toc_needed = bool(sections) and body_count >= 6
    auto_pc = has_paper and "paper_card" not in page_types
    auto_toc = toc_needed and "toc" not in page_types

    numbers = []
    ptr = 1
    cover_done = False
    pc_done = "paper_card" in page_types
    toc_done = "toc" in page_types
    pc_number = None
    toc_number = None
    for p in pages:
        t = p.get("type", "")
        if t == "cover":
            numbers.append(1)
            ptr = 2
            cover_done = True
            continue
        if not cover_done:
            ptr = 2  # 自动封面占第 1 页
            cover_done = True
        if t == "paper_card":
            numbers.append(ptr)
            pc_number = ptr
            ptr += 1
            pc_done = True
            continue
        if t == "toc":
            numbers.append(ptr)
            toc_number = ptr
            ptr += 1
            toc_done = True
            continue
        if has_paper and not pc_done:
            ptr += 1
            pc_done = True
            if pc_number is None:
                pc_number = ptr - 1
        if toc_needed and not toc_done:
            ptr += 1
            toc_done = True
            if toc_number is None:
                toc_number = ptr - 1
        numbers.append(ptr)
        ptr += 1
    if pc_number is None:
        pc_number = 2 if has_paper else None
    return {
        "numbers": numbers,
        "paper_card_page": pc_number,
        "toc_page": toc_number,
        "auto_prefix": (1 if not cover_explicit else 0) + (1 if auto_pc else 0) + (1 if auto_toc else 0),
        "body_count": body_count,
    }


def iter_text_fields(page, path="page"):
    """遍历一页里需要回查的文本字段：[(field_path, text, kind)]。

    kind: title/eyebrow/lead/item_k/item_t/block_h/block_p/table_cell/tnote/
          caption/line/col_item/step/qa/roadmap/quote/formula/notes
    """
    t = page.get("type", "")
    out = []
    if page.get("title"):
        out.append(("title", page["title"], "title"))
    if page.get("eyebrow"):
        out.append(("eyebrow", page["eyebrow"], "eyebrow"))
    if page.get("lead"):
        out.append(("lead", page["lead"], "lead"))
    for i, it in enumerate(page.get("items") or []):
        if isinstance(it, dict):
            if it.get("k"):
                out.append(("items[%d].k" % i, it["k"], "item_k"))
            if it.get("t"):
                out.append(("items[%d].t" % i, it["t"], "item_t"))
            if it.get("num"):
                ctx = " ".join(str(it.get(x) or "") for x in ("t", "unit"))
                ext = any(k in ctx for k in EXTERNAL_CLAIM_KW)
                out.append(("items[%d].num" % i, it["num"], "keynum_num_ext" if ext else "keynum_num"))
            if it.get("unit"):
                ctx = " ".join(str(it.get(x) or "") for x in ("t", "num"))
                ext = any(k in ctx for k in EXTERNAL_CLAIM_KW)
                out.append(("items[%d].unit" % i, it["unit"], "keynum_unit_ext" if ext else "keynum_unit"))
            if it.get("h"):
                out.append(("items[%d].h" % i, it["h"], "block_h"))
            if it.get("p"):
                out.append(("items[%d].p" % i, it["p"], "block_p"))
        elif isinstance(it, str):
            out.append(("items[%d]" % i, it, "item_t"))
    for i, b in enumerate(page.get("blocks") or []):
        if isinstance(b, dict):
            if b.get("h"):
                out.append(("blocks[%d].h" % i, b["h"], "block_h"))
            if b.get("p"):
                out.append(("blocks[%d].p" % i, b["p"], "block_p"))
    for ri, row in enumerate(page.get("rows") or []):
        for ci, cell in enumerate(row if isinstance(row, list) else [row]):
            if cell is not None and str(cell).strip():
                out.append(("rows[%d][%d]" % (ri, ci), str(cell), "table_cell"))
    for key in ("tnote", "note"):
        if page.get(key):
            out.append((key, page[key], "tnote"))
    if page.get("caption"):
        out.append(("caption", page["caption"], "caption"))
    if page.get("caption2"):
        out.append(("caption2", page["caption2"], "caption"))
    for i, f in enumerate(page.get("figs") or []):
        if isinstance(f, dict) and f.get("caption"):
            out.append(("figs[%d].caption" % i, f["caption"], "caption"))
    for i, line in enumerate(page.get("lines") or []):
        if isinstance(line, str) and line.strip():
            out.append(("lines[%d]" % i, line, "line"))
    side = page.get("side") or {}
    for i, it in enumerate(side.get("items") or []):
        if isinstance(it, dict):
            if it.get("k"):
                out.append(("side.items[%d].k" % i, it["k"], "item_k"))
            if it.get("t") or it.get("d"):
                out.append(("side.items[%d].t" % i, it.get("t") or it.get("d"), "item_t"))
            if it.get("h"):
                out.append(("side.items[%d].h" % i, it["h"], "block_h"))
    for i, col in enumerate(page.get("cols") or []):
        if not isinstance(col, dict):
            continue
        if col.get("h"):
            out.append(("cols[%d].h" % i, col["h"], "block_h"))
        for j, it in enumerate(col.get("items") or []):
            if isinstance(it, str):
                out.append(("cols[%d].items[%d]" % (i, j), it, "col_item"))
    for i, st in enumerate(page.get("steps") or []):
        if isinstance(st, dict):
            if st.get("h"):
                out.append(("steps[%d].h" % i, st["h"], "block_h"))
            if st.get("t"):
                out.append(("steps[%d].t" % i, st["t"], "item_t"))
    for i, it in enumerate(page.get("items") or []):
        if isinstance(it, dict) and it.get("q"):
            out.append(("items[%d].q" % i, it["q"], "qa"))
            if it.get("a"):
                out.append(("items[%d].a" % i, it["a"], "qa"))
    if page.get("text"):
        out.append(("text", page["text"], "quote"))
    if page.get("from"):
        out.append(("from", page["from"], "quote"))
    if page.get("num") and t == "chapter":
        pass
    if page.get("sub"):
        out.append(("sub", page["sub"], "lead"))
    if t == "chapter":
        for i, b in enumerate(page.get("bullets") or []):
            if isinstance(b, str):
                out.append(("bullets[%d]" % i, b, "item_t"))
    if page.get("formulas"):
        for i, fm in enumerate(page["formulas"]):
            if isinstance(fm, dict) and fm.get("t"):
                out.append(("formulas[%d].t" % i, fm["t"], "formula"))
    if page.get("notes"):
        out.append(("notes", page["notes"], "notes"))
    return out


# ---------------------------------------------------------------------------
# 检查器
# ---------------------------------------------------------------------------

class Checker:
    def __init__(self, content, paper_index, figures_doc=None, source_docs=None):
        self.content = content
        self.paper_index = paper_index
        self.paper_card = content.get("paper") or {}
        self.figures = (figures_doc or {}).get("figures") or []
        self.source_text = "\n".join(source_docs or [])
        self.source_norm = normalize_text(self.source_text)
        self.source_lower = self.source_norm.lower()
        self.source_key = normalize_key(self.source_text)
        self.page_map = resolve_page_numbers(content)
        self.issues = []
        self.coverage = {
            "fields_checked": 0,
            "numbers_checked": 0,
            "numbers_grounded": 0,
            "en_tokens_checked": 0,
            "en_tokens_grounded": 0,
            "figures_checked": 0,
            "pages_checked": 0,
        }
        self._fig_no_map = {}

    # ----- 基础 -----
    def add(self, page_num, page_type, field, severity, category, message,
            evidence="", suggestion=""):
        self.issues.append(
            {
                "page": page_num,
                "page_type": page_type,
                "field": field,
                "severity": severity,
                "category": category,
                "message": message,
                "evidence": evidence,
                "suggestion": suggestion,
            }
        )

    # ----- 1. paper_card -----
    def check_paper_card(self):
        pc = self.paper_card
        pno = self.page_map.get("paper_card_page")
        if not pc:
            self.add(pno, "paper_card", "paper", "warn", "字段",
                     "content.json 没有 paper 对象；文献汇报会缺少论文简介页",
                     suggestion="补 paper.title_cn / authors / summary / venue / doi")
            return
        has_cover = bool(self.content.get("pages")) and self.content["pages"][0].get("type") == "cover"
        if pno is None:
            pno = 2 if has_cover or True else 2
        title = pc.get("title_cn") or pc.get("title_en") or ""
        fields = [
            ("title_cn/title_en", title, "题名"),
            ("authors", pc.get("authors"), "作者"),
            ("summary", pc.get("summary"), "一句话概述"),
            ("venue/journal", pc.get("venue") or pc.get("journal"), "期刊/会议"),
            ("doi", pc.get("doi"), "DOI"),
        ]
        for f, v, name in fields:
            if not (str(v).strip() if v is not None else ""):
                self.add(pno, "paper_card", "paper." + f, "error", "缺少字段",
                         "paper.%s（%s）为空；论文简介页会缺内容" % (f, name),
                         suggestion="从 paper_card.json 补齐")
        for f, name in [("volume", "卷期页码"), ("affiliations", "单位"),
                        ("keywords", "关键词"), ("ranks", "分区"), ("metrics", "影响因子")]:
            if not pc.get(f):
                self.add(pno, "paper_card", "paper." + f, "hint", "缺少字段",
                         "paper.%s（%s）为空（非必填，但讲解稿/简介页会用）" % (f, name))
        # summary 长度
        if pc.get("summary"):
            n = len(clean_inline(pc["summary"]))
            self.coverage["fields_checked"] += 1
            if n > LIMITS["paper_summary"]:
                self.add(pno, "paper_card", "paper.summary", "error", "文案长度",
                         "paper.summary 长度 %d 字 > 上限 %d 字" % (n, LIMITS["paper_summary"]),
                         suggestion="压成一句话结论")
        # 分区/影响因子必须有来源
        if pc.get("ranks") or pc.get("metrics"):
            if not pc.get("source_note"):
                self.add(pno, "paper_card", "paper.source_note", "error", "缺少字段",
                         "写了 rankings/metrics 但没有 source_note（分区/影响因子必须现查并写来源）",
                         suggestion="补 source_note，如：数据来源 LetPub / 期刊官网（查询日期 2026-11）")
            else:
                self._check_external_metadata(pno, pc)
        # 作者是否能在论文里找到
        authors = pc.get("authors")
        if isinstance(authors, list):
            authors = "、".join(str(a) for a in authors)
        if authors:
            names = [x.strip() for x in re.split(r"[、,;；/]| and ", str(authors)) if x.strip()]
            for name in names:
                name = re.sub(r"\s*(?:Member|Fellow|Senior|Student|IEEE).*$", "", name).strip()
                if not name:
                    continue
                self.coverage["fields_checked"] += 1
                toks = [w for w in re.findall(r"[A-Za-z][A-Za-z\-]+", name) if len(w) >= 3]
                if toks and not all(self._find_in_paper(w, is_name=True) for w in toks):
                    self.add(pno, "paper_card", "paper.authors", "warn", "术语",
                             "作者「%s」在论文原文里没找到（可能姓在前/缩写）" % name,
                             suggestion="核对作者拼写；英文名以论文首页为准")
        # paper.figures 里的文件是否存在
        for f in pc.get("figures") or []:
            fn = f.get("file") if isinstance(f, dict) else None
            if fn and self.figures and not self._find_figure(fn):
                self.add(pno, "paper_card", "paper.figures", "warn", "图注",
                         "paper.figures 里的 %s 不在 figures.json（抽图结果）里" % fn,
                         suggestion="跑 extract_figures.py 或改成实际文件名")
        # contributions / limits 的英文术语回查
        for key in ("contributions", "limits"):
            for i, it in enumerate(pc.get(key) or []):
                if isinstance(it, dict):
                    for sub in ("k", "t"):
                        if it.get(sub):
                            self._check_claim_text(
                                pno, "paper_card", "paper.%s[%d].%s" % (key, i, sub), it[sub],
                                "item_k" if sub == "k" else "item_t")

    def _check_external_metadata(self, pno, pc):
        """paper.ranks / metrics / metrics_note 这类期刊外部数据：用 --source-doc 核验。

        volume / doi / venue 来自 PDF metadata 与页眉，不作为正文数据回查对象。
        """
        texts = []
        for key in ("ranks", "metrics", "metrics_note"):
            v = pc.get(key)
            if isinstance(v, list):
                texts.extend(str(x) for x in v)
            elif v:
                texts.append(str(v))
        for t in texts:
            if not t:
                continue
            if self.source_text:
                self._check_text_against_source(pno, "paper.meta", t)
            else:
                self.add(pno, "paper_card", "paper.metrics", "hint", "外部数据",
                         "「%s」属于期刊外部数据，论文原文无法核验；建议用 --source-doc 提供现查来源" % t[:40],
                         suggestion="把 LetPub/期刊官网记录存成 md 后 --source-doc 传进来")

    def _check_text_against_source(self, pno, field, text):
        self.coverage["fields_checked"] += 1
        clean = strip_numbering(re.sub(r"<[^>]+>", " ", text))
        for value, unit, raw in extract_numbers(clean):
            self.coverage["numbers_checked"] += 1
            ok = self._source_has_number(value, unit)
            if ok:
                self.coverage["numbers_grounded"] += 1
            else:
                self.add(pno, "paper_card", field, "error", "数字",
                         "外部数据里的数字「%s」在 --source-doc 里找不到" % raw,
                         self._source_snippet(),
                         suggestion="核对来源文件；论文没有写的数字不要写成来源数据")
        for tok in extract_en_tokens(clean):
            if len(tok) < 3:
                continue
            self.coverage["en_tokens_checked"] += 1
            ok = bool(re.search(r"(?<![A-Za-z0-9])" + re.escape(tok.lower()) + r"(?![A-Za-z0-9])", self.source_lower)) or (
                len(normalize_key(tok)) >= 5 and normalize_key(tok) in self.source_key
            )
            if ok:
                self.coverage["en_tokens_grounded"] += 1
            else:
                self.add(pno, "paper_card", field, "error", "术语",
                         "外部数据里的术语「%s」在 --source-doc 里找不到" % tok,
                         self._source_snippet(),
                         suggestion="核对来源文件拼写")

    def _source_has_number(self, value, unit):
        if not self.source_text:
            return False
        for v, u, raw in extract_numbers(self.source_norm):
            if v == value and (not unit or not u or norm_unit(u) == norm_unit(unit)):
                return True
        return False

    def _source_snippet(self):
        if not self.source_text:
            return ""
        return re.sub(r"\s+", " ", self.source_text)[:140]

    def _find_in_paper(self, word, is_name=False):
        if not word:
            return True
        ok, _, _ = self.paper_index.find_token(word)
        return ok

    def _find_figure(self, filename):
        bn = os.path.basename(filename)
        for f in self.figures:
            if os.path.basename(str(f.get("file", ""))) == bn:
                return f
        return None

    # ----- 2. 文案长度 -----
    def check_lengths(self, page, pno):
        t = page.get("type", "")
        if page.get("title"):
            n = len(clean_inline(page["title"]))
            self.coverage["fields_checked"] += 1
            if n > LIMITS["title"]:
                self.add(pno, t, "title", "error", "文案长度",
                         "title 长度 %d 字 > 上限 %d 字" % (n, LIMITS["title"]),
                         suggestion="改成一句结论，≤%d 字" % LIMITS["title"])
        if page.get("eyebrow"):
            n = len(clean_inline(page["eyebrow"]))
            self.coverage["fields_checked"] += 1
            if n > LIMITS["eyebrow"]:
                self.add(pno, t, "eyebrow", "error", "文案长度",
                         "eyebrow 长度 %d 字 > 上限 %d 字" % (n, LIMITS["eyebrow"]),
                         suggestion="只保留编号式小标题")
        if page.get("lead"):
            n = len(clean_inline(page["lead"]))
            self.coverage["fields_checked"] += 1
            if n > LIMITS["lead"]:
                self.add(pno, t, "lead", "error", "文案长度",
                         "lead 长度 %d 字 > 上限 %d 字" % (n, LIMITS["lead"]),
                         suggestion="压成一句结论")
        for i, it in enumerate(page.get("items") or []):
            if not isinstance(it, dict):
                continue
            if it.get("k"):
                n = len(clean_inline(it["k"]))
                self.coverage["fields_checked"] += 1
                if n > LIMITS["k"]:
                    self.add(pno, t, "items[%d].k" % i, "error", "文案长度",
                             "要点核心词长度 %d 字 > 上限 %d 字" % (n, LIMITS["k"]),
                             suggestion="压缩成 10 字内的关键词")
            if it.get("t"):
                n = len(clean_inline(it["t"]))
                self.coverage["fields_checked"] += 1
                if n > LIMITS["t"]:
                    self.add(pno, t, "items[%d].t" % i, "error", "文案长度",
                             "要点说明长度 %d 字 > 上限 %d 字" % (n, LIMITS["t"]),
                             suggestion="一句话说明，≤62 字（超了换行）")
        if page.get("rows"):
            for ri, row in enumerate(page["rows"]):
                for ci, cell in enumerate(row if isinstance(row, list) else [row]):
                    if cell is None:
                        continue
                    n = len(clean_inline(str(cell)))
                    self.coverage["fields_checked"] += 1
                    if n > LIMITS["table_cell"]:
                        self.add(pno, t, "rows[%d][%d]" % (ri, ci), "warn", "文案长度",
                                 "表格单元格 %d 字 > 上限 %d 字（生成器可能溢出）" % (n, LIMITS["table_cell"]),
                                 suggestion="拆成短语或换行")
        for key in ("caption", "caption2"):
            if page.get(key):
                n = len(clean_inline(page[key]))
                self.coverage["fields_checked"] += 1
                if n > LIMITS["caption"]:
                    self.add(pno, t, key, "error", "文案长度",
                             "图注 %d 字 > 上限 %d 字" % (n, LIMITS["caption"]),
                             suggestion="图注只写图是什么，不下结论")
        for i, f in enumerate(page.get("figs") or []):
            if isinstance(f, dict) and f.get("caption"):
                n = len(clean_inline(f["caption"]))
                self.coverage["fields_checked"] += 1
                if n > LIMITS["caption"]:
                    self.add(pno, t, "figs[%d].caption" % i, "error", "文案长度",
                             "图注 %d 字 > 上限 %d 字" % (n, LIMITS["caption"]))
        if page.get("notes"):
            n = len(clean_inline(page["notes"]))
            self.coverage["fields_checked"] += 1
            if n > LIMITS["notes"]:
                self.add(pno, t, "notes", "warn", "文案长度",
                         "口头讲稿 %d 字 > 上限 %d 字" % (n, LIMITS["notes"]),
                         suggestion="压到 1–3 句短句")
        # 数量约束（提示级）
        if t == "bullets" and page.get("items"):
            n = len(page["items"])
            if not 3 <= n <= 5:
                self.add(pno, t, "items", "hint", "结构",
                         "bullets 页建议 3–5 条要点，当前 %d 条" % n)
        if t == "keynums" and page.get("items"):
            n = len(page["items"])
            if not 2 <= n <= 4:
                self.add(pno, t, "items", "hint", "结构",
                         "keynums 页建议 2–4 个数字，当前 %d 个" % n)
            for i, it in enumerate(page["items"]):
                if isinstance(it, dict):
                    if it.get("num") and len(str(it["num"])) > LIMITS["keynum_num"]:
                        self.add(pno, t, "items[%d].num" % i, "warn", "文案长度",
                                 "大数字 %r 超过 %d 字符" % (it["num"], LIMITS["keynum_num"]))
                    if it.get("unit") and len(str(it["unit"])) > LIMITS["keynum_unit"]:
                        self.add(pno, t, "items[%d].unit" % i, "warn", "文案长度",
                                 "单位 %r 超过 %d 字符" % (it["unit"], LIMITS["keynum_unit"]))
        if t == "text" and page.get("blocks"):
            if len(page["blocks"]) > 4:
                self.add(pno, t, "blocks", "hint", "结构", "text 页最多 4 块，当前 %d 块" % len(page["blocks"]))
        if t == "chapter" and page.get("bullets"):
            if len(page["bullets"]) > 4:
                self.add(pno, t, "bullets", "hint", "结构", "chapter 页 bullets 最多 4 条")
        if t == "compare" and page.get("cols"):
            if not 2 <= len(page["cols"]) <= 3:
                self.add(pno, t, "cols", "hint", "结构", "compare 页建议 2–3 列")
        if t in ("timeline", "pipeline") and page.get("steps"):
            if not 3 <= len(page["steps"]) <= 5:
                self.add(pno, t, "steps", "hint", "结构", "%s 建议 3–5 步，当前 %d 步" % (t, len(page["steps"])))
        if t == "summary" and page.get("cols"):
            for i, col in enumerate(page["cols"]):
                for j, it in enumerate((col or {}).get("items") or []):
                    n = len(clean_inline(str(it)))
                    self.coverage["fields_checked"] += 1
                    if n > LIMITS["t"]:
                        self.add(pno, t, "cols[%d].items[%d]" % (i, j), "warn", "文案长度",
                                 "summary 条目 %d 字 > %d 字，建议精简" % (n, LIMITS["t"]))

    # ----- 3. 数字 / 英文术语回查 -----
    def _is_external_claim(self, text):
        t = text or ""
        if any(k in t for k in EXTERNAL_CLAIM_KW):
            return True
        return False

    def _check_claim_text(self, pno, page_type, field, text, kind):
        if not text:
            return
        text = re.sub(r"<[^>]+>", " ", str(text))
        # 分句，豁免「论文没有写」句式
        clauses = split_sentences(text)
        if not clauses:
            clauses = [text]
        allow_pure = kind in ("eyebrow", "title", "lead", "caption", "block_h", "sub")
        allow_external = kind.endswith("_ext")
        for clause in clauses:
            if has_exemption(clause):
                continue
            clean = strip_numbering(clause, allow_pure_numbering=allow_pure)
            # 数字
            for value, unit, raw in extract_numbers(clean):
                self.coverage["numbers_checked"] += 1
                ok, paper_unit, ev = self.paper_index.find_number(value, unit)
                if ok:
                    self.coverage["numbers_grounded"] += 1
                    if unit and paper_unit and norm_unit(unit) != norm_unit(paper_unit):
                        self.add(pno, page_type, field, "warn", "单位",
                                 "数字「%s」在原文存在，但单位不一致：你写 %s，原文是 %s" % (value, unit, paper_unit),
                                 ev[0]["snippet"] if ev else "",
                                 suggestion="核对单位换算；不确定就写「论文没有写」")
                elif allow_external or self._is_external_claim(clause):
                    if self.source_text:
                        if self._source_has_number(value, unit):
                            self.coverage["numbers_grounded"] += 1
                        else:
                            self.add(pno, page_type, field, "error", "数字",
                                     "外部指标数字「%s」在 --source-doc 里找不到" % raw,
                                     self._source_snippet(),
                                     suggestion="核对来源记录；影响因子/分区必须现查")
                    else:
                        self.add(pno, page_type, field, "hint", "外部数据",
                                 "「%s」是期刊外部指标（影响因子/分区），论文原文无法核验" % raw,
                                 suggestion="把 LetPub/期刊官网记录存成 md 后 --source-doc 传进来；或在 paper.source_note 写明来源")
                else:
                    near = self.paper_index.similar_numbers(value)
                    ev_text = near[0]["snippet"] if near else ""
                    suggestion = "论文原文没有这个数字，别编。"
                    if near:
                        suggestion += "原文相近的是「%s」（第 %s 页），确认是不是抄错。" % (
                            near[0]["value"], near[0]["page"])
                    self.add(pno, page_type, field, "error", "数字",
                             "数字「%s」在论文原文中找不到（容差 0）" % raw,
                             ev_text, suggestion)
            # 英文术语
            for tok in extract_en_tokens(clean):
                self.coverage["en_tokens_checked"] += 1
                ok, page, ev = self.paper_index.find_token(tok)
                if ok:
                    self.coverage["en_tokens_grounded"] += 1
                elif allow_external or self._is_external_claim(clause):
                    if self.source_text:
                        sl = self.source_lower
                        comp = normalize_key(tok)
                        ok2 = bool(re.search(r"(?<![a-z0-9])" + re.escape(tok.lower()) + r"(?![a-z0-9])", sl)) or (
                            len(comp) >= 4 and comp in self.source_key
                        )
                        if ok2:
                            self.coverage["en_tokens_grounded"] += 1
                        else:
                            self.add(pno, page_type, field, "error", "术语",
                                     "外部指标术语「%s」在 --source-doc 里找不到" % tok,
                                     self._source_snippet(),
                                     suggestion="核对来源记录拼写")
                    else:
                        self.add(pno, page_type, field, "hint", "外部数据",
                                 "「%s」是期刊外部指标术语，论文原文无法核验" % tok,
                                 suggestion="把 LetPub/期刊官网记录存成 md 后 --source-doc 传进来")
                else:
                    near = self.paper_index.similar_tokens(tok)
                    ev_text = near[0]["snippet"] if near else ""
                    suggestion = "论文原文没有这个英文术语/缩写，确认是否自己起的名字。"
                    if near:
                        suggestion += "原文相近的是「%s」。" % near[0]["token"]
                    self.add(pno, page_type, field, "error", "术语",
                             "英文术语「%s」在论文原文中找不到" % tok,
                             ev_text, suggestion)

    # ----- 4. 图注 / 图号 -----
    def check_figures(self, page, pno):
        t = page.get("type", "")
        files = []
        if page.get("fig"):
            files.append((page["fig"], page.get("caption")))
        for i, f in enumerate(page.get("figs") or []):
            if isinstance(f, dict):
                files.append((f.get("file"), f.get("caption")))
        if not files and t in ("left_fig", "fig_bullets", "big_fig", "two_figs"):
            self.add(pno, t, "fig", "error", "图片", "该页型需要图片（fig/figs），但没有给")
            return
        for i, (fn, cap) in enumerate(files):
            if not fn:
                self.add(pno, t, "fig", "error", "图片", "第 %d 张图没有文件名" % (i + 1))
                continue
            self.coverage["figures_checked"] += 1
            self.coverage["fields_checked"] += 1
            fig = self._find_figure(fn)
            if self.figures and fig is None:
                self.add(pno, t, "fig", "error", "图注",
                         "图片文件 %s 不在 figures.json 里" % fn,
                         suggestion="从抽图结果里选文件名（或先跑 extract_figures.py）")
                continue
            if not cap:
                self.add(pno, t, "fig", "error", "图注", "图片 %s 没有 caption" % fn)
                continue
            m = re.search(r"(?:图|Fig\.?|Figure)\s*(\d+(?:\.\d+)?)\s*(?:[（(]([a-zA-Z])[)）])?", str(cap))
            if not m:
                self.add(pno, t, "fig", "warn", "图注",
                         "图注里没有「图 N」编号：%s" % str(cap)[:30],
                         suggestion="图注编号用论文原图号，如「图 2.6（a）」")
            else:
                num, panel = m.group(1), (m.group(2) or "").lower()
                key = (num, panel)
                prev = self._fig_no_map.get(key)
                if prev and prev[1] != os.path.basename(fn):
                    self.add(pno, t, "fig", "error", "图注",
                             "图号「图 %s%s」同时指向两张不同的图：%s / %s" % (
                                 num, "(%s)" % panel if panel else "", prev[1], os.path.basename(fn)),
                             suggestion="同一图号只能对应一张图；重编或改图号")
                else:
                    self._fig_no_map[key] = (pno, os.path.basename(fn))
                if fig:
                    src_no = str(fig.get("paper_no") or "")
                    src_panel = (fig.get("panel") or "").lower()
                    if src_panel and panel and src_panel != panel:
                        self.add(pno, t, "fig", "warn", "图注",
                                 "图注写的是子图（%s），但文件 %s 是子图（%s）" % (panel, fn, src_panel),
                                 suggestion="检查 figNNx.png 是否与图注的 (a)/(b) 对应")
                    if src_no and panel and not src_panel:
                        self.add(pno, t, "fig", "hint", "图注",
                                 "图注带子图（%s），但 %s 是整张图；需要单张可加 --split-panels 重抽" % (panel, fn))

    # ----- 5. 标题空泛 -----
    def check_vague_title(self, page, pno):
        t = page.get("type", "")
        title = page.get("title")
        if not title:
            if t not in ("cover", "thanks", "paper_card", "toc", "quote", "chapter"):
                self.add(pno, t, "title", "error", "结构", "页面没有 title")
            return
        core = clean_inline(SECTION_NUM_RE.sub("", clean_inline(title)))
        core = re.sub(r"^[0-9.、．\s　]+", "", core).strip()
        if core in VAGUE_TITLES:
            self.add(pno, t, "title", "hint", "标题",
                     "标题空泛：「%s」读者看不出这页结论" % title,
                     suggestion="标题写成具体结论或问题，如「现有方法为什么不够用」（≤22 字）")
            return
        if len(core) <= 4 and re.search(r"(背景|方法|模型|实验|结论|总结|展望)$", core):
            self.add(pno, t, "title", "hint", "标题",
                     "标题偏空泛：「%s」" % title,
                     suggestion="补上这页的具体结论（≤22 字）")

    # ----- 总控 -----
    def run(self):
        pages = self.content.get("pages") or []
        numbers = self.page_map["numbers"]
        self.check_paper_card()
        for idx, page in enumerate(pages):
            pno = numbers[idx] if idx < len(numbers) else idx + 1
            self.coverage["pages_checked"] += 1
            self.check_lengths(page, pno)
            self.check_vague_title(page, pno)
            self.check_figures(page, pno)
            for field, text, kind in iter_text_fields(page):
                severity_notes = kind == "notes"
                before = len(self.issues)
                self._check_claim_text(pno, page.get("type", ""), field, text, kind)
                if severity_notes:
                    for iss in self.issues[before:]:
                        iss["severity"] = "hint"
                        iss["category"] = "讲稿(提示)"
        # 图号一致性来自 content，figures.json 只在文件缺失时报错
        return self.issues


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def _snip(s, n=120):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    return s[:n] + ("…" if len(s) > n else "")


def print_report(content, paper_path, issues, coverage, strict, page_map, fig_path):
    sev_order = {"error": 0, "warn": 1, "hint": 2}
    counts = {"error": 0, "warn": 0, "hint": 0}
    for it in issues:
        counts[it["severity"]] = counts.get(it["severity"], 0) + 1
    print("=" * 96)
    print("verify_claims：内容回查报告")
    print("  content：%s" % os.path.abspath(content.get("_path", "")) if content.get("_path") else "  content：content.json")
    print("  论文：%s" % paper_path)
    if fig_path:
        print("  抽图：%s" % fig_path)
    pm = page_map
    print("  正文页码：%s–%s（自动封面=1，论文简介=%s，目录=%s）" % (
        min(pm["numbers"]) if pm["numbers"] else "?", max(pm["numbers"]) if pm["numbers"] else "?",
        pm.get("paper_card_page") or "无", pm.get("toc_page") or "无"))
    print("-" * 96)
    groups = {}
    for it in issues:
        key = (it["page"] or 0, it["page_type"] or "")
        groups.setdefault(key, []).append(it)
    if not groups:
        print("没有发现问题。")
    for (pno, ptype) in sorted(groups):
        its = sorted(groups[(pno, ptype)], key=lambda x: (sev_order.get(x["severity"], 3), x["field"]))
        errs = sum(1 for x in its if x["severity"] == "error")
        warns = sum(1 for x in its if x["severity"] == "warn")
        hints = sum(1 for x in its if x["severity"] == "hint")
        print("[第 %s 页 · %s] 错误 %d / 警告 %d / 提示 %d" % (pno, ptype or "?", errs, warns, hints))
        for it in its:
            tag = {"error": "错误", "warn": "警告", "hint": "提示"}[it["severity"]]
            print("  · [%s] %s | %s" % (tag, it["field"], it["message"]))
            if it.get("evidence"):
                print("      证据（原文）：%s" % _snip(it["evidence"]))
            if it.get("suggestion"):
                print("      建议：%s" % it["suggestion"])
        print("-" * 96)
    total = len(issues)
    print("汇总：错误 %d / 警告 %d / 提示 %d（共 %d 条）" % (counts["error"], counts["warn"], counts["hint"], total))
    print("覆盖：页面 %d 个；字段 %d 个；数字断言 %d 条（命中 %d，%s）；英文术语 %d 个（命中 %d，%s）；图 %d 张" % (
        coverage["pages_checked"], coverage["fields_checked"],
        coverage["numbers_checked"], coverage["numbers_grounded"],
        ("%.0f%%" % (100.0 * coverage["numbers_grounded"] / coverage["numbers_checked"])) if coverage["numbers_checked"] else "—",
        coverage["en_tokens_checked"], coverage["en_tokens_grounded"],
        ("%.0f%%" % (100.0 * coverage["en_tokens_grounded"] / coverage["en_tokens_checked"])) if coverage["en_tokens_checked"] else "—",
        coverage["figures_checked"]))
    verdict = "未通过（有错误）" if counts["error"] else "通过"
    print("结论：%s%s" % (verdict, "；--strict 模式 exit 1" if (strict and counts["error"]) else ""))
    print("=" * 96)
    return counts


def build_report_json(issues, coverage, counts, content, paper_path, fig_path, page_map):
    return {
        "tool": "verify_claims",
        "generated_at": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "content": os.path.abspath(content.get("_path", "")) if content.get("_path") else "",
        "paper": os.path.abspath(paper_path) if paper_path else "",
        "figures": os.path.abspath(fig_path) if fig_path else "",
        "page_map": page_map,
        "summary": {
            "errors": counts["error"],
            "warns": counts["warn"],
            "hints": counts["hint"],
            "total": len(issues),
            "ok": counts["error"] == 0,
        },
        "coverage": coverage,
        "issues": issues,
    }


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="content.json 内容回查（数字/术语/字段/长度/图注/空标题）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  /opt/anaconda3/bin/python3 verify_claims.py --content content.json \\\n"
            "      --paper work/a/paper/paper.md --figures work/a/figures/figures.json --strict\n"
        ),
    )
    ap.add_argument("--content", required=True, help="content.json 路径")
    ap.add_argument("--paper", required=True, help="paper_extract 产出的 paper.md")
    ap.add_argument("--figures", default=None, help="extract_figures 产出的 figures.json")
    ap.add_argument("--json", default=None, help="把报告写成 JSON")
    ap.add_argument("--source-doc", action="append", default=None,
                    help="期刊外部数据来源（可多次传）：LetPub/官网记录 md/txt，用于核验 paper.metrics/ranks")
    ap.add_argument("--strict", action="store_true", help="有 error 时 exit 1")
    args = ap.parse_args(argv)

    if not os.path.exists(args.content):
        print("找不到 content.json：%s" % args.content, file=sys.stderr)
        return 2
    if not os.path.exists(args.paper):
        print("找不到 paper.md：%s\n先跑：paper_extract.py --pdf x.pdf --out work/xxx/paper" % args.paper, file=sys.stderr)
        return 2
    try:
        with open(args.content, "r", encoding="utf-8") as f:
            content = json.load(f)
    except json.JSONDecodeError as e:
        print("content.json 不是合法 JSON：%s（第 %d 行）" % (e.msg, e.lineno), file=sys.stderr)
        return 2
    content["_path"] = args.content
    figures_doc = None
    if args.figures:
        if not os.path.exists(args.figures):
            print("找不到 figures.json：%s" % args.figures, file=sys.stderr)
            return 2
        with open(args.figures, "r", encoding="utf-8") as f:
            figures_doc = json.load(f)
    source_docs = []
    for p in args.source_doc or []:
        if not os.path.exists(p):
            print("找不到 --source-doc 文件：%s" % p, file=sys.stderr)
            return 2
        with open(p, "r", encoding="utf-8", errors="replace") as f:
            source_docs.append(f.read())

    paper = PaperIndex(args.paper)
    checker = Checker(content, paper, figures_doc, source_docs)
    issues = checker.run()
    counts = {"error": 0, "warn": 0, "hint": 0}
    for it in issues:
        counts[it["severity"]] = counts.get(it["severity"], 0) + 1
    print_report(content, args.paper, issues, checker.coverage, args.strict,
                 checker.page_map, args.figures)
    if args.json:
        rep = build_report_json(issues, checker.coverage, counts, content, args.paper,
                                args.figures, checker.page_map)
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(rep, f, ensure_ascii=False, indent=1)
            f.write("\n")
        print("JSON 报告已写入：%s" % os.path.abspath(args.json))
    if args.strict and counts["error"]:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

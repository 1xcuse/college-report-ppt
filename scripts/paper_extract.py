#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""paper_extract.py — 论文 PDF / 纯文本 → 结构化素材。

输入：
    --pdf   论文 PDF（也接受 .txt / .md 文本，便于测试与纯文本材料）
    --out   输出目录

输出（写进 --out 目录）：
    paper.md         正文（带 <!-- page:N --> 页标记；已去页眉页脚；参考文献只保留标题与条数）
    paper_card.json  题名 / 作者 / 单位 / 摘要 / 关键词 / 期刊卷期 DOI / 章节结构 /
                     图表标题清单 / 公式候选 / 参考文献条数 / 解析告警
    layout.json      每页文本块位置（bbox / 字号 / 粗斜体 / 所在列 / kind），供图区检测

用法：
    python3 paper_extract.py --pdf x.pdf --out work/xxx/paper [--pages 1-2] [--max-figures 10]

依赖：PyMuPDF（import pymupdf 或 import fitz）。脚本会先探测本进程，再尝试用
      /opt/anaconda3/bin/python3 等已知解释器自动换解释器运行；都失败时给中文报错。
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import subprocess
import sys
import unicodedata
import warnings

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# 依赖探测与自动换解释器
# ---------------------------------------------------------------------------

_PY_CANDIDATES = [
    os.environ.get("PAPER_PYTHON", ""),
    "/opt/anaconda3/bin/python3",
    "/usr/local/bin/python3",
    "/opt/homebrew/bin/python3",
]

_FITZ = None


def require_pymupdf():
    """导入 PyMuPDF；本进程没有则尝试换解释器重跑；仍失败给中文报错。"""
    global _FITZ
    if _FITZ is not None:
        return _FITZ
    for name in ("pymupdf", "fitz"):
        try:
            mod = __import__(name)
            _FITZ = mod
            return mod
        except Exception:
            continue
    _reexec_with_pymupdf()
    # 正常情况下上面已经 exec 换进程；到这里说明换了也没成
    raise SystemExit(
        "缺少 PyMuPDF（pymupdf / fitz），无法解析 PDF。\n"
        "请用带 PyMuPDF 的解释器运行，例如：\n"
        "  /opt/anaconda3/bin/python3 paper_extract.py --pdf x.pdf --out work/xxx/paper\n"
        "或安装依赖：\n"
        "  /opt/anaconda3/bin/python3 -m pip install pymupdf\n"
        "（DSH 自带的 python 装了 pymupdf 但签名不可用，建议直接用 Anaconda。）"
    )


def _reexec_with_pymupdf():
    if os.environ.get("PAPER_PY_REEXEC") == "1":
        return  # 已经换过，别死循环
    tried = []
    for cand in _PY_CANDIDATES:
        if not cand or not os.path.exists(cand):
            continue
        if os.path.realpath(cand) == os.path.realpath(sys.executable):
            continue
        tried.append(cand)
        try:
            r = subprocess.run([cand, "-c", "import pymupdf"], capture_output=True, timeout=30)
            ok = r.returncode == 0
            if not ok:
                r = subprocess.run([cand, "-c", "import fitz"], capture_output=True, timeout=30)
                ok = r.returncode == 0
        except Exception:
            ok = False
        if ok:
            env = dict(os.environ, PAPER_PY_REEXEC="1")
            entry = os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else ""
            if entry and os.path.exists(entry):
                os.execve(cand, [cand, entry] + sys.argv[1:], env)
            return
    sys.stderr.write(
        "缺少 PyMuPDF（pymupdf / fitz），且没能自动切换到可用解释器。\n"
        "尝试过的解释器：%s\n"
        "请手动用 Anaconda 运行：/opt/anaconda3/bin/python3 %s\n"
        % (", ".join(tried) or "（无）", os.path.abspath(sys.argv[0]) if sys.argv else "paper_extract.py")
    )
    raise SystemExit(2)


# ---------------------------------------------------------------------------
# 文本归一化 / 通用工具
# ---------------------------------------------------------------------------

_DASHES = "\u2010\u2011\u2012\u2013\u2014\u2015\u2212\u00ad"
_DASH_RE = re.compile("[" + re.escape(_DASHES) + "]")
_FULLWIDTH = {ord(c): ord(c) - 0xFEE0 for c in "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ％．，；：（）"}


def _strip_control(s: str) -> str:
    return "".join(ch for ch in s if ch >= " " or ch in "\n\t")


def normalize_text(s: str) -> str:
    """用于「在原文里找」的归一化：全半角、破折号、连字符、空白。"""
    if not s:
        return ""
    s = _strip_control(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(_FULLWIDTH)
    s = _DASH_RE.sub("-", s)
    s = s.replace("\u00a0", " ").replace("\u2009", " ").replace("\u200b", "")
    s = re.sub(r"[ \t\r\f\v]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    return s.strip()


def normalize_key(s: str) -> str:
    """更强的归一化：小写、去所有空白与连字符，用于英文术语兜底匹配。"""
    s = normalize_text(s).lower()
    s = re.sub(r"[\s\-_/]+", "", s)
    return s


def clean_inline(s: str) -> str:
    s = normalize_text(s)
    s = re.sub(r"\s+([,.;:])", r"\1", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()


def dehyphenate_join(lines, strip_leading_space=True):
    """把一段块内的行拼成段落；处理行尾连字符断词。"""
    out = ""
    for ln in lines:
        ln = ln.strip()
        if not ln:
            continue
        if not out:
            out = ln
        elif out.endswith("-") and re.match(r"^[A-Za-z0-9]", ln):
            prev = out[-2] if len(out) >= 2 else ""
            if prev.isalpha() or prev.isdigit():
                # 词内断词：communi-\ncation → communication
                out = out[:-1] + ln
            else:
                # 形如 (LAE)- oriented：保留连字符
                out = out + ln
        else:
            out = out + " " + ln
    return clean_inline(out) if out else ""


# ---------------------------------------------------------------------------
# 粘连文本修补：部分 PDF 的断行/两端对齐会让 PyMuPDF 丢掉空格，
# 例如 "TableIsummarizesthemainnotations"。用全文词表做保守切分。
# ---------------------------------------------------------------------------

_STOP_WORDS = {
    "a", "an", "the", "of", "and", "or", "for", "with", "in", "on", "at", "by", "to", "as",
    "is", "are", "was", "were", "be", "been", "this", "that", "these", "those", "we", "our",
    "it", "its", "from", "via", "per", "than", "then", "so", "such", "both", "all", "each",
    "more", "less", "high", "low", "new", "can", "may", "not", "no", "up", "down", "out",
    "under", "over", "between", "within", "into", "two", "three", "four", "one", "zero",
    "first", "second", "third", "ii", "iii", "iv", "vi", "vii", "viii", "ix",
}


def build_word_lexicon(pages_lines):
    """从全文行里统计词表（含重复词才收，降低误切风险）。"""
    from collections import Counter

    freq = Counter()
    for lines in pages_lines:
        for ln in lines:
            for w in re.findall(r"[A-Za-z][A-Za-z']+", ln["text"]):
                if len(w) >= 2:
                    freq[w.lower()] += 1
    return {w for w, c in freq.items() if c >= 2 or len(w) >= 6}


def _wordbreak_dp(s, lexicon, min_len=2, max_words=6):
    """把粘连字符串切成词表里的词；成功返回词列表，失败返回 None。"""
    n = len(s)
    if n < 4 or n > 90:
        return None
    best = {0: []}
    for i in range(n):
        if i not in best:
            continue
        cur = best[i]
        if len(cur) >= max_words:
            continue
        for j in range(n, i + min_len - 1, -1):
            if j - i < min_len:
                continue
            w = s[i:j]
            if w.lower() in lexicon or (len(w) >= 5 and re.fullmatch(r"[A-Z][a-z]+", w) and w.lower() in _STOP_WORDS):
                if j not in best or len(best[j]) > len(cur) + 1:
                    best[j] = cur + [w]
    if n not in best:
        return None
    return best[n]


def maybe_respace_glued(text, lexicon):
    """对超长无空格英文词块做保守切分（只在能切成词表词时动手）。"""
    if not text or not lexicon:
        return text

    def _fix(m):
        tok = m.group(0)
        seg = _wordbreak_dp(tok, lexicon)
        if seg and len(seg) >= 2 and sum(1 for w in seg if len(w) >= 3) >= 2:
            return " ".join(seg)
        return tok

    return re.sub(r"[A-Za-z]{12,}", _fix, text)


def is_ascii_upper_ratio(s: str):
    letters = [c for c in s if c.isalpha()]
    if not letters:
        return 0.0
    return sum(1 for c in letters if c.isupper()) / len(letters)


# ---------------------------------------------------------------------------
# 页内行提取（保留 bbox / 字号 / 粗斜体）
# ---------------------------------------------------------------------------

def _line_from_pymupdf(line):
    spans = line.get("spans") or []
    if not spans:
        return None
    text = _strip_control("".join(s.get("text", "") for s in spans))
    text = re.sub(r"[ \t]+", " ", text).strip()
    if not text:
        return None
    size = max((s.get("size", 0) or 0) for s in spans)
    # 取字符数最多的 span 的字体
    longest = max(spans, key=lambda s: len(s.get("text", "")))
    font = longest.get("font", "") or ""
    flags = 0
    for s in spans:
        flags |= int(s.get("flags", 0) or 0)
    return {
        "text": text,
        "bbox": tuple(round(v, 2) for v in line["bbox"]),
        "size": round(float(size), 2),
        "font": font,
        "bold": bool(flags & 16) or ("bold" in font.lower()),
        "italic": bool(flags & 2) or bool(re.search(r"(italic|oblique)", font, re.I)),
    }


def _rawdict_line_texts(page):
    """用字符间空隙恢复「丢空格」的行文本（只给粘连行用）。

    部分 PDF 两端对齐时 content stream 里没有空格字符，PyMuPDF 拼行就变成
    'TableIsummarizesthemainnotations' 这种。按字符 bbox 间距可以还原。
    """
    try:
        rd = page.get_text("rawdict")
    except Exception:
        return []
    texts = []
    for b in rd.get("blocks", []):
        if b.get("type") != 0:
            continue
        for l in b.get("lines", []):
            spans = l.get("spans", [])
            chars = [c for s in spans for c in s.get("chars", [])]
            if not chars:
                texts.append("")
                continue
            size = max((s.get("size", 0) or 0) for s in spans) or 8.0
            rec = []
            prev = None
            for c in chars:
                ch = c.get("c", "")
                if prev is not None:
                    gap = c["bbox"][0] - prev["bbox"][2]
                    pch = prev.get("c", "")
                    if gap > 0.12 * size and ch and not ch.isspace() and pch and not pch.isspace():
                        rec.append(" ")
                rec.append(ch)
                prev = c
            texts.append(re.sub(r"\s+", " ", "".join(rec)).strip())
    return texts


def _looks_glued(text):
    if len(text) < 40:
        return False
    return text.count(" ") <= max(3, len(text) // 30)


def extract_page_lines(page, pymupdf_module=None):
    """返回 (lines, images, drawings_count)。lines 保持 PDF 内容流顺序。"""
    fitz = pymupdf_module or require_pymupdf()
    out = []
    d = page.get_text("dict")
    raw_texts = _rawdict_line_texts(page)
    flat_idx = 0
    for bi, b in enumerate(d.get("blocks", [])):
        if b.get("type") != 0:
            continue
        for li, line in enumerate(b.get("lines", [])):
            item = _line_from_pymupdf(line)
            if item:
                if flat_idx < len(raw_texts) and _looks_glued(item["text"]):
                    recovered = raw_texts[flat_idx]
                    if recovered and recovered.count(" ") > item["text"].count(" "):
                        item["text"] = recovered
                item["block_index"] = bi
                item["line_index"] = li
                out.append(item)
            flat_idx += 1
    images = []
    try:
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                info = fitz.Pixmap  # noqa: F841  (仅确保模块可用)
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            for r in rects:
                images.append(
                    {
                        "xref": xref,
                        "bbox": [round(v, 2) for v in (r.x0, r.y0, r.x1, r.y1)],
                        "width": int(img[2]) if len(img) > 2 else 0,
                        "height": int(img[3]) if len(img) > 3 else 0,
                    }
                )
    except Exception:
        pass
    try:
        drawings_count = len(page.get_drawings())
    except Exception:
        drawings_count = 0
    return out, images, drawings_count


# ---------------------------------------------------------------------------
# 版面：列检测、页眉页脚识别
# ---------------------------------------------------------------------------

HEADER_JUNK_PATTERNS = [
    r"authorized licensed use limited to",
    r"downloaded on .*from ieee xplore",
    r"restrictions apply",
    r"personal use is permitted",
    r"republication/redistribution requires ieee permission",
    r"all rights reserved",
    r"see https?://www\.ieee\.org/publications/rights",
    r"©\s*20\d{2}\s*ieee",
    r"digital object identifier",
    r"^\d{4,5}$",
    r"^\d{4,5}\s+ieee\s",
    r"^ieee transactions on\b",
    r"^ieee journal on\b",
    r"^ieee communications\b",
    r"^vol\.\s*\d+",
    r"^no\.\s*\d+",
]
_HEADER_JUNK_RE = re.compile("|".join(HEADER_JUNK_PATTERNS), re.I)


def body_font_size(pages_lines):
    """按字符数统计最常见的字号 = 正文字号。"""
    counter = {}
    for lines in pages_lines:
        for ln in lines:
            counter[round(ln["size"], 1)] = counter.get(round(ln["size"], 1), 0) + len(ln["text"])
    if not counter:
        return 10.0
    return max(counter.items(), key=lambda kv: kv[1])[0]


def is_junk_line(ln, page_h, page_w, body_size):
    t = ln["text"].strip()
    if not t:
        return True
    if _HEADER_JUNK_RE.search(t):
        return True
    # 页眉页脚带
    top_band = 56.0 * page_h / 792.0
    bot_band = page_h - 36.0 * page_h / 792.0
    if ln["bbox"][3] <= top_band and ln["size"] <= body_size * 1.08:
        # 页眉：短、含大写/数字/期刊信息的行
        if len(t) <= 140 and (is_ascii_upper_ratio(t) > 0.55 or re.search(r"\d", t)):
            return True
    if ln["bbox"][1] >= bot_band and ln["size"] <= body_size * 1.08:
        return True
    return False


def detect_columns(lines, page_rect, body_size):
    """粗略列检测：返回 [[x0,x1], ...]；单列返回 1 个。"""
    page_w, page_h = page_rect.width, page_rect.height
    cand = []
    for ln in lines:
        x0, y0, x1, y1 = ln["bbox"]
        if is_junk_line(ln, page_h, page_w, body_size):
            continue
        if ln["size"] < body_size - 1.2:  # 图表内小字不算正文列
            continue
        w = x1 - x0
        if w < 40:
            continue
        if w > page_w * 0.72:  # 跨栏块
            continue
        cand.append((x0, x1))
    if not cand:
        return [[0.0, page_w]]
    mid = page_w / 2.0
    left = [c for c in cand if c[1] <= mid + 18]
    right = [c for c in cand if c[0] >= mid - 18]
    if len(left) >= 8 and len(right) >= 8:
        lx0 = min(c[0] for c in left)
        lx1 = max(c[1] for c in left)
        rx0 = min(c[0] for c in right)
        rx1 = max(c[1] for c in right)
        return [[round(lx0, 1), round(lx1 + 2.0, 1)], [round(rx0, 1), round(rx1, 1)]]
    x0 = min(c[0] for c in cand)
    x1 = max(c[1] for c in cand)
    return [[round(x0, 1), round(x1, 1)]]


def column_of_x(x, columns, tol=6.0):
    if not columns:
        return 0
    best, bestd = 0, 1e9
    for i, (x0, x1) in enumerate(columns):
        d = 0 if x0 <= x <= x1 else min(abs(x - x0), abs(x - x1))
        if d < bestd:
            best, bestd = i, d
    return best


# ---------------------------------------------------------------------------
# 分类正则
# ---------------------------------------------------------------------------

CAP_FIG_RE = re.compile(r"^(?:Fig\.?|FIGURE|Figure)\s*(\d{1,3}[a-z]?)\s*[.:]\s*", re.I)
CAP_TAB_RE = re.compile(r"^(?:TABLE|Table)\s+([IVXLC]{1,5}|\d{1,3})\s*[.:]?\s*", re.I)
H1_RE = re.compile(r"^([IVX]{1,8})\.\s+([A-Z0-9][A-Z0-9 \-&,:/'’\.()]{2,110})$")
H2_RE = re.compile(r"^([A-Z])\.\s+([A-Z][A-Za-z0-9\-'’ ]{1,80})$")
SPECIAL_HEADINGS = {
    "REFERENCES": "REFERENCES",
    "REFERENCE": "REFERENCES",
    "ACKNOWLEDGMENT": "ACKNOWLEDGMENT",
    "ACKNOWLEDGEMENT": "ACKNOWLEDGMENT",
    "AUTHOR BIOGRAPHIES": "AUTHOR BIOGRAPHIES",
    "BIOGRAPHIES": "AUTHOR BIOGRAPHIES",
}
EQ_NO_RE = re.compile(r"\((?:[1-9]\d?)[a-z]?\)\s*$")
MATH_CHARS = set("=∑∫∏≤≥≈∈∉∂∇αβγδεζηθικλμνξοπρστυφχψωΓΔΘΛΞΠΣΦΨΩ√∞⋅·×→←↔‖⟨⟩")


def classify_block(text, lines, body_size):
    """返回 kind ∈ heading1/heading2/special/caption_fig/caption_tab/equation/body。"""
    t = text.strip()
    flat = clean_inline(t)
    if not flat:
        return "empty"
    up = flat.upper()
    if up in SPECIAL_HEADINGS:
        return "special"
    m = H1_RE.match(flat)
    if m and len(m.group(2)) >= 3 and is_ascii_upper_ratio(m.group(2)) > 0.75:
        return "heading1"
    if CAP_FIG_RE.match(flat):
        return "caption_fig"
    if CAP_TAB_RE.match(flat) and len(flat) < 240:
        return "caption_tab"
    if H2_RE.match(flat):
        italic = any(ln.get("italic") for ln in lines)
        no_end_punct = not re.search(r"[,;:.\-]$", flat)
        if (italic or len(flat.split()) <= 9) and no_end_punct:
            return "heading2"
    if EQ_NO_RE.search(flat) and ("=" in flat or any(c in flat for c in MATH_CHARS)):
        return "equation"
    return "body"


def strip_dropcap_line(text):
    """'I. INTRODUCTION A' -> 'I. INTRODUCTION'（首字母下沉产生的孤立字母）。"""
    parts = text.rsplit(" ", 1)
    if len(parts) == 2 and len(parts[1]) == 1 and parts[1].isalpha() and parts[1].isupper():
        return parts[0].strip()
    return text


# ---------------------------------------------------------------------------
# 页眉页脚过滤后的块清单
# ---------------------------------------------------------------------------

def build_block_records(page, body_size, pymupdf_module=None, lexicon=None):
    """把一页的行组装成块记录（过滤页眉页脚），并给出列信息。"""
    lines, images, drawings_count = extract_page_lines(page, pymupdf_module)
    page_h = page.rect.height
    page_w = page.rect.width
    kept = [ln for ln in lines if not is_junk_line(ln, page_h, page_w, body_size)]
    blocks = {}
    for ln in kept:
        blocks.setdefault(ln["block_index"], []).append(ln)
    recs = []
    for bi in sorted(blocks):
        blines = blocks[bi]
        # 去掉紧跟在标题后的孤立下沉首字母行
        cleaned = []
        for ln in blines:
            t = ln["text"].strip()
            if len(t) == 1 and t.isalpha() and t.isupper() and cleaned:
                cleaned.append(ln)
                continue
            cleaned.append(ln)
        text = dehyphenate_join([ln["text"] for ln in cleaned])
        text = maybe_respace_glued(text, lexicon)
        x0 = min(ln["bbox"][0] for ln in cleaned)
        y0 = min(ln["bbox"][1] for ln in cleaned)
        x1 = max(ln["bbox"][2] for ln in cleaned)
        y1 = max(ln["bbox"][3] for ln in cleaned)
        recs.append(
            {
                "block_index": bi,
                "lines": cleaned,
                "text": text,
                "bbox": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
                "size": round(max(ln["size"] for ln in cleaned), 2),
                "italic": any(ln["italic"] for ln in cleaned),
                "bold": any(ln["bold"] for ln in cleaned),
                "kind": classify_block(text, cleaned, body_size),
            }
        )
    return recs, images, drawings_count, lines


# ---------------------------------------------------------------------------
# 首部信息：题名 / 作者 / 单位 / 摘要 / 关键词 / 期刊 / DOI
# ---------------------------------------------------------------------------

def _find_title_blocks(first_recs, body_size):
    return [
        r
        for r in first_recs
        if r["size"] >= body_size * 1.45
        and r["bbox"][1] < 200
        and len(clean_inline(r["text"])) >= 8
    ]


def _find_title(first_recs, body_size):
    big = _find_title_blocks(first_recs, body_size)
    if not big:
        return ""
    parts = [r["text"] for r in sorted(big, key=lambda r: r["bbox"][1])]
    title = " ".join(p.strip() for p in parts)
    title = re.sub(r"\s+", " ", title).strip()
    return title


def _find_authors(first_recs, body_size, title_bottom):
    """作者可能被排版拆成多行/多块（如 TCOMM），按窗口合并。"""
    abstract_y = None
    for r in first_recs:
        if clean_inline(r["text"]).lower().startswith("abstract"):
            abstract_y = r["bbox"][1]
            break
    cands = []
    for r in first_recs:
        y0 = r["bbox"][1]
        if y0 <= title_bottom - 2:
            continue
        if abstract_y is not None and y0 >= abstract_y - 1:
            continue
        t = clean_inline(r["text"])
        if not t or "abstract" in t.lower()[:30]:
            continue
        if r["size"] < body_size * 0.95 or r["size"] > body_size * 1.6:
            continue
        if len(t) < 6 or len(t) > 500:
            continue
        if H1_RE.match(t) or H2_RE.match(t) or CAP_FIG_RE.match(t) or CAP_TAB_RE.match(t):
            continue
        if t.upper() in SPECIAL_HEADINGS:
            continue
        if re.search(r"[A-Z][a-z]+ [A-Z]", t) or re.search(r"(?:Member|Fellow),?\s*IEEE", t):
            cands.append(r)
    if not cands:
        return ""
    cands.sort(key=lambda r: (r["bbox"][1], r["bbox"][0]))
    merged = " ".join(clean_inline(r["text"]) for r in cands)
    return clean_inline(merged)


def clean_authors(raw):
    if not raw:
        return []
    t = re.sub(r"\s*,\s*(?:(?:Senior|Student|Life)\s+)?(?:Member|Fellow)\s*,?\s*IEEE", "", raw)
    t = re.sub(r"\s*,(?:\s*(?:Senior|Student|Life)\s+)?(?:Member|Fellow)\s+IEEE", "", t)
    t = re.sub(r"\bIEEE\b", "", t)
    t = re.sub(r"[\d\*†‡§¶]+", "", t)
    t = re.sub(r"\s+", " ", t).strip().strip(",")
    parts = re.split(r",\s*|\s+and\s+", t)
    names = []
    for p in parts:
        p = re.sub(r"^and\s+", "", p.strip().strip(","))
        if not p:
            continue
        if len(p.split()) >= 2 and re.search(r"[A-Za-z]", p):
            names.append(p)
    return names


def _space_out_glued_names(text, authors):
    """把 'XiaowenYe,HengyiLin,andYiWu' 这类粘连作者名按已知作者表切开。"""
    if not text:
        return text

    def _match_author(chunk):
        key = re.sub(r"[^A-Za-z]", "", chunk).lower()
        if not key:
            return None
        for a in authors:
            if re.sub(r"[^A-Za-z]", "", a).lower() == key:
                return a
        # 允许姓在前 / 空格位置不同
        for a in authors:
            if key.endswith(re.sub(r"[^A-Za-z]", "", a).lower()[-3:]):
                return a
        return None

    chunks = re.split(r"\s*,\s*|\s+and\s+|,and", text)
    fixed = []
    for c in chunks:
        c = re.sub(r"^and", "", c.strip())
        if not c:
            continue
        m = _match_author(c)
        fixed.append(m if m else c)
    return fixed


def _find_affiliations(first_recs, authors=None, lexicon=None):
    authors = authors or []
    chunks = []
    for r in first_recs:
        t = clean_inline(r["text"])
        # 注意：部分 PDF 会把 "are with the" 粘成 "arewiththe"
        if not re.search(r"\b(?:is|are)\s*with\s*(?:the\s*)?", t, re.I) and not re.search(
            r"(?:is|are)with(?:the)?", t, re.I
        ):
            continue
        if lexicon:
            t = maybe_respace_glued(t, lexicon)
        # 按句切分（句号后接空格+大写）
        sents = re.split(r"(?<=[.。])\s+(?=[A-Z])", t)
        for s in sents:
            m = re.match(r"^(.*?)\s*(?:are|is)\s*with\s*(?:the\s*)?(.*)$", s, re.S | re.I)
            if not m:
                continue
            names_raw, aff = m.group(1), m.group(2)
            aff = re.split(r"\s*\(e-?\s*mail", aff)[0]
            aff = aff.split(". Digital Object")[0]
            aff = aff.split("(Corresponding")[0]
            aff = clean_inline(aff).rstrip(".").strip()
            # 修复粘连：Elec-tronic → Electronic；Collegeof... → College of ...
            aff = re.sub(r"([A-Za-z])-\s+([A-Za-z])", r"\1-\2", aff)  # 保留原连字符
            names = clean_authors(names_raw)
            if not names and authors:
                names = _space_out_glued_names(names_raw, authors)
            if aff and names:
                chunks.append("%s（%s）" % ("、".join(names), aff))
            elif aff:
                chunks.append(aff)
    return "；".join(chunks)


def _find_abstract_keywords(first_recs):
    abstract = ""
    keywords = ""
    for r in first_recs:
        t = r["text"]
        low = t.lower()
        if not abstract and (low.startswith("abstract") or re.match(r"^abstract\s*[—\-–]", low)):
            abstract = re.sub(r"^abstract\s*[—\-–:]*\s*", "", t, flags=re.I)
            m = re.search(r"\bIndex Terms\b", abstract, re.I)
            if m:
                keywords = abstract[m.end():].lstrip("—-– :")
                abstract = abstract[: m.start()]
            abstract = clean_inline(abstract)
        elif not keywords and re.match(r"^index terms\s*[—\-–:]", low):
            keywords = re.sub(r"^index terms\s*[—\-–:]+\s*", "", t, flags=re.I)
            keywords = clean_inline(keywords)
    return abstract.strip(), keywords.strip()


def _parse_subject_meta(subject, page1_text):
    """从 PDF metadata subject（IEEE 格式 Journal;year;vol;issue;doi）解析期刊信息。"""
    info = {"journal": "", "year": "", "volume": "", "issue": "", "doi": ""}
    if not subject:
        return info
    parts = [p.strip() for p in subject.split(";")]
    if len(parts) >= 5 and parts[-1].startswith("10."):
        info["journal"] = parts[0]
        info["year"] = parts[1]
        info["volume"] = parts[2]
        info["issue"] = parts[3]
        info["doi"] = parts[-1]
    elif subject.startswith("10."):
        info["doi"] = subject.split()[0]
    return info


def _parse_first_page_header(page1_lines):
    """从首页页眉拿期刊名 / 卷期 / 月份年（用未过滤的原始行）。"""
    out = {"venue": "", "volume": "", "issue": "", "month_year": ""}
    texts = [ln["text"] for ln in page1_lines[:8]] if page1_lines else []
    # 页眉可能被拆成多行（先页码、再期刊行），拼起来再找
    joined = " ".join(texts)
    m = re.search(
        r"(IEEE[^,]*?),\s*VOL\.?\s*(\d+),\s*NO\.?\s*([A-Z0-9]+),\s*([A-Z]+)\s+(\d{4})", joined
    )
    if m:
        out["venue"] = clean_inline(m.group(1))
        out["volume"] = m.group(2)
        out["issue"] = m.group(3)
        out["month_year"] = "%s %s" % (m.group(4).title(), m.group(5))
    return out


def _extract_page_numbers_from_headers(pages_lines, body_size):
    """收集所有页顶的独立页码数字，推断页码范围（如 17472–17488）。"""
    nums = []
    for lines in pages_lines:
        for ln in lines:
            t = ln["text"].strip()
            if re.fullmatch(r"\d{4,6}", t) and ln["bbox"][3] < 70:
                try:
                    nums.append(int(t))
                except ValueError:
                    pass
    if len(nums) >= 3:
        return min(nums), max(nums)
    return None


def _extract_funding(first_recs):
    for r in first_recs:
        t = r["text"]
        if "supported in part by" in t or "supported by the" in t:
            t = re.split(r"\s*(?:Digital Object Identifier|Xiaowen|The authors|Manuscript received)", t)[0]
            return clean_inline(t)
    return ""


# ---------------------------------------------------------------------------
# 图表标题 / 公式候选 / 参考文献计数
# ---------------------------------------------------------------------------

def classify_figure_kind(caption: str) -> str:
    t = (caption or "").lower()
    framework_kw = ["framework", "architecture", "structure", "overview", "block diagram",
                    "flowchart", "pipeline", "schematic", "process of", "procedure"]
    result_kw = ["sum-rate", "sum rate", "snr", "nmse", "pareto", "trajectory", "convergence",
                 "ablation", "comparison", "versus", "vs.", "performance", "results", "achieved",
                 "error", "trade-off", "tradeoff", "throughput", "accuracy", "robustness",
                 "evaluation", "coverage", "kpi"]
    scene_kw = ["system model", "scenario", "deployment", "systems", "setup", "illustration",
                "low-altitude", "network"]
    if any(k in t for k in framework_kw):
        return "framework"
    if any(k in t for k in result_kw):
        return "result"
    if any(k in t for k in scene_kw):
        return "scene"
    return "concept"


def _collect_captions(records_by_page):
    figs, tabs = [], []
    for page_no, recs in records_by_page:
        for r in recs:
            flat = clean_inline(r["text"])
            mf = CAP_FIG_RE.match(flat)
            if mf:
                figs.append(
                    {
                        "no": "Fig. " + mf.group(1),
                        "label": "Fig. " + mf.group(1),
                        "caption": flat,
                        "page": page_no,
                        "kind": classify_figure_kind(flat),
                        "caption_bbox": r["bbox"],
                    }
                )
                continue
            mt = CAP_TAB_RE.match(flat)
            if mt:
                tabs.append(
                    {
                        "no": "TABLE " + mt.group(1).upper(),
                        "label": "TABLE " + mt.group(1).upper(),
                        "caption": flat,
                        "page": page_no,
                        "caption_bbox": r["bbox"],
                    }
                )
    # 去重（同一编号可能被拆成两个块）
    def _dedup(items):
        seen = {}
        for it in items:
            if it["no"] not in seen:
                seen[it["no"]] = it
        return list(seen.values())

    return _dedup(figs), _dedup(tabs)


def _collect_formula_candidates(records_by_page):
    cands = []
    seen = set()
    for page_no, recs in records_by_page:
        for r in recs:
            flat = clean_inline(r["text"])
            if not flat:
                continue
            m = EQ_NO_RE.search(flat)
            if not m:
                continue
            no = m.group(0).strip()
            rest = clean_inline(flat[: flat.rfind("(")]).strip(" .,:;")
            if len(rest) < 3 and not any(c in rest for c in "=<>∈"):
                continue
            key = (page_no, no)
            if key in seen:
                continue
            seen.add(key)
            cands.append({"page": page_no, "no": no, "text": flat[:400]})
    return cands


# ---------------------------------------------------------------------------
# 主解析流程
# ---------------------------------------------------------------------------

def parse_page_ranges(spec, page_count):
    """'1-2' / '1,3' / '1-2,5' → [1,2,5]（1 起）。"""
    if not spec:
        return list(range(1, page_count + 1))
    pages = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a = int(a.strip()) if a.strip() else 1
            b = int(b.strip()) if b.strip() else page_count
            pages.extend(range(max(1, a), min(page_count, b) + 1))
        else:
            p = int(part)
            if 1 <= p <= page_count:
                pages.append(p)
    return sorted(set(pages))


def extract_pdf(pdf_path: str, pages_spec=None, max_figures=None, verbose=True):
    fitz = require_pymupdf()
    doc = fitz.open(pdf_path)
    page_count = doc.page_count
    page_numbers = parse_page_ranges(pages_spec, page_count)
    if not page_numbers:
        raise SystemExit("--pages 没有选中任何页面：%r（全文共 %d 页）" % (pages_spec, page_count))

    # 第一遍：收集每页行，估计正文字号
    raw_lines = {}
    for pno in page_numbers:
        page = doc[pno - 1]
        lines, _imgs, _dc = extract_page_lines(page)
        raw_lines[pno] = lines
    body_size = body_font_size([raw_lines[p] for p in page_numbers])
    first_page = page_numbers[0]
    page1_lines = raw_lines.get(first_page, [])
    lexicon = build_word_lexicon([raw_lines[p] for p in page_numbers])

    # 第二遍：块记录 + 版面信息
    records_by_page = []
    layout_pages = []
    page_meta = {}
    for pno in page_numbers:
        page = doc[pno - 1]
        recs, images, drawings_count, lines = build_block_records(page, body_size, None, lexicon)
        columns = detect_columns(lines, page.rect, body_size)
        for r in recs:
            r["column"] = column_of_x((r["bbox"][0] + r["bbox"][2]) / 2.0, columns)
        records_by_page.append((pno, recs))
        page_meta[pno] = {"width": page.rect.width, "height": page.rect.height}
        if pno == first_page:
            first_recs = recs
        layout_pages.append(
            {
                "page": pno,
                "width": round(page.rect.width, 1),
                "height": round(page.rect.height, 1),
                "columns": columns,
                "block_count": len(recs),
                "drawings_count": drawings_count,
                "images": images,
                "blocks": [
                    {
                        "bbox": r["bbox"],
                        "text": r["text"][:600],
                        "size": r["size"],
                        "bold": r["bold"],
                        "italic": r["italic"],
                        "kind": r["kind"],
                        "column": r["column"],
                    }
                    for r in recs
                ],
            }
        )

    # ---------- 首部信息 ----------
    title = _find_title(first_recs, body_size)
    meta_title = (doc.metadata or {}).get("title", "") or ""
    meta_subject = (doc.metadata or {}).get("subject", "") or ""
    if meta_title and (
        not title
        or normalize_key(meta_title)[:40] in normalize_key(" ".join(ln["text"] for ln in page1_lines))
        or normalize_key(title)[:40] in normalize_key(meta_title)
    ):
        title = clean_inline(meta_title)
    title_blocks = _find_title_blocks(first_recs, body_size)
    title_bottom = max((r["bbox"][3] for r in title_blocks), default=0)
    authors_raw = _find_authors(first_recs, body_size, title_bottom)
    authors = clean_authors(authors_raw)
    affiliations = _find_affiliations(first_recs, authors, lexicon)
    abstract, keywords_raw = _find_abstract_keywords(first_recs)
    keywords = [clean_inline(k) for k in re.split(r"[;,，、]", keywords_raw) if clean_inline(k)]
    abstract_y = None
    for r in first_recs:
        if clean_inline(r["text"]).lower().startswith("abstract"):
            abstract_y = r["bbox"][1]
            break
    header_info = _parse_first_page_header(page1_lines)
    subject_info = _parse_subject_meta(meta_subject, " ".join(ln["text"] for ln in page1_lines))
    funding = _extract_funding(first_recs)

    page_range_nums = _extract_page_numbers_from_headers([raw_lines[p] for p in page_numbers], body_size)
    page_range = ""
    if page_range_nums:
        page_range = "%d–%d" % page_range_nums

    journal = subject_info["journal"] or header_info["venue"]
    volume = subject_info["volume"] or header_info["volume"]
    issue = subject_info["issue"] or header_info["issue"]
    year = subject_info["year"] or (header_info["month_year"].split()[-1] if header_info["month_year"] else "")
    if not year:
        m_year = re.search(r"(20\d{2})", (doc.metadata or {}).get("creationDate", "") or "")
        if m_year:
            year = m_year.group(1)
    doi = subject_info["doi"]
    if not doi:
        m = re.search(r"10\.\d{4,9}/[^\s,;]+", " ".join(ln["text"] for ln in page1_lines))
        if m:
            doi = m.group(0).rstrip(".")

    # ---------- 页内容 / 章节 / 图表 / 公式 ----------
    sections = []          # 一级
    current_section = None
    current_sub = None
    in_references = False
    ref_count = 0
    ref_seen_heading = False
    md_lines = []
    warnings = []
    equation_count = 0
    caption_fig_count = 0
    pending_dropcap = ""

    if title:
        md_lines.append("# " + title)
    if authors:
        md_lines.append("**作者**：" + "、".join(authors))
    if affiliations:
        md_lines.append("**机构**：" + affiliations)
    venue_line = " · ".join([x for x in [journal, year] if x])
    if venue_line:
        md_lines.append("**期刊**：" + venue_line)
    if doi:
        md_lines.append("**DOI**：" + doi)
    md_lines.append("")

    for pno, recs in records_by_page:
        md_lines.append("<!-- page:%d -->" % pno)
        if pno == first_page:
            pass
        emitted_on_page = False
        for r in recs:
            flat = clean_inline(r["text"])
            if not flat:
                continue
            if r["kind"] == "special":
                key = SPECIAL_HEADINGS[flat.upper()]
                if key == "REFERENCES":
                    in_references = True
                    ref_seen_heading = True
                    md_lines.append("")
                    md_lines.append("## REFERENCES（参考文献条目已省略）")
                    md_lines.append("")
                    continue
                in_references = False
                md_lines.append("")
                md_lines.append("## " + key)
                md_lines.append("")
                continue
            if in_references:
                for mm in re.finditer(r"(?:^|\s)\[(\d{1,3})\]", flat):
                    ref_count = max(ref_count, int(mm.group(1)))
                continue
            if pno == first_page:
                # 跳过题名 / 作者块（已在文首单列）
                ry0 = r["bbox"][1]
                if ry0 < title_bottom - 2 and r["size"] >= body_size * 1.45:
                    continue
                if abstract_y is not None and title_bottom - 2 <= ry0 < abstract_y - 1 and r["size"] >= body_size * 0.95:
                    continue
                if r["kind"] == "body" and r["size"] > body_size * 1.35:
                    continue  # 其余大字号块（题名残留）
            kind = r["kind"]
            dropcap_prefix = ""
            if pending_dropcap and kind in ("body", "caption_fig", "caption_tab", "equation", "heading2"):
                dropcap_prefix = pending_dropcap
                flat = dropcap_prefix + flat
                pending_dropcap = ""
            if kind == "heading1":
                h = strip_dropcap_line(flat)
                if h != flat:
                    tail = flat.rsplit(" ", 1)[-1]
                    if len(tail) == 1 and tail.isupper():
                        pending_dropcap = tail
                m = H1_RE.match(h)
                if not m:
                    md_lines.append("")
                    md_lines.append(h)
                    md_lines.append("")
                    continue
                num, sec_title = m.group(1), clean_inline(m.group(2))
                current_section = {"num": num, "title": sec_title, "page": pno, "subsections": []}
                sections.append(current_section)
                current_sub = None
                md_lines.append("")
                md_lines.append("## %s. %s" % (num, sec_title))
                md_lines.append("")
                continue
            if kind == "heading2":
                h = strip_dropcap_line(flat)
                if h != flat:
                    tail = flat.rsplit(" ", 1)[-1]
                    if len(tail) == 1 and tail.isupper():
                        pending_dropcap = tail
                if not h or len(h) > 90 or EQ_NO_RE.search(h):
                    kind = "body"
                else:
                    m = H2_RE.match(h)
                    if current_section is not None and m:
                        current_sub = {"num": m.group(1), "title": clean_inline(m.group(2)), "page": pno}
                        current_section["subsections"].append(current_sub)
                    md_lines.append("")
                    md_lines.append("### " + h)
                    md_lines.append("")
                    continue
            if kind == "caption_fig":
                caption_fig_count += 1
                md_lines.append("")
                md_lines.append(flat)
                md_lines.append("")
                continue
            if kind == "caption_tab":
                md_lines.append("")
                md_lines.append(flat)
                md_lines.append("")
                continue
            if kind == "equation":
                equation_count += 1
                md_lines.append(flat)
                continue
            # 正文块：逐行拼接（处理断词）
            body_text = dropcap_prefix + dehyphenate_join([ln["text"] for ln in r["lines"]])
            if not body_text:
                continue
            if body_text.lower().startswith("abstract"):
                body_text = re.sub(r"^abstract\s*[—\-–:]*\s*", "", body_text, flags=re.I)
                md_lines.append("")
                md_lines.append("## Abstract")
                md_lines.append(body_text)
                md_lines.append("")
                continue
            if body_text.lower().startswith("index terms"):
                body_text = re.sub(r"^index terms\s*[—\-–:]+\s*", "", body_text, flags=re.I)
                md_lines.append("")
                md_lines.append("## Index Terms")
                md_lines.append(body_text)
                md_lines.append("")
                continue
            md_lines.append(body_text)
            md_lines.append("")
            emitted_on_page = True
        if in_references:
            pass
        if not emitted_on_page and pno not in (first_page,):
            pass

    # 若正文完全没有解析出 REFERENCES 标题（例如参考条目被拆散），用正则兜底计数
    if ref_count == 0:
        full_text = "\n".join(ln["text"] for pnos, recs in records_by_page for ln in [l for r in recs for l in r["lines"]])
        nums = [int(x) for x in re.findall(r"(?:^|\s)\[(\d{1,3})\]\s+[A-Z]", full_text)]
        if nums:
            ref_count = max(nums)

    figs, tabs = _collect_captions(records_by_page)
    formulas = _collect_formula_candidates(records_by_page)
    if max_figures:
        figs = figs[: int(max_figures)]

    if not sections:
        warnings.append("没有识别出一级章节标题（可能是扫描版 PDF 或非 IEEE 版式），章节结构为空")
    if not abstract:
        warnings.append("没有识别出摘要（Abstract）；请人工确认首页版面")
    if not authors:
        warnings.append("没有识别出作者；请人工核对首页")
    if not affiliations:
        warnings.append("没有识别出单位；请人工核对首页脚注")
    if not figs:
        warnings.append("没有识别出图注（Fig. n）；抽图前请检查该 PDF 是否为扫描件")
    if not page_range:
        warnings.append("未能从页眉推断页码范围")
    if not journal:
        warnings.append("未能解析期刊名，请从 paper_card.json 手动补齐")

    # ---------- paper_card.json ----------
    year = str(year or "")
    month_year = header_info.get("month_year", "")
    if page_range:
        pages_str = page_range
    else:
        pages_str = ""
    volume_str = ""
    if volume:
        bits = ["第 %s 卷" % volume]
        if issue:
            bits.append("第 %s 期" % issue)
        if month_year:
            bits.append(month_year)
        if pages_str:
            bits.append("第 %s 页" % pages_str)
        volume_str = " · ".join(bits)

    paper_card = {
        "source": {
            "path": os.path.abspath(pdf_path),
            "file": os.path.basename(pdf_path),
            "pages_total": page_count,
            "pages_parsed": page_numbers if len(page_numbers) != page_count else "全部",
            "parser": "PyMuPDF",
            "extracted_at": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        },
        "title": title,
        "title_en": title,
        "authors": authors,
        "affiliations": affiliations,
        "abstract": abstract,
        "keywords": keywords,
        "journal": journal,
        "volume": volume_str,
        "volume_number": volume,
        "issue": issue,
        "year": year,
        "month_year": month_year,
        "pages": pages_str,
        "doi": doi,
        "funding": funding,
        "sections": sections,
        "figures": figs,
        "tables": tabs,
        "formula_candidates": formulas,
        "references_count": ref_count,
        "stats": {
            "body_font_size": body_size,
            "sections": len(sections),
            "subsections": sum(len(s.get("subsections", [])) for s in sections),
            "figure_captions": len(figs),
            "table_captions": len(tabs),
            "formula_candidates": len(formulas),
            "equation_numbered_blocks": equation_count,
        },
        "warnings": warnings,
    }

    paper_md = "\n".join(md_lines).strip() + "\n"
    layout = {
        "source": paper_card["source"],
        "pages": layout_pages,
        "note": "bbox 单位 pt；kind ∈ heading1/heading2/special/caption_fig/caption_tab/equation/body；column 为列序号（0 起）",
    }

    doc.close()
    return {
        "paper_md": paper_md,
        "paper_card": paper_card,
        "layout": layout,
        "page_numbers": page_numbers,
        "page_count": page_count,
    }


# ---------------------------------------------------------------------------
# 纯文本输入（.txt / .md）支持
# ---------------------------------------------------------------------------

def extract_text_file(path, pages_spec=None, max_figures=None):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        raw = f.read()
    warnings = []
    title = ""
    m = re.search(r"^#\s+(.+)$", raw, re.M)
    if m:
        title = clean_inline(m.group(1))
    else:
        for line in raw.splitlines():
            if line.strip():
                title = clean_inline(line.strip())
                break
    sections = []
    for mm in re.finditer(r"^##\s+([IVX]+)\.\s+(.+)$", raw, re.M):
        sections.append({"num": mm.group(1), "title": clean_inline(mm.group(2)), "page": 1, "subsections": []})
    abstract = ""
    ma = re.search(r"^##\s+Abstract\s*\n(.+?)(?=\n##\s|\Z)", raw, re.M | re.S)
    if ma:
        abstract = clean_inline(ma.group(1))
    figs = []
    for mm in re.finditer(r"^(?:Fig\.?|Figure)\s*(\d{1,3}[a-z]?)\s*[.:]\s*(.+)$", raw, re.M):
        figs.append(
            {
                "no": "Fig. " + mm.group(1),
                "label": "Fig. " + mm.group(1),
                "caption": clean_inline(mm.group(0)),
                "page": 1,
                "kind": classify_figure_kind(mm.group(0)),
                "caption_bbox": [],
            }
        )
    refs = 0
    for mm in re.finditer(r"^\s*\[(\d{1,3})\]", raw, re.M):
        refs = max(refs, int(mm.group(1)))
    paper_card = {
        "source": {
            "path": os.path.abspath(path),
            "file": os.path.basename(path),
            "pages_total": 1,
            "pages_parsed": "全部",
            "parser": "text",
            "extracted_at": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        },
        "title": title,
        "title_en": title,
        "authors": [],
        "affiliations": "",
        "abstract": abstract,
        "keywords": [],
        "journal": "",
        "volume": "",
        "volume_number": "",
        "issue": "",
        "year": "",
        "month_year": "",
        "pages": "",
        "doi": "",
        "funding": "",
        "sections": sections,
        "figures": figs if not max_figures else figs[: int(max_figures)],
        "tables": [],
        "formula_candidates": [],
        "references_count": refs,
        "stats": {
            "body_font_size": 0,
            "sections": len(sections),
            "subsections": 0,
            "figure_captions": len(figs),
            "table_captions": 0,
            "formula_candidates": 0,
            "equation_numbered_blocks": 0,
        },
        "warnings": ["文本输入：作者/单位/期刊等信息需人工补充"],
    }
    layout = {"source": paper_card["source"], "pages": [], "note": "文本输入没有版面信息"}
    return {
        "paper_md": raw,
        "paper_card": paper_card,
        "layout": layout,
        "page_numbers": [1],
        "page_count": 1,
    }


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="论文 PDF / 文本 → paper.md + paper_card.json + layout.json",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  /opt/anaconda3/bin/python3 paper_extract.py --pdf paper.pdf --out work/a/paper\n"
            "  /opt/anaconda3/bin/python3 paper_extract.py --pdf paper.pdf --out work/a/paper --pages 1-2\n"
        ),
    )
    ap.add_argument("--pdf", required=True, help="论文 PDF 路径（也接受 .txt/.md）")
    ap.add_argument("--out", required=True, help="输出目录（写 paper.md / paper_card.json / layout.json）")
    ap.add_argument("--pages", default=None, help="只解析这些页，1 起，如 1-2 或 1,3,5-6")
    ap.add_argument("--max-figures", type=int, default=None, help="最多保留多少条图注（默认全部）")
    ap.add_argument("--quiet", action="store_true", help="只打印结果摘要")
    args = ap.parse_args(argv)

    pdf_path = args.pdf
    if not os.path.exists(pdf_path):
        raise SystemExit("找不到输入文件：%s" % pdf_path)
    os.makedirs(args.out, exist_ok=True)

    ext = os.path.splitext(pdf_path)[1].lower()
    if ext in (".txt", ".md"):
        result = extract_text_file(pdf_path, args.pages, args.max_figures)
    else:
        result = extract_pdf(pdf_path, args.pages, args.max_figures)

    md_path = os.path.join(args.out, "paper.md")
    card_path = os.path.join(args.out, "paper_card.json")
    layout_path = os.path.join(args.out, "layout.json")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(result["paper_md"])
    with open(card_path, "w", encoding="utf-8") as f:
        json.dump(result["paper_card"], f, ensure_ascii=False, indent=1)
        f.write("\n")
    with open(layout_path, "w", encoding="utf-8") as f:
        json.dump(result["layout"], f, ensure_ascii=False, indent=1)
        f.write("\n")

    card = result["paper_card"]
    stats = card["stats"]
    print("=" * 72)
    print("paper_extract 完成：%s" % os.path.basename(pdf_path))
    print("  输出目录：%s" % os.path.abspath(args.out))
    print("  paper.md：%d 字符 / %d 页标记" % (len(result["paper_md"]), len(re.findall(r"<!-- page:\d+ -->", result["paper_md"]))))
    print("  题名：%s" % (card["title"][:90] or "（未识别）"))
    print("  作者：%s" % ("、".join(card["authors"]) or "（未识别）"))
    print("  期刊：%s  年：%s  卷期：%s %s  页码：%s" % (card["journal"] or "?", card["year"] or "?", card["volume_number"] or "?", card["issue"] or "?", card["pages"] or "?"))
    print("  DOI：%s" % (card["doi"] or "（未识别）"))
    print("  章节：%d 一级 / %d 二级；图注：%d；表注：%d；公式候选：%d；参考文献：%s 条" % (
        stats["sections"], stats["subsections"], stats["figure_captions"], stats["table_captions"],
        stats["formula_candidates"], card["references_count"] or "?"))
    print("  摘要：%d 字；关键词：%d 个" % (len(card["abstract"]), len(card["keywords"])))
    if card["warnings"]:
        print("  告警：")
        for w in card["warnings"]:
            print("    - " + w)
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())

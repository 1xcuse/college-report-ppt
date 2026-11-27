#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_speech.py — 从 content.json 生成详细的演讲稿（.docx + .md）。

输入：
    content.json      format_spec v2（兼容 v1 的 blocks / bullets / lines 写法）
    profile.json      可选，字段见 references/profile_schema.md
    --pptx            可选，用来核对页数、读演讲者备注（需要 python-pptx）

输出结构（按 purpose 微调）：
    封面          题目 / 学校 / 学院 / 汇报人 / 日期 / 时长 / 目的
    一、时间分配表  按页数与时长分配，中文口播按 220–260 字/分钟估算
    二、逐页讲稿    页码 · 标题 · 建议用时 · 可直接念的正文 · 过渡句 · 时间紧可跳过哪句
    三、论文背景补充 PPT 里没有的：研究脉络、关键概念、方法直觉、实验细节、相邻工作
    四、术语表      中英对照 + 一句白话解释
    五、预设问答    8–12 条，按组会 / 课程 / 答辩不同侧重
    六、口头表达提示 开场、看图表、讲公式、控场、收尾

讲稿正文优先用 content.json 里的自定义字段（可选扩展）：
    页面级： {"speech": {"text": "整段可念的讲稿", "transition": "过渡句", "skip": "可跳过的句子", "bg": "本页背景补充"}}
    全局级： {"speech": {"greeting": "...", "auto_pages": {"cover": "...", "paper_card": "...", "toc": "..."},
                        "background": [{"h": "...", "p": "..."}],
                        "terms": [{"cn": "...", "en": "...", "plain": "..."}],
                        "qa": [{"q": "...", "a": "...", "scope": "组会"}],
                        "tips": [{"h": "...", "p": "..."}]}}
没有这些字段时，脚本用每页 notes + 页面上的要点/表格/图注自动拼一份可念的草稿，
补背景靠 paper / sections / table 等已有事实，不编论文里没有的东西。

用法：
    python3 build_speech.py --content content.json --profile profile.json
    python3 build_speech.py --content content.json --out 讲稿.docx --md 讲稿.md
    python3 build_speech.py --content content.json --pptx 汇报.pptx --duration 20

依赖：python-docx（DSH 自带 Python 已装；Anaconda 没装，可用 pip install python-docx）。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence

VERSION = "1.0"

# 中文口播速度区间 220–260 字/分钟，默认取中值
DEFAULT_SPEED = 240
SPEED_RANGE = (220, 260)

PURPOSE_INFO = {
    "course_report": {
        "label": "个人课程汇报",
        "audience": "课程班",
        "duration": 9,
        "tone": "课程汇报：多解释概念和背景，老师更在意你有没有真读懂",
    },
    "group_paper": {
        "label": "组会文献汇报",
        "audience": "课题组组会",
        "duration": 20,
        "tone": "组会：重点是方法和实验，背景一页带过，多讲自己的判断",
    },
    "group_progress": {
        "label": "组会进展汇报",
        "audience": "课题组组会",
        "duration": 12,
        "tone": "进展汇报：先说结论，再讲这期做了什么、卡在哪、下一步",
    },
    "proposal": {
        "label": "开题 / 中期检查",
        "audience": "开题/中期评审",
        "duration": 18,
        "tone": "开题：讲清楚为什么值得做、怎么做、能不能做完",
    },
    "defense": {
        "label": "毕业答辩",
        "audience": "答辩委员会",
        "duration": 25,
        "tone": "答辩：创新点一句话说清，每个结论都要有数据支撑",
    },
    "project": {
        "label": "项目 / 竞赛汇报",
        "audience": "项目评审",
        "duration": 15,
        "tone": "项目汇报：痛点、方案、数据、团队，少讲技术细节",
    },
}

TYPE_LABEL = {
    "cover": "封面",
    "paper_card": "论文简介",
    "toc": "目录",
    "chapter": "章节页",
    "bullets": "要点页",
    "keynums": "关键数字",
    "text": "文字页",
    "left_fig": "图文页",
    "fig_bullets": "图文页",
    "big_fig": "大图页",
    "two_figs": "双图页",
    "table": "表格页",
    "formula": "公式页",
    "compare": "对比页",
    "timeline": "时间线",
    "pipeline": "流程图",
    "quote": "金句页",
    "summary": "总结页",
    "roadmap": "路线图",
    "qa": "问答页",
    "thanks": "致谢",
}

# 自动过渡句库：按「下一页类型」分组，每份讲稿内不重复
TRANSITION_BANK = {
    # 进入论文简介 / 目录
    "paper_card": [
        "先交代一下这篇论文的来路。",
        "在讲内容之前，先说说它是谁做的、发在哪。",
        "按惯例先报出处：期刊、作者、单位。",
    ],
    "toc": [
        "报告按三节展开，先看整体结构。",
        "先把路线摆出来，后面按这个顺序走。",
        "这页是目录，我说一下每节回答什么。",
    ],
    # 进入正文页
    "bullets": [
        "接着看下一页：{next}。",
        "这一页拆开说：{next}。",
        "把上一条结论展开一下：{next}。",
        "往下落一层：{next}。",
        "这里停下来解释一下：{next}。",
    ],
    "text": [
        "用文字交代清楚：{next}。",
        "换一段文字说明：{next}。",
        "这一页不配图，直接说结论：{next}。",
    ],
    "chapter": [
        "下面进这一节：{next}。",
        "接下来看第 {sec} 节：{next}。",
        "这一节要回答的是：{next}。",
    ],
    "fig": [
        "看图更直观：{next}。",
        "对照图说：{next}。",
        "这张图对应刚才的内容：{next}。",
        "下面这张图信息量比较大：{next}。",
    ],
    "table": [
        "下面把参数摆出来：{next}。",
        "先看实验条件：{next}。",
        "表格里信息多，我挑重点：{next}。",
        "报几个关键设置：{next}。",
    ],
    "formula": [
        "下面进入建模：{next}。",
        "把刚才的描述写成公式：{next}。",
        "这一页是公式，不念符号，只讲它在算什么：{next}。",
    ],
    "compare": [
        "换个角度对比一下：{next}。",
        "把两个做法摆在一起看：{next}。",
        "这里做一个对照：{next}。",
    ],
    "pipeline": [
        "看整条流程：{next}。",
        "把步骤串起来：{next}。",
        "按时间顺序过一遍：{next}。",
    ],
    "summary": [
        "最后收一下：{next}。",
        "讲到这里可以下结论了：{next}。",
        "进入总结：{next}。",
        "把前面几条收拢：{next}。",
    ],
    "thanks": [
        "我的汇报就讲到这里。",
        "以上就是这篇论文的主要内容。",
        "内容讲完了，谢谢大家。",
    ],
    "default": [
        "接着往下看：{next}。",
        "下一页是{next}。",
        "把刚才的思路再往前推一步：{next}。",
        "顺着这个逻辑往下：{next}。",
        "下一块内容是{next}。",
    ],
}


def _next_bank_key(next_page: Page) -> str:
    if next_page.kind in ("paper_card", "toc", "cover"):
        return next_page.kind
    if next_page.type == "thanks":
        return "thanks"
    if next_page.type == "chapter":
        return "chapter"
    if next_page.type in ("left_fig", "big_fig", "two_figs", "fig_bullets"):
        return "fig"
    if next_page.type in ("table",):
        return "table"
    if next_page.type in ("formula",):
        return "formula"
    if next_page.type in ("compare",):
        return "compare"
    if next_page.type in ("pipeline",):
        return "pipeline"
    if next_page.type in ("summary",):
        return "summary"
    if next_page.type in ("bullets", "keynums", "qa"):
        return "bullets"
    if next_page.type in ("text", "left_right"):
        return "text"
    return "default"


def make_transition(page: Page, next_page: Page | None, deck: Deck) -> str:
    """生成过渡句：自定义 > 按下一页页型从库里挑未用过的 > 轮换。"""
    if page.transition:
        return page.transition
    if next_page is None:
        return "我的汇报就到这里，谢谢各位老师同学。"
    bank_key = _next_bank_key(next_page)
    bank = TRANSITION_BANK[bank_key]
    nxt = strip_number_prefix(next_page.title) if next_page.kind == "body" else next_page.display_title
    sec = next_page.num or next_page.sec_num
    used = deck.used_transitions
    offset = (page.no + len(bank_key)) % len(bank)
    for i in range(len(bank)):
        template = bank[(offset + i) % len(bank)]
        text = template.format(next=nxt, sec=sec)
        if not text.endswith(("。", "！", "？")):
            text += "。"
        if text not in used:
            used.add(text)
            return text
    text = bank[page.no % len(bank)].format(next=nxt, sec=sec)
    return text if text.endswith(("。", "！", "？")) else text + "。"


def pick_skip_sentence(page: Page, body: str) -> str:
    """挑一句「时间紧可以跳过」的具体句子（只返回句子，不返回模板）。"""
    if page.skip:
        return page.skip
    sentences = [
        clean(s) for s in re.split(r"(?<=[。！？])", body.replace("\n", "")) if len(clean(s)) > 8
    ]
    if len(sentences) < 4:
        return ""
    for sentence in reversed(sentences[1:-1]):
        if not re.search(r"\d", sentence) and 10 <= han_count(sentence) <= 55:
            return sentence
    return sentences[-2] if len(sentences) >= 2 else ""


SKIP_PREFIXES = [
    "时间紧就省这句：",
    "来不及就略过：",
    "这句可以跳过：",
    "压缩时间时，这句可以不讲：",
]


def build_skip_lines(deck: Deck) -> None:
    """控制「时间紧可跳过」的密度：每 4 页最多 1 条，优先有自定义 skip 的页。"""
    total_minutes = sum(page.seconds for page in deck.pages) / 60.0
    over_time = total_minutes > deck.duration_min * 1.05
    limit = max(3, (len(deck.pages) + 3) // 4)
    last_no = -99
    count = 0
    for page in deck.pages:
        page.skip_line = ""
        if page.type == "thanks":
            continue
        if count >= limit or page.no - last_no < 4:
            continue
        body = page_speech(page, deck)
        sentence = pick_skip_sentence(page, body)
        if not sentence:
            continue
        if page.skip:
            # 自定义跳过说明直接沿用，保证是作者写的具体句子
            page.skip_line = sentence
        else:
            if page.kind != "body" or page.chars < 250:
                continue
            if not over_time and page.seconds < 80:
                continue
            page.skip_line = SKIP_PREFIXES[count % len(SKIP_PREFIXES)] + sentence
        last_no = page.no
        count += 1


# 静态口头表达提示（不依赖内容，任何汇报都适用）
COMMON_TIPS = [
    ("开场三句话", "第一句报题目和来路（哪篇论文、发在哪），第二句报时长和结构，第三句给一句话结论。三句说完再翻页。"),
    ("每页停半秒", "翻页后先让听众看两秒再开口，别踩着上一句的尾巴直接念下一页。"),
    ("看图表怎么说", "先说横轴纵轴和单位，再说趋势，最后给一个具体数字。不要从图例开始念。"),
    ("讲公式怎么说", "先说这个公式在算什么，再点一两个关键符号，其余符号让听众自己看。不要从头念到尾。"),
    ("讲对比怎么说", "先说共同点，再说差在哪，最后说这个差别意味着什么。"),
    ("控时", "超时先砍两类内容：背景举例、公式推导。结论和数字不能砍。"),
    ("被问住", "直接说「这块我没细看，记下来会后查」，别现编。组会汇报不丢人。"),
    ("数字口径", "同一指标全篇用同一个单位和小数位；发言前把关键数字抄在纸的第一行。"),
    ("收尾", "收尾讲三件事：做了什么、得到什么、还差什么。最后一句停住，别拖长音。"),
    ("语速", "按 220–260 字/分钟准备；每 3–5 句留一个自然停顿，念稿也要抬头看人。"),
]

PURPOSE_TIPS = {
    "defense": [("答辩创新点", "创新点用一句话说完，最多三点；每点后面立刻跟一个证据（数字或对比）。")],
    "course_report": [("课程汇报", "老师更在意你有没有读懂背景和假设；每个术语第一次出现时补一句白话解释。")],
    "group_paper": [("组会侧重", "背景控制在三分钟以内，方法和实验至少占一半时间，留五分钟给提问。")],
    "proposal": [("开题节奏", "研究内容和可行性要占一半；计划表讲到季度即可，不用念每个月。")],
    "group_progress": [("进展汇报", "先给这期结论，再讲做了什么；卡点要具体到「卡在哪个实验/哪一步」。")],
}


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------


@dataclass
class Page:
    no: int
    type: str
    kind: str  # cover / paper_card / toc / body
    title: str = ""
    subtitle: str = ""
    eyebrow: str = ""
    num: str = ""
    notes: str = ""
    speech_text: str = ""
    transition: str = ""
    skip: str = ""
    bg: str = ""
    sec_num: str = ""
    sec_title: str = ""
    seconds: int = 0
    chars: int = 0
    skip_line: str = ""   # 渲染用的「时间紧可跳过」句子（密度受控）
    speech_cache: str = ""  # 讲稿正文缓存，保证多段渲染一致
    raw: dict = field(default_factory=dict)

    @property
    def display_title(self) -> str:
        parts = []
        if self.kind == "cover":
            return "封面"
        if self.kind == "paper_card":
            return "论文简介"
        if self.kind == "toc":
            return "目录"
        if self.type == "thanks":
            return self.title or "致谢"
        if self.type == "chapter":
            num = self.num or self.sec_num
            return f"第 {num} 节　{self.title}" if num else self.title
        if self.eyebrow and self.eyebrow not in self.title:
            parts.append(self.eyebrow)
        parts.append(self.title)
        return "　".join(p for p in parts if p)


@dataclass
class Deck:
    title: str
    subtitle: str
    pages: list[Page]
    paper: dict
    sections: list[dict]
    profile: dict
    purpose: str
    duration_min: int
    speed: int
    content: dict
    speech: dict
    background: list[dict] = field(default_factory=list)
    terms: list[dict] = field(default_factory=list)
    qa: list[dict] = field(default_factory=list)
    tips: list[dict] = field(default_factory=list)
    greeting: str = "各位老师、同学好。"
    stats: dict = field(default_factory=dict)
    used_transitions: set = field(default_factory=set)
    used_phrases: set = field(default_factory=set)


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------


def clean(text: str) -> str:
    text = str(text or "").replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip(" \t\r\n")


def han_count(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def speak_len(text: str) -> int:
    """口播字数：汉字 + 英文词/数字，不含标点空白。"""
    return han_count(text) + len(re.findall(r"[A-Za-z]+|\d+(?:\.\d+)?", text))


def ensure_end(text: str, mark: str = "。") -> str:
    text = clean(text)
    if not text:
        return ""
    return text if text[-1] in "。！？；…" else text + mark


def join_sentences(parts: Sequence[str]) -> str:
    out = [ensure_end(p) for p in parts if p and clean(p)]
    return "".join(out)


def hhmmss(seconds: int) -> str:
    seconds = int(round(seconds))
    minutes, sec = divmod(max(seconds, 0), 60)
    if minutes and sec:
        return f"{minutes} 分 {sec:02d} 秒"
    if minutes:
        return f"{minutes} 分钟"
    return f"{sec} 秒"


def fmt_time(seconds: int) -> str:
    seconds = int(round(seconds))
    minutes, sec = divmod(max(seconds, 0), 60)
    if minutes and sec:
        return f"{minutes} 分 {sec} 秒"
    if minutes:
        return f"{minutes} 分钟"
    return f"{sec} 秒"


def strip_number_prefix(text: str) -> str:
    """去掉标题里的「1.1　」「2、」等编号，口语里不念编号。"""
    return re.sub(r"^\s*(?:\d+(?:\.\d+)*\s*[、.．]?\s*)", "", clean(text))


# ---------------------------------------------------------------------------
# 载入 content / profile
# ---------------------------------------------------------------------------


def load_json(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise SystemExit(f"[错误] 找不到文件：{path}")
    except json.JSONDecodeError as exc:
        raise SystemExit(f"[错误] JSON 解析失败：{path}\n        {exc}")
    if not isinstance(data, dict):
        raise SystemExit(f"[错误] {path} 顶层必须是 JSON 对象")
    return data


def build_deck(content: dict, profile: dict, args: argparse.Namespace) -> Deck:
    purpose = clean(profile.get("purpose") or "group_paper")
    if purpose not in PURPOSE_INFO:
        purpose = "group_paper"
    info = PURPOSE_INFO[purpose]

    # 时长：命令行 > profile > 目的默认值
    duration = args.duration or profile.get("duration_min") or info["duration"]
    try:
        duration_min = max(3, int(round(float(duration))))
    except (TypeError, ValueError):
        duration_min = info["duration"]

    paper = content.get("paper") if isinstance(content.get("paper"), dict) else {}
    sections = [s for s in (content.get("sections") or []) if isinstance(s, dict)]

    title = clean(content.get("title") or paper.get("title_cn") or "")
    subtitle = clean(content.get("subtitle") or paper.get("title_en") or "")
    if not title:
        title = "学术汇报"

    pages_raw = [p for p in (content.get("pages") or []) if isinstance(p, dict)]
    body_raw = [p for p in pages_raw if p.get("type") not in ("cover", "paper_card", "toc")]
    explicit_cover = next((p for p in pages_raw if p.get("type") == "cover"), None)
    explicit_paper = next((p for p in pages_raw if p.get("type") == "paper_card"), None)
    explicit_toc = next((p for p in pages_raw if p.get("type") == "toc"), None)

    pages: list[Page] = []

    def make_page(raw: dict, kind: str, ptype: str, title_: str = "") -> Page:
        raw = raw or {}
        speech = raw.get("speech")
        page = Page(
            no=0,
            type=ptype,
            kind=kind,
            title=clean(title_ or raw.get("title") or ""),
            subtitle=clean(raw.get("sub") or raw.get("subtitle") or ""),
            eyebrow=clean(raw.get("eyebrow") or ""),
            num=clean(str(raw.get("num") or "")),
            notes=clean(raw.get("notes") or ""),
            raw=raw,
        )
        if isinstance(speech, str):
            page.speech_text = clean(speech)
        elif isinstance(speech, dict):
            page.speech_text = clean(speech.get("text") or "")
            page.transition = clean(speech.get("transition") or "")
            page.skip = clean(speech.get("skip") or "")
            page.bg = clean(speech.get("bg") or "")
        return page

    # 1) 封面
    if explicit_cover is not None:
        pages.append(make_page(explicit_cover, "cover", "cover"))
    else:
        pages.append(make_page({}, "cover", "cover", title))

    # 2) 论文简介
    if paper and (paper.get("title_cn") or paper.get("title_en")):
        if explicit_paper is not None:
            pages.append(make_page(explicit_paper, "paper_card", "paper_card"))
        else:
            pc = {"notes": clean(content.get("paper_notes") or "")}
            pages.append(make_page(pc, "paper_card", "paper_card"))

    # 3) 目录
    if sections and len(body_raw) >= 6:
        if explicit_toc is not None:
            pages.append(make_page(explicit_toc, "toc", "toc"))
        else:
            pages.append(make_page({}, "toc", "toc"))

    # 4) 正文
    for raw in body_raw:
        ptype = clean(raw.get("type") or "text")
        pages.append(make_page(raw, "body", ptype))

    deck = Deck(
        title=title,
        subtitle=subtitle,
        pages=pages,
        paper=paper,
        sections=sections,
        profile=profile,
        purpose=purpose,
        duration_min=duration_min,
        speed=args.speed or DEFAULT_SPEED,
        content=content,
        speech=content.get("speech") if isinstance(content.get("speech"), dict) else {},
    )

    for idx, page in enumerate(pages, start=1):
        page.no = idx

    assign_sections(deck)
    return deck


def assign_sections(deck: Deck) -> None:
    """给正文页归到 sections 里，用于时间分配表。"""
    sections = deck.sections
    for page in deck.pages:
        if page.kind == "cover":
            page.sec_title = "开场"
        elif page.kind == "paper_card":
            page.sec_title = "论文简介"
        elif page.kind == "toc":
            page.sec_title = "目录"
        elif page.type == "thanks":
            page.sec_title = "致谢"
    if not sections:
        for page in deck.pages:
            if page.kind == "body" and not page.sec_title:
                page.sec_title = "正文"
        return

    sec_by_num = {clean(s.get("num") or str(i + 1)): clean(s.get("title") or "") for i, s in enumerate(sections)}
    nums = list(sec_by_num.keys())
    current = 0
    matched_any = False
    for page in deck.pages:
        if page.kind != "body" or page.type == "thanks":
            continue
        candidate = page.eyebrow or page.title
        m = re.match(r"^\s*(\d+)", candidate or "")
        if page.type == "chapter" and page.num and page.num in sec_by_num:
            current = nums.index(page.num)
            matched_any = True
        elif m and m.group(1) in sec_by_num:
            current = nums.index(m.group(1))
            matched_any = True
        page.sec_num = nums[current]
        page.sec_title = sec_by_num[nums[current]]

    if not matched_any:
        # 没有编号线索时，按页面顺序均分到各节
        body = [p for p in deck.pages if p.kind == "body" and p.type != "thanks"]
        if body and sections:
            per = max(1, (len(body) + len(sections) - 1) // len(sections))
            for idx, page in enumerate(body):
                si = min(idx // per, len(sections) - 1)
                page.sec_num = clean(sections[si].get("num") or str(si + 1))
                page.sec_title = clean(sections[si].get("title") or "")


# ---------------------------------------------------------------------------
# 讲稿正文生成
# ---------------------------------------------------------------------------


def _items_text(page: Page, key: str = "items", k_field: str = "k", t_field: str = "t") -> list[tuple[str, str]]:
    out = []
    for item in page.raw.get(key) or []:
        if isinstance(item, dict):
            k = clean(item.get(k_field) or item.get("h") or item.get("num") or "")
            t = clean(item.get(t_field) or item.get("p") or item.get("t") or "")
            if k or t:
                out.append((k, t))
        elif isinstance(item, str):
            out.append(("", item))
    return out


def _blocks_text(page: Page) -> list[tuple[str, str]]:
    out = []
    for block in page.raw.get("blocks") or []:
        if isinstance(block, dict):
            out.append((clean(block.get("h") or ""), clean(block.get("p") or "")))
        elif isinstance(block, str):
            out.append(("", block))
    return out


def _fig_pages_text(page: Page) -> list[str]:
    lines = [clean(x) for x in (page.raw.get("lines") or []) if clean(x)]
    for kind in ("items", "blocks"):
        for item in page.raw.get(kind) or []:
            if isinstance(item, dict):
                k = clean(item.get("k") or item.get("h") or "")
                t = clean(item.get("t") or item.get("p") or "")
                if k or t:
                    lines.append("：".join(x for x in (k, t) if x))
    return lines


def _page_notes(page: Page) -> str:
    return page.notes or clean(page.raw.get("note") or "")


def compose_note_speech(page: Page, deck: Deck) -> str:
    """没有自定义讲稿时：用 notes 打底，按页型补上下文，拼成可念的草稿。"""
    notes = _page_notes(page)
    ptype = page.type
    title = strip_number_prefix(page.display_title if page.kind != "body" else page.title)

    if page.kind == "cover":
        paper = deck.paper
        venue = clean(paper.get("venue") or paper.get("journal") or "")
        authors = clean(paper.get("authors") or "")
        rank = "、".join(clean(x) for x in (paper.get("ranks") or [])[:2] if clean(x))
        parts = [f"我汇报的题目是《{deck.title}》。"]
        if deck.subtitle and not re.search(r"[\u4e00-\u9fff]", deck.subtitle):
            parts.append(f"英文题目是 {deck.subtitle}。")
        elif deck.subtitle:
            parts.append(f"副标题是{deck.subtitle}。")
        if venue:
            parts.append(f"这是一篇发表在 {venue} 上的论文。")
        if authors:
            parts.append(f"作者是 {authors}。")
        if rank:
            parts.append(f"期刊级别是 {rank}。")
        parts.append(
            f"我计划用大约 {deck.duration_min} 分钟，讲清楚三件事：这篇论文为什么做、怎么做、做出来什么。"
        )
        return _paragraphs(parts, first=deck.greeting)

    if page.kind == "paper_card":
        paper = deck.paper
        parts = ["先交代一下这篇论文的来路。"]
        journal = clean(paper.get("venue") or paper.get("journal") or "")
        volume = clean(paper.get("volume") or "")
        if journal or volume:
            parts.append(f"期刊是 {journal}{('，' + volume) if volume else ''}。")
        ranks = [clean(x) for x in (paper.get("ranks") or []) if clean(x)]
        if ranks:
            parts.append(f"级别是 {'、'.join(ranks)}。")
        metrics = clean(paper.get("metrics") or "")
        if metrics:
            parts.append(f"{metrics}。")
        authors = clean(paper.get("authors") or "")
        affil = clean(paper.get("affiliations") or "")
        if authors:
            parts.append(f"作者是 {authors}。")
        if affil:
            parts.append(f"单位是{affil}。")
        summary = clean(paper.get("summary") or "")
        if summary:
            parts.append(f"一句话概括：{summary}。")
        return "\n\n".join(parts)

    if page.kind == "toc":
        sections = deck.sections
        parts = [f"报告分 {len(sections)} 节。"]
        for sec in sections:
            num = clean(sec.get("num") or "")
            name = clean(sec.get("title") or "")
            summary = clean(sec.get("summary") or "")
            pages = [
                p
                for p in deck.pages
                if p.sec_num == num and p.kind == "body" and p.type != "thanks"
            ]
            page_range = f"（第 {pages[0].no} 到 {pages[-1].no} 页）" if pages else ""
            parts.append(f"{('第 ' + num + ' 节 ' if num else '')}{name}：{summary or '这一节的主要页面'}{page_range}。")
        return "\n\n".join(parts)

    if ptype == "chapter":
        parts = [f"进入第 {page.num or page.sec_num} 节：{page.title}。"]
        if page.subtitle:
            parts.append(f"{page.subtitle}。")
        bullets = _items_text(page, key="bullets", k_field="h", t_field="p")
        if bullets:
            names = [b[0] or b[1] for b in bullets if b[0] or b[1]]
            parts.append(f"这一节看 {'、'.join(names)}。")
        if notes:
            parts.append(notes)
        elif deck.sections:
            sec = next(
                (x for x in deck.sections if clean(str(x.get("num") or "")) == (page.num or page.sec_num)),
                None,
            )
            summary = clean((sec or {}).get("summary") or "")
            if summary:
                parts.append(f"这一节回答的是：{ensure_end(summary)}")
        return "\n\n".join(parts)

    # 正文页：notes 打底
    parts: list[str] = []
    if notes:
        parts.append(notes)

    if ptype in ("bullets", "fig_bullets"):
        items = _items_text(page)
        missing = [(k, t) for k, t in items[:5] if (t and t not in notes) or (k and k not in notes)]
        if missing:
            lines = []
            for k, t in missing:
                if k and t:
                    lines.append(f"{k}：{ensure_end(t)}")
                elif t:
                    lines.append(ensure_end(t))
            if lines:
                if "bullets_intro" not in deck.used_phrases:
                    deck.used_phrases.add("bullets_intro")
                    parts.append("页面上列了几条，我一条一条说：" + "".join(lines))
                else:
                    parts.append("具体是：" + "".join(lines))
    elif ptype == "text":
        blocks = _blocks_text(page)
        missing = [(h, p) for h, p in blocks[:4] if p and p not in notes]
        if missing:
            parts.append("".join(ensure_end((h + "：" if h else "") + p) for h, p in missing))
    elif ptype == "keynums":
        lines = []
        for item in page.raw.get("items") or []:
            if not isinstance(item, dict):
                continue
            num = clean(item.get("num") or "")
            unit = clean(item.get("unit") or "")
            t = clean(item.get("t") or "")
            if num:
                lines.append(f"{num}{(' ' + unit) if unit else ''}，{ensure_end(t)}")
        if lines:
            parts.append("先摆几个数字：" + "；".join(lines[:4]))
    elif ptype == "table":
        rows = [r for r in (page.raw.get("rows") or []) if isinstance(r, list) and len(r) >= 2]
        if rows:
            picked = rows[1:5] if len(rows) > 1 else rows
            parts.append("表里几个数先记住：" + "；".join(f"{clean(r[0])} {clean(r[1])}" for r in picked) + "。")
        tnote = clean(page.raw.get("tnote") or "")
        if tnote:
            parts.append(f"论文没有写的是：{tnote}。")
    elif ptype == "formula":
        note = clean(page.raw.get("note") or "")
        formulas = [f for f in (page.raw.get("formulas") or []) if isinstance(f, dict)]
        nos = "、".join(clean(f.get("no") or "") for f in formulas if clean(f.get("no") or ""))
        parts.append(f"公式不用念符号，先抓住它在算什么{('，' + nos) if nos else ''}。")
        if note:
            parts.append(f"提醒一句：{note}。")
    elif ptype in ("left_fig", "big_fig", "two_figs"):
        lines = _fig_pages_text(page)
        if lines:
            parts.append("图里的信息我挑重点说：" + "".join(ensure_end(x) for x in lines[:3]))
        caption = clean(page.raw.get("caption") or "")
        if caption:
            parts.append(f"图注是：{caption}。")
    elif ptype in ("compare", "summary"):
        cols = [c for c in (page.raw.get("cols") or []) if isinstance(c, dict)]
        for col in cols:
            head = clean(col.get("h") or "")
            items = [clean(x) for x in (col.get("items") or []) if clean(x)]
            if items:
                parts.append(f"{head}：" + "".join(ensure_end(x) for x in items[:4]))
    elif ptype == "pipeline":
        steps = _items_text(page, key="steps", k_field="h", t_field="t")
        if steps:
            parts.append("流程按顺序说：" + "".join(ensure_end((k + "：" if k else "") + t) for k, t in steps[:5]))
    elif ptype in ("timeline", "roadmap"):
        items = _items_text(page, key="steps", k_field="h", t_field="t") or _items_text(
            page, key="items", k_field="when", t_field="t"
        )
        if items:
            parts.append("计划按时间段说：" + "".join(ensure_end(f"{k}，{t}") for k, t in items[:5]))
    elif ptype == "quote":
        text = clean(page.raw.get("text") or "")
        if text:
            parts.append(f"作者的原话可以这么理解：{text}")
    elif ptype == "qa":
        items = _items_text(page, key="items", k_field="q", t_field="a")
        if items:
            parts.append("这里先把可能被问到的两点摆出来：" + "".join(
                ensure_end(f"问：{k}") + ensure_end(f"答：{t}") for k, t in items[:3]
            ))
    elif ptype == "thanks":
        parts.append(clean(page.raw.get("subtitle") or "我的汇报就到这里，请老师和同学批评指正。"))

    if page.bg:
        parts.append(page.bg)

    # 兜底：什么都不够时给一句结构说明，避免空页
    if not parts:
        parts.append(f"这一页是{page.title or TYPE_LABEL.get(ptype, '正文')}。")

    return "\n\n".join(p for p in parts if p)


def _paragraphs(parts: Sequence[str], first: str = "") -> str:
    """把句子列表合并成 2–3 个自然段（开场页用）。"""
    sentences = [p for p in parts if clean(p)]
    if not sentences:
        return ""
    if len(sentences) <= 3:
        return "".join(ensure_end(s) for s in sentences)
    mid = 2
    para1 = "".join(ensure_end(s) for s in sentences[:mid])
    para2 = "".join(ensure_end(s) for s in sentences[mid:])
    head = first if first else ""
    return f"{head}{para1}\n\n{para2}".strip()


def page_speech(page: Page, deck: Deck) -> str:
    """页面讲稿正文：自定义优先；自动页用 speech.auto_pages；否则按 notes 拼。

    - 页面有 speech.text 时整段使用，不和 notes / 页面要点拼接。
    - 结果缓存在 Page.speech_cache，保证时间估算、md、docx 三处文字一致。
    """
    if page.speech_cache:
        return page.speech_cache
    if page.speech_text:
        text = page.speech_text
    elif page.kind in ("cover", "paper_card", "toc"):
        auto = deck.speech.get("auto_pages") if isinstance(deck.speech.get("auto_pages"), dict) else {}
        text = clean(auto.get(page.kind) or "")
        if not text:
            text = compose_note_speech(page, deck)
        if page.kind == "cover":
            text = f"{deck.greeting}{text}" if deck.greeting and not text.startswith(deck.greeting[:2]) else text
    else:
        text = compose_note_speech(page, deck)
    page.speech_cache = text
    return text


# ---------------------------------------------------------------------------
# 时间分配
# ---------------------------------------------------------------------------


def allocate_time(deck: Deck) -> None:
    for page in deck.pages:
        text = page_speech(page, deck)
        page.chars = max(speak_len(text), 30)
    weights = [page.chars for page in deck.pages]
    total_sec = deck.duration_min * 60
    raw = [total_sec * w / sum(weights) for w in weights]
    # 每页至少 15 秒，按 5 秒对齐
    secs = [max(15, int(round(x / 5.0) * 5)) for x in raw]
    diff = total_sec - sum(secs)
    if diff:
        order = sorted(range(len(secs)), key=lambda i: -secs[i])
        step = 5 if diff > 0 else -5
        idx = 0
        guard = 0
        while diff != 0 and guard < 10000:
            i = order[idx % len(order)]
            if secs[i] + step >= 15:
                secs[i] += step
                diff -= step
            idx += 1
            guard += 1
    for page, sec in zip(deck.pages, secs):
        page.seconds = sec


def time_groups(deck: Deck) -> list[dict]:
    order: list[str] = []
    groups: dict[str, dict] = {}
    for page in deck.pages:
        key = page.sec_title or "正文"
        if key not in groups:
            groups[key] = {"name": key, "pages": [], "seconds": 0}
            order.append(key)
        groups[key]["pages"].append(page)
        groups[key]["seconds"] += page.seconds
    out = []
    for key in order:
        group = groups[key]
        pages = group["pages"]
        span = f"{pages[0].no}" if pages[0].no == pages[-1].no else f"{pages[0].no}–{pages[-1].no}"
        out.append({"name": key, "span": span, "seconds": group["seconds"]})
    return out


# ---------------------------------------------------------------------------
# 背景补充 / 术语 / 问答 / 提示
# ---------------------------------------------------------------------------


def build_background(deck: Deck) -> list[dict]:
    custom = [x for x in (deck.speech.get("background") or []) if isinstance(x, dict)]
    if custom:
        return [
            {"h": clean(x.get("h") or "背景补充"), "p": clean(x.get("p") or "")}
            for x in custom
            if clean(x.get("p") or "")
        ]

    sections: list[dict] = []
    # 1) 研究脉络：用 sections 的 summary 与页码范围
    if deck.sections:
        paras = [f"论文分了 {len(deck.sections)} 节，按这个顺序推进："]
        for sec in deck.sections:
            num = clean(sec.get("num") or "")
            name = clean(sec.get("title") or "")
            summary = clean(sec.get("summary") or "")
            pages = [
                p for p in deck.pages if p.sec_num == num and p.kind == "body" and p.type != "thanks"
            ]
            page_range = f"第 {pages[0].no}–{pages[-1].no} 页" if pages else ""
            line = f"第{num}节「{name}」"
            if page_range:
                line += f"（{page_range}）"
            if summary:
                line += f"回答的是：{ensure_end(summary)}"
            else:
                titles = [strip_number_prefix(p.title) for p in pages[:3] if p.title]
                if titles:
                    line += f"主要讲 {'、'.join(titles)}。"
            paras.append(line)
        sections.append({"h": "研究脉络", "p": "\n\n".join(paras)})

    # 2) 关键概念：从页面要点里收 k→t，作为术语背景
    concepts: list[str] = []
    seen: set[str] = set()
    for page in deck.pages:
        if page.kind != "body":
            continue
        for k, t in _items_text(page) + _blocks_text(page):
            if not k or not t or k in seen:
                continue
            if len(k) > 14 or len(t) < 8:
                continue
            seen.add(k)
            concepts.append(f"{k}：{ensure_end(t)}")
    if concepts:
        intro = "下面这些是页面上出现过的关键词，听众不熟时可以在提问环节补讲："
        sections.append({"h": "关键概念", "p": intro + "\n\n" + "".join(concepts[:10])})

    # 3) 方法直觉：优先 quote / pipeline / formula 的说明文字
    intuition: list[str] = []
    for page in deck.pages:
        if page.type == "quote":
            text = clean(page.raw.get("text") or "")
            if text:
                intuition.append(f"作者的核心直觉：{ensure_end(text)}")
        elif page.type == "pipeline":
            steps = _items_text(page, key="steps", k_field="h", t_field="t")
            if steps:
                intuition.append(
                    "方法流程："
                    + "".join(ensure_end((k + " → ") if k else "" + t) for k, t in steps[:5]).replace(" → 。", "。")
                )
        elif page.type == "formula":
            note = clean(page.raw.get("note") or "")
            if note:
                intuition.append(f"建模时容易忽略的一点：{ensure_end(note)}")
    if intuition:
        sections.append({"h": "方法直觉", "p": "\n\n".join(intuition[:4])})

    # 4) 实验设置细节：表格页 + 数据页
    exp: list[str] = []
    for page in deck.pages:
        if page.type == "table":
            rows = [r for r in (page.raw.get("rows") or []) if isinstance(r, list) and len(r) >= 2]
            if len(rows) > 1:
                exp.append(
                    "实验设置里，论文给出的是："
                    + "；".join(f"{clean(r[0])} {clean(r[1])}" for r in rows[1:7])
                    + "。"
                )
            tnote = clean(page.raw.get("tnote") or "")
            if tnote:
                exp.append(f"论文没有交代的地方：{ensure_end(tnote)}")
    if exp:
        sections.append({"h": "实验设置细节", "p": "\n\n".join(exp)})

    # 5) 和相邻工作的关系：对比页
    related: list[str] = []
    for page in deck.pages:
        if page.type == "compare":
            cols = [c for c in (page.raw.get("cols") or []) if isinstance(c, dict)]
            names = [clean(c.get("h") or "") for c in cols if clean(c.get("h") or "")]
            if len(names) >= 2:
                related.append(f"论文把 {' 和 '.join(names[:3])} 放在一起比较，差别主要在实现方式和适用条件上。")
            for col in cols:
                items = [clean(x) for x in (col.get("items") or []) if clean(x)]
                if items:
                    related.append(f"{clean(col.get('h') or '')}：{'；'.join(items[:3])}。")
    if related:
        sections.append({"h": "和相邻工作的关系", "p": "\n\n".join(related[:3])})

    return sections


def build_terms(deck: Deck) -> list[dict]:
    custom = [x for x in (deck.speech.get("terms") or []) if isinstance(x, dict)]
    if custom:
        return [
            {
                "cn": clean(x.get("cn") or ""),
                "en": clean(x.get("en") or ""),
                "plain": clean(x.get("plain") or ""),
            }
            for x in custom
            if clean(x.get("cn") or "")
        ]

    terms: list[dict] = []
    seen: set[str] = set()
    for page in deck.pages:
        if page.kind != "body":
            continue
        pairs = _items_text(page) + _blocks_text(page)
        for k, t in pairs:
            if not k or not t or k in seen or len(k) > 14:
                continue
            seen.add(k)
            terms.append({"cn": k, "en": "", "plain": ensure_end(t)})
            if len(terms) >= 12:
                break
        if len(terms) >= 12:
            break
    if not terms:
        keywords = clean(deck.paper.get("keywords") or "")
        for kw in [x for x in re.split(r"[；;、,，]", keywords) if clean(x)][:10]:
            terms.append({"cn": clean(kw), "en": "", "plain": "论文关键词，论文简介页列出了中英文题目与关键词。"})
    return terms


def _paper_answer_bits(deck: Deck) -> dict:
    paper = deck.paper
    bits = {
        "summary": clean(paper.get("summary") or ""),
        "journal": clean(paper.get("venue") or paper.get("journal") or ""),
        "ranks": "、".join(clean(x) for x in (paper.get("ranks") or []) if clean(x)),
        "metrics": clean(paper.get("metrics") or ""),
        "keywords": clean(paper.get("keywords") or ""),
        "title": clean(paper.get("title_cn") or ""),
        "venue": clean(paper.get("venue") or paper.get("journal") or ""),
    }
    return bits


def auto_qa(deck: Deck) -> list[dict]:
    """从 content.json 已有事实生成通用问答；自定义 qa 优先。"""
    paper = deck.paper
    bits = _paper_answer_bits(deck)
    items: list[dict] = []

    def add(q: str, a: str, scope: str = "组会") -> None:
        if q and a and not any(clean(x.get("q")) == q for x in items):
            items.append({"q": q, "a": ensure_end(a), "scope": scope})

    if bits["summary"]:
        add("这篇论文一句话讲什么？", bits["summary"], "组会")

    contribs = [
        f"{clean(c.get('k') or '')}：{ensure_end(clean(c.get('t') or ''))}"
        for c in (paper.get("contributions") or [])
        if isinstance(c, dict) and clean(c.get("t") or "")
    ]
    if contribs:
        add("论文的贡献怎么概括？", "论文自己列了 " + str(len(contribs)) + " 点：" + "".join(contribs), "答辩")

    limits = [
        f"{clean(x.get('k') or '')}：{ensure_end(clean(x.get('t') or ''))}"
        for x in (paper.get("limits") or [])
        if isinstance(x, dict) and clean(x.get("t") or "")
    ]
    if limits:
        add("论文自己承认的不足有哪些？", "和刚才页面上说的一致：" + "".join(limits), "答辩")
        add("如果让你改，你先改哪里？", "我会先补论文自己列的第一条不足：" + limits[0] + "这是投入产出比最高的改法。", "答辩")

    method_notes = []
    for page in deck.pages:
        if page.kind == "body" and page.type in ("pipeline", "formula", "big_fig", "bullets"):
            note = _page_notes(page)
            if note and len(note) > 20:
                method_notes.append(f"「{strip_number_prefix(page.title)}」：{ensure_end(note)}")
            if len(method_notes) >= 2:
                break
    if method_notes:
        add("方法的核心设计是什么？", "".join(method_notes), "组会")

    table_rows: list[str] = []
    for page in deck.pages:
        if page.type == "table":
            rows = [r for r in (page.raw.get("rows") or []) if isinstance(r, list) and len(r) >= 2]
            table_rows += [f"{clean(r[0])}={clean(r[1])}" for r in rows[1:6]]
    if table_rows:
        add("实验设置里哪些数字要记住？", "最关键的是：" + "、".join(table_rows[:6]) + "。", "组会")

    if bits["journal"] or bits["ranks"]:
        add(
            "这篇论文发表在什么水平的地方？",
            f"{bits['journal']}{('，' + bits['ranks']) if bits['ranks'] else ''}"
            f"{('，' + bits['metrics']) if bits['metrics'] else ''}。",
            "课程",
        )

    if bits["venue"] or bits["title"]:
        add(
            "为什么选这篇论文来汇报？",
            f"它发表在 {bits['venue'] or '期刊'}，题目是《{bits['title'] or deck.title}》，"
            f"关键词是 {bits['keywords'] or '论文关键词'}。"
            "选它的具体理由建议换成你自己的：比如和你的课题、课程哪一部分直接相关。",
            "课程",
        )

    terms = build_terms(deck)
    if terms:
        first = terms[0]
        add(
            f"「{first['cn']}」是什么意思？",
            first["plain"] or "见论文简介页的关键词。",
            "课程",
        )

    sim_note = ""
    for page in deck.pages:
        if page.type == "table":
            tnote = clean(page.raw.get("tnote") or "")
            if "仿真" in tnote or "实测" in tnote or "仿真" in _page_notes(page):
                sim_note = tnote or _page_notes(page)
                break
    if sim_note:
        add("实验是仿真还是实测？", sim_note, "组会")

    # 对比方案
    baselines: list[str] = []
    for page in deck.pages:
        side = page.raw.get("side")
        if isinstance(side, dict):
            for item in side.get("items") or []:
                if isinstance(item, dict):
                    k = clean(item.get("k") or item.get("h") or "")
                    t = clean(item.get("t") or item.get("d") or "")
                    if k:
                        baselines.append(f"{k}（{ensure_end(t)}）" if t else k)
    if baselines:
        add("对比了哪些方案？", "四个基线：" + "、".join(baselines[:5]) + "。", "组会")

    # 总结页的结论 / 不足
    for page in deck.pages:
        if page.type != "summary":
            continue
        cols = [c for c in (page.raw.get("cols") or []) if isinstance(c, dict)]
        if len(cols) >= 1:
            conclusions = [clean(x) for x in (cols[0].get("items") or []) if clean(x)]
            if conclusions:
                add("这篇论文的结论怎么概括？", "三条结论：" + "".join(ensure_end(x) for x in conclusions[:4]), "答辩")
        if len(cols) >= 2:
            limits = [clean(x) for x in (cols[1].get("items") or []) if clean(x)]
            if limits and not any(clean(x.get("k")) == "不足" for x in []):
                add("论文的不足有哪些？", "页面右侧列了：" + "".join(ensure_end(x) for x in limits[:4]), "答辩")

    if deck.sections:
        structure = "；".join(
            f"第{clean(s.get('num') or '')}节{clean(s.get('title') or '')}"
            for s in deck.sections
        )
        add("这次汇报分几块讲？", f"分 {len(deck.sections)} 块：" + structure + "。", "课程")

    if not items:
        pages = "；".join(f"第 {p.no} 页「{strip_number_prefix(p.title)}」" for p in deck.pages if p.kind == "body" and p.title)[:120]
        add("这篇汇报是什么结构？", pages, "课程")

    return items


def build_qa(deck: Deck) -> list[dict]:
    custom = [x for x in (deck.speech.get("qa") or []) if isinstance(x, dict) and clean(x.get("q") or "")]
    pool = [
        {"q": clean(x.get("q")), "a": clean(x.get("a") or ""), "scope": clean(x.get("scope") or "通用")}
        for x in custom
    ]
    if len(pool) < 8:
        existing = {x["q"] for x in pool}
        for item in auto_qa(deck):
            if len(pool) >= 8:
                break
            if item["q"] not in existing:
                pool.append(item)
                existing.add(item["q"])

    # 按 purpose 调侧重
    priority = {
        "group_paper": ["组会", "通用", "答辩", "课程"],
        "course_report": ["课程", "通用", "组会", "答辩"],
        "group_progress": ["组会", "通用", "答辩", "课程"],
        "proposal": ["答辩", "通用", "组会", "课程"],
        "defense": ["答辩", "通用", "组会", "课程"],
        "project": ["通用", "课程", "组会", "答辩"],
    }.get(deck.purpose, ["通用", "组会", "课程", "答辩"])

    def rank(item: dict) -> int:
        scope = item.get("scope") or "通用"
        return priority.index(scope) if scope in priority else len(priority)

    unique: list[dict] = []
    seen_q: set[str] = set()
    for item in sorted(pool, key=rank):
        if item["q"] in seen_q:
            continue
        seen_q.add(item["q"])
        unique.append(item)
    return unique[:12]


def build_tips(deck: Deck) -> list[dict]:
    tips = []
    for h, p in COMMON_TIPS:
        tips.append({"h": h, "p": p})
    for h, p in PURPOSE_TIPS.get(deck.purpose, []):
        tips.append({"h": h, "p": p})
    for item in deck.speech.get("tips") or []:
        if isinstance(item, dict) and clean(item.get("p") or ""):
            tips.append({"h": clean(item.get("h") or "补充提示"), "p": clean(item.get("p") or "")})
    return tips


# ---------------------------------------------------------------------------
# docx 排版
# ---------------------------------------------------------------------------


def _require_docx():
    try:
        import docx  # noqa: F401
        from docx import Document
        from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Cm, Pt, RGBColor

        return {
            "Document": Document,
            "WD_ALIGN_PARAGRAPH": WD_ALIGN_PARAGRAPH,
            "WD_ALIGN_VERTICAL": WD_ALIGN_VERTICAL,
            "WD_TABLE_ALIGNMENT": WD_TABLE_ALIGNMENT,
            "WD_LINE_SPACING": WD_LINE_SPACING,
            "OxmlElement": OxmlElement,
            "qn": qn,
            "Cm": Cm,
            "Pt": Pt,
            "RGBColor": RGBColor,
        }
    except ImportError:
        raise SystemExit(
            "[错误] 没有装 python-docx，无法生成 .docx。\n"
            "        DSH 自带 Python 已包含：\n"
            "        DSH 自带 Python（通常在 ~/.dsh/dsh-runtimes/*/dependencies/python/bin/python3）\n"
            "        或执行：pip install python-docx\n"
            "        只想要 Markdown 时可以加 --md 并把 --out 指到 .md 文件。"
        )


def _set_run_font(run, latin: str, eastasia: str, size: float | None = None, bold: bool | None = None,
                  color: tuple[int, int, int] | None = None, dx=None) -> None:
    qn = dx["qn"]
    run.font.name = latin
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = dx["OxmlElement"]("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), eastasia)
    if size:
        run.font.size = dx["Pt"](size)
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = dx["RGBColor"](*color)


def _style_font(style, latin: str, eastasia: str, size: float | None = None, bold: bool | None = None,
                color: tuple[int, int, int] | None = None, dx=None) -> None:
    qn = dx["qn"]
    style.font.name = latin
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = dx["OxmlElement"]("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), eastasia)
    if size:
        style.font.size = dx["Pt"](size)
    if bold is not None:
        style.font.bold = bold
    if color:
        style.font.color.rgb = dx["RGBColor"](*color)


def _setup_docx(doc, deck: Deck, args: argparse.Namespace) -> dict:
    dx = _require_docx()
    section = doc.sections[0]
    section.page_width = dx["Cm"](21.0)
    section.page_height = dx["Cm"](29.7)
    section.top_margin = dx["Cm"](2.54)
    section.bottom_margin = dx["Cm"](2.54)
    section.left_margin = dx["Cm"](2.7)
    section.right_margin = dx["Cm"](2.7)

    accent = _hex_to_rgb(args.accent)
    body_font = args.body_font
    title_font = args.title_font

    normal = doc.styles["Normal"]
    _style_font(normal, "Times New Roman", body_font, args.body_size, False, (0x1A, 0x1A, 0x1A), dx)
    normal.paragraph_format.line_spacing = args.line_spacing
    normal.paragraph_format.space_before = dx["Pt"](0)
    normal.paragraph_format.space_after = dx["Pt"](4)

    for name, size, color in (("Heading 1", args.body_size + 4, accent), ("Heading 2", args.body_size + 2, accent),
                              ("Heading 3", args.body_size + 1, (0x33, 0x33, 0x33))):
        style = doc.styles[name]
        _style_font(style, "Arial", title_font, size, True, color, dx)
        style.paragraph_format.space_before = dx["Pt"](16 if name == "Heading 1" else 12)
        style.paragraph_format.space_after = dx["Pt"](8 if name == "Heading 1" else 6)
        style.paragraph_format.line_spacing = 1.3
        style.paragraph_format.keep_with_next = True

    # 页眉页脚
    header = section.header.paragraphs[0]
    header.alignment = dx["WD_ALIGN_PARAGRAPH"].RIGHT
    run = header.add_run(f"《{deck.title}》· 演讲稿")
    _set_run_font(run, "Arial", title_font, 9, False, (0x80, 0x80, 0x80), dx)

    footer = section.footer.paragraphs[0]
    footer.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
    run = footer.add_run("第 ")
    _set_run_font(run, "Times New Roman", body_font, 9, False, (0x60, 0x60, 0x60), dx)
    _add_field(footer, "PAGE", dx)
    run = footer.add_run(" 页 / 共 ")
    _set_run_font(run, "Times New Roman", body_font, 9, False, (0x60, 0x60, 0x60), dx)
    _add_field(footer, "NUMPAGES", dx)
    run = footer.add_run(" 页")
    _set_run_font(run, "Times New Roman", body_font, 9, False, (0x60, 0x60, 0x60), dx)

    doc.core_properties.title = f"《{deck.title}》演讲稿"
    doc.core_properties.author = clean(deck.profile.get("reporter") or "")
    doc.core_properties.subject = "学院汇报演讲稿"
    return {"dx": dx, "accent": accent, "body_font": body_font, "title_font": title_font}


def _add_field(paragraph, field_name: str, dx) -> None:
    run = paragraph.add_run()
    fld_begin = dx["OxmlElement"]("w:fldChar")
    fld_begin.set(dx["qn"]("w:fldCharType"), "begin")
    instr = dx["OxmlElement"]("w:instrText")
    instr.set(dx["qn"]("xml:space"), "preserve")
    instr.text = f" {field_name} "
    fld_sep = dx["OxmlElement"]("w:fldChar")
    fld_sep.set(dx["qn"]("w:fldCharType"), "separate")
    text = dx["OxmlElement"]("w:t")
    text.text = "1"
    fld_end = dx["OxmlElement"]("w:fldChar")
    fld_end.set(dx["qn"]("w:fldCharType"), "end")
    for node in (fld_begin, instr, fld_sep, text, fld_end):
        run._r.append(node)
    _set_run_font(run, "Times New Roman", "宋体", 9, False, (0x60, 0x60, 0x60), dx)


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = (value or "#00367F").strip().lstrip("#")
    if len(value) != 6:
        value = "00367F"
    try:
        return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]
    except ValueError:
        return (0x00, 0x36, 0x7F)


def _first_line_indent(paragraph, chars: int = 2, size: float = 12.0, dx=None) -> None:
    """首行缩进 2 字符：同时写 firstLineChars（Word）和 firstLine（兼容其他阅读器）。"""
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.find(dx["qn"]("w:ind"))
    if ind is None:
        ind = dx["OxmlElement"]("w:ind")
        ppr.append(ind)
    ind.set(dx["qn"]("w:firstLineChars"), str(chars * 100))
    ind.set(dx["qn"]("w:firstLine"), str(int(chars * size * 20)))


def _body_paragraph(doc, text: str, env: dict, indent: bool = True, space_after: int = 4) -> None:
    dx = env["dx"]
    paragraph = doc.add_paragraph()
    paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].JUSTIFY
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.space_after = dx["Pt"](space_after)
    run = paragraph.add_run(text)
    _set_run_font(run, "Times New Roman", env["body_font"], ARGS.body_size, False, (0x1A, 0x1A, 0x1A), dx)
    if indent:
        _first_line_indent(paragraph, 2, ARGS.body_size, dx)


def _label_paragraph(doc, label: str, text: str, env: dict, color: tuple[int, int, int] = (0x60, 0x60, 0x60)) -> None:
    dx = env["dx"]
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.line_spacing = 1.3
    paragraph.paragraph_format.space_after = dx["Pt"](4)
    run = paragraph.add_run(label)
    _set_run_font(run, "Arial", env["title_font"], ARGS.body_size - 1, True, color, dx)
    run2 = paragraph.add_run(text)
    _set_run_font(run2, "Times New Roman", env["body_font"], ARGS.body_size - 1, False, (0x33, 0x33, 0x33), dx)


def _heading(doc, text: str, level: int, env: dict) -> None:
    paragraph = doc.add_heading(level=level)
    run = paragraph.add_run(text)
    dx = env["dx"]
    size = ARGS.body_size + (4 if level == 1 else 2 if level == 2 else 1)
    _set_run_font(run, "Arial", env["title_font"], size, True, env["accent"], dx)


def _three_line_table(doc, headers: Sequence[str], rows: Sequence[Sequence[str]], env: dict,
                      widths: Sequence[float] | None = None, font_size: float | None = None,
                      align: Sequence[str] | None = None) -> None:
    dx = env["dx"]
    font_size = font_size or ARGS.body_size - 1.5
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = dx["WD_TABLE_ALIGNMENT"].CENTER
    table.autofit = False

    # 三线表：上下粗线 + 表头下线；其余边框去掉
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = dx["OxmlElement"]("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = dx["OxmlElement"](f"w:{edge}")
        element.set(dx["qn"]("w:val"), "none")
        element.set(dx["qn"]("w:sz"), "0")
        element.set(dx["qn"]("w:space"), "0")
        borders.append(element)
    for edge, size in (("top", 12), ("bottom", 12)):
        element = dx["OxmlElement"](f"w:{edge}")
        element.set(dx["qn"]("w:val"), "single")
        element.set(dx["qn"]("w:sz"), str(size))
        element.set(dx["qn"]("w:color"), "000000")
        borders.append(element)
    tbl_pr.append(borders)

    if widths:
        total = sum(widths)
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                cell.width = dx["Cm"](widths[idx] / total * 15.0)

    header_row = table.rows[0]
    header_row._tr.get_or_add_trPr().append(dx["OxmlElement"]("w:tblHeader"))
    for idx, head in enumerate(headers):
        cell = header_row.cells[idx]
        cell.vertical_alignment = dx["WD_ALIGN_VERTICAL"].CENTER
        paragraph = cell.paragraphs[0]
        paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
        run = paragraph.add_run(str(head))
        _set_run_font(run, "Arial", env["title_font"], font_size + 0.5, True, (0x1A, 0x1A, 0x1A), dx)
        _cell_bottom_border(cell, env)

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.vertical_alignment = dx["WD_ALIGN_VERTICAL"].CENTER
            paragraph = cell.paragraphs[0]
            mode = (align[c_idx] if align and c_idx < len(align) else "left")
            paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER if mode == "center" else dx["WD_ALIGN_PARAGRAPH"].LEFT
            run = paragraph.add_run(str(value))
            _set_run_font(run, "Times New Roman", env["body_font"], font_size, False, (0x1A, 0x1A, 0x1A), dx)


def _cell_bottom_border(cell, env: dict) -> None:
    dx = env["dx"]
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = dx["OxmlElement"]("w:tcBorders")
    bottom = dx["OxmlElement"]("w:bottom")
    bottom.set(dx["qn"]("w:val"), "single")
    bottom.set(dx["qn"]("w:sz"), "6")
    bottom.set(dx["qn"]("w:color"), "000000")
    borders.append(bottom)
    tc_pr.append(borders)


# ---------------------------------------------------------------------------
# 各章节写入 docx
# ---------------------------------------------------------------------------


def write_cover(doc, deck: Deck, env: dict) -> None:
    dx = env["dx"]
    for _ in range(3):
        doc.add_paragraph()
    paragraph = doc.add_paragraph()
    paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
    run = paragraph.add_run("演 讲 稿")
    _set_run_font(run, "Arial", env["title_font"], 18, False, env["accent"], dx)

    paragraph = doc.add_paragraph()
    paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
    paragraph.paragraph_format.space_before = dx["Pt"](18)
    run = paragraph.add_run(f"《{deck.title}》")
    _set_run_font(run, "Arial", env["title_font"], 24, True, (0x1A, 0x1A, 0x1A), dx)

    if deck.subtitle:
        paragraph = doc.add_paragraph()
        paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
        run = paragraph.add_run(deck.subtitle)
        _set_run_font(run, "Arial", env["title_font"], 13, False, (0x55, 0x55, 0x55), dx)

    for _ in range(2):
        doc.add_paragraph()

    prof = deck.profile
    purpose_label = PURPOSE_INFO.get(deck.purpose, {}).get("label", deck.purpose)
    info = [
        ("学　　校", clean(prof.get("school") or "")),
        ("学　　院", clean(prof.get("college") or "")),
        ("汇 报 人", clean(prof.get("reporter") or "")),
        ("身　　份", clean(prof.get("role") or "")),
        ("汇报目的", purpose_label),
        ("汇报对象", clean(prof.get("audience") or PURPOSE_INFO.get(deck.purpose, {}).get("audience", ""))),
        ("汇报日期", clean(prof.get("date") or "")),
        ("预计时长", f"{deck.duration_min} 分钟"),
        ("讲稿字数", f"{deck.stats.get('speech_chars', 0):,} 字"),
    ]
    for label, value in info:
        if not value:
            continue
        paragraph = doc.add_paragraph()
        paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
        paragraph.paragraph_format.space_after = dx["Pt"](3)
        run = paragraph.add_run(f"{label}：")
        _set_run_font(run, "Arial", env["title_font"], 11, True, (0x44, 0x44, 0x44), dx)
        run2 = paragraph.add_run(value)
        _set_run_font(run2, "Times New Roman", env["body_font"], 11, False, (0x44, 0x44, 0x44), dx)

    doc.add_paragraph()
    paragraph = doc.add_paragraph()
    paragraph.alignment = dx["WD_ALIGN_PARAGRAPH"].CENTER
    note = "本讲稿按页给出口播正文、过渡句和可跳过句；页码对应 PPT 页序，可直接照念。"
    run = paragraph.add_run(note)
    _set_run_font(run, "Times New Roman", env["body_font"], 9.5, False, (0x80, 0x80, 0x80), dx)


def write_time_table(doc, deck: Deck, env: dict) -> None:
    _heading(doc, "一、时间分配表", 1, env)
    _body_paragraph(
        doc,
        f"共 {len(deck.pages)} 页，设定时长 {deck.duration_min} 分钟；口播按 {deck.speed} 字/分钟估算，"
        f"全篇讲稿正文约 {deck.stats.get('speech_chars', 0):,} 字，"
        f"读出来约 {deck.stats.get('speech_minutes', 0):.1f} 分钟。"
        "各页用时按该页正文字数比例分配，不是平均分。",
        env,
    )
    rows = [[g["name"], g["span"], fmt_time(g["seconds"])] for g in time_groups(deck)]
    rows.append(["合计", f"1–{len(deck.pages)}", fmt_time(sum(p.seconds for p in deck.pages))])
    _three_line_table(doc, ["部分", "页码", "预计用时"], rows, env, widths=[5, 3, 4], align=["left", "center", "center"])

    over = deck.stats.get("speech_minutes", 0)
    if over and abs(over - deck.duration_min) > max(1.5, deck.duration_min * 0.15):
        _body_paragraph(
            doc,
            f"⚠ 重要提醒：按 {deck.speed} 字/分钟，正文需要约 {over:.1f} 分钟，和设定时长差得比较多。"
            "要么删掉每页标了「可跳过」的句子，要么把讲稿补到目标字数。",
            env,
            indent=True,
        )


def write_page_scripts(doc, deck: Deck, env: dict) -> None:
    _heading(doc, "二、逐页讲稿", 1, env)
    _body_paragraph(
        doc,
        "每页先给建议用时和标题，下面是可直接念的正文。过渡句用来接下一页。"
        "标了〔时间紧可跳过〕的句子可以略过；没标的页面本身不长，不用跳。"
        "不管怎么压时间，结论和数字都不要跳。",
        env,
    )
    total_pages = len(deck.pages)
    for idx, page in enumerate(deck.pages):
        next_page = deck.pages[idx + 1] if idx + 1 < total_pages else None
        _heading(doc, f"第 {page.no} 页　{page.display_title}（{fmt_time(page.seconds)}）", 2, env)
        body = page_speech(page, deck)
        for para in [p for p in body.split("\n") if p.strip()]:
            _body_paragraph(doc, para.strip(), env)
        if next_page is not None:
            _label_paragraph(doc, "〔过渡〕", page.transition, env, env["accent"])
        if page.skip_line:
            _label_paragraph(doc, "〔时间紧可跳过〕", page.skip_line, env)


def write_background(doc, deck: Deck, env: dict) -> None:
    _heading(doc, "三、论文背景补充", 1, env)
    _body_paragraph(
        doc,
        "这一节是 PPT 上没写的背景，用来回答「为什么这么做」「这个概念是什么」「实验为什么这么设」。"
        "讲的时候不用背，被问到或者时间富余时展开。",
        env,
    )
    for section in deck.background:
        _heading(doc, section["h"], 2, env)
        for para in [p for p in section["p"].split("\n") if p.strip()]:
            _body_paragraph(doc, para.strip(), env)


def write_terms(doc, deck: Deck, env: dict) -> None:
    _heading(doc, "四、术语表", 1, env)
    if not deck.terms:
        _body_paragraph(doc, "这份 content.json 没有提供术语数据；需要时可在 speech.terms 里补。", env)
        return
    has_en = any(t.get("en") for t in deck.terms)
    if has_en:
        rows = [[t["cn"], t.get("en") or "—", t.get("plain") or ""] for t in deck.terms]
        _three_line_table(doc, ["术语", "英文", "一句白话解释"], rows, env, widths=[3, 5, 8], align=["left", "left", "left"])
    else:
        rows = [[t["cn"], t.get("plain") or ""] for t in deck.terms]
        _three_line_table(doc, ["术语", "一句白话解释"], rows, env, widths=[3, 10], align=["left", "left"])


def write_qa(doc, deck: Deck, env: dict) -> None:
    _heading(doc, "五、预设问答", 1, env)
    _body_paragraph(doc, f"按组会 / 课程 / 答辩三种场合标了侧重，共 {len(deck.qa)} 条。答不上来的，用最后那句话兜底。", env)
    for idx, item in enumerate(deck.qa, start=1):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = env["dx"]["Pt"](6)
        paragraph.paragraph_format.space_after = env["dx"]["Pt"](2)
        scope = clean(item.get("scope") or "通用")
        run = paragraph.add_run(f"Q{idx}〔{scope}〕")
        _set_run_font(run, "Arial", env["title_font"], ARGS.body_size, True, env["accent"], env["dx"])
        run2 = paragraph.add_run(clean(item.get("q") or ""))
        _set_run_font(run2, "Arial", env["title_font"], ARGS.body_size, True, (0x1A, 0x1A, 0x1A), env["dx"])
        answer = doc.add_paragraph()
        answer.paragraph_format.line_spacing = 1.5
        answer.paragraph_format.space_after = env["dx"]["Pt"](4)
        _first_line_indent(answer, 2, ARGS.body_size, env["dx"])
        run3 = answer.add_run("A：" + clean(item.get("a") or ""))
        _set_run_font(run3, "Times New Roman", env["body_font"], ARGS.body_size, False, (0x1A, 0x1A, 0x1A), env["dx"])
    _body_paragraph(
        doc,
        "被问到不会的，就说「这块我没细看，记下来会后查」，不要现编。组会汇报不丢人，编错了才丢人。",
        env,
    )


def write_tips(doc, deck: Deck, env: dict) -> None:
    _heading(doc, "六、口头表达提示", 1, env)
    for tip in deck.tips:
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = env["dx"]["Pt"](5)
        paragraph.paragraph_format.space_after = env["dx"]["Pt"](2)
        run = paragraph.add_run(f"· {tip['h']}：")
        _set_run_font(run, "Arial", env["title_font"], ARGS.body_size, True, (0x1A, 0x1A, 0x1A), env["dx"])
        run2 = paragraph.add_run(tip["p"])
        _set_run_font(run2, "Times New Roman", env["body_font"], ARGS.body_size, False, (0x1A, 0x1A, 0x1A), env["dx"])


def build_docx(deck: Deck, out_path: Path, args: argparse.Namespace) -> None:
    dx = _require_docx()
    doc = dx["Document"]()
    env = _setup_docx(doc, deck, args)

    write_cover(doc, deck, env)
    doc.add_page_break()
    write_time_table(doc, deck, env)
    doc.add_page_break()
    write_page_scripts(doc, deck, env)
    doc.add_page_break()
    write_background(doc, deck, env)
    doc.add_page_break()
    write_terms(doc, deck, env)
    doc.add_page_break()
    write_qa(doc, deck, env)
    doc.add_page_break()
    write_tips(doc, deck, env)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


# ---------------------------------------------------------------------------
# Markdown 版本
# ---------------------------------------------------------------------------


def build_markdown(deck: Deck) -> str:
    lines: list[str] = []
    purpose_label = PURPOSE_INFO.get(deck.purpose, {}).get("label", deck.purpose)
    prof = deck.profile
    lines.append(f"# 演讲稿：《{deck.title}》")
    if deck.subtitle:
        lines.append(f"> {deck.subtitle}")
    lines.append("")
    lines.append(f"- 学校：{clean(prof.get('school') or '')}")
    lines.append(f"- 学院：{clean(prof.get('college') or '')}")
    lines.append(f"- 汇报人：{clean(prof.get('reporter') or '')}（{clean(prof.get('role') or '')}）")
    lines.append(f"- 汇报目的：{purpose_label}")
    lines.append(f"- 汇报对象：{clean(prof.get('audience') or PURPOSE_INFO.get(deck.purpose, {}).get('audience', ''))}")
    lines.append(f"- 日期：{clean(prof.get('date') or '')}　预计时长：{deck.duration_min} 分钟")
    lines.append(
        f"- 讲稿正文：{deck.stats.get('speech_chars', 0):,} 字，按 {deck.speed} 字/分钟约 "
        f"{deck.stats.get('speech_minutes', 0):.1f} 分钟"
    )
    lines.append("")

    lines.append("## 一、时间分配表")
    lines.append("")
    lines.append("| 部分 | 页码 | 预计用时 |")
    lines.append("|---|---|---|")
    for group in time_groups(deck):
        lines.append(f"| {group['name']} | {group['span']} | {fmt_time(group['seconds'])} |")
    lines.append(f"| 合计 | 1–{len(deck.pages)} | {fmt_time(sum(p.seconds for p in deck.pages))} |")
    lines.append("")

    lines.append("## 二、逐页讲稿")
    lines.append("")
    for idx, page in enumerate(deck.pages):
        next_page = deck.pages[idx + 1] if idx + 1 < len(deck.pages) else None
        lines.append(f"### 第 {page.no} 页　{page.display_title}（{fmt_time(page.seconds)}）")
        lines.append("")
        body = page_speech(page, deck)
        lines.append(body)
        lines.append("")
        if next_page is not None:
            lines.append(f"〔过渡〕{page.transition}")
        if page.skip_line:
            lines.append("")
            lines.append(f"〔时间紧可跳过〕{page.skip_line}")
        lines.append("")

    lines.append("## 三、论文背景补充")
    lines.append("")
    for section in deck.background:
        lines.append(f"### {section['h']}")
        lines.append("")
        lines.append(section["p"])
        lines.append("")

    lines.append("## 四、术语表")
    lines.append("")
    has_en = any(t.get("en") for t in deck.terms)
    if has_en:
        lines.append("| 术语 | 英文 | 一句白话解释 |")
        lines.append("|---|---|---|")
        for t in deck.terms:
            lines.append(f"| {t['cn']} | {t.get('en') or '—'} | {t.get('plain') or ''} |")
    else:
        lines.append("| 术语 | 一句白话解释 |")
        lines.append("|---|---|")
        for t in deck.terms:
            lines.append(f"| {t['cn']} | {t.get('plain') or ''} |")
    lines.append("")

    lines.append("## 五、预设问答")
    lines.append("")
    for idx, item in enumerate(deck.qa, start=1):
        lines.append(f"**Q{idx}〔{clean(item.get('scope') or '通用')}〕{clean(item.get('q') or '')}**")
        lines.append("")
        lines.append(f"A：{clean(item.get('a') or '')}")
        lines.append("")
    lines.append("被问到不会的，就说「这块我没细看，记下来会后查」，不要现编。")
    lines.append("")

    lines.append("## 六、口头表达提示")
    lines.append("")
    for tip in deck.tips:
        lines.append(f"- **{tip['h']}**：{tip['p']}")
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------


def read_pptx_info(path: Path) -> dict | None:
    try:
        from pptx import Presentation
    except ImportError:
        return None
    try:
        prs = Presentation(str(path))
    except Exception as exc:  # noqa: BLE001 - 只用于提示，不阻断
        print(f"[提示] 读不了 pptx：{path}（{exc}）", file=sys.stderr)
        return None
    slides = 0
    notes_chars = 0
    for slide in prs.slides:
        slides += 1
        if slide.has_notes_slide:
            notes_chars += han_count(slide.notes_slide.notes_text_frame.text)
    return {"slides": slides, "notes_chars": notes_chars}


def main(argv: Sequence[str] | None = None) -> int:
    global ARGS
    parser = argparse.ArgumentParser(
        description="从 content.json 生成演讲稿 .docx / .md",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--content", required=True, help="content.json 路径（format_spec v2，兼容 v1）")
    parser.add_argument("--profile", help="profile.json 路径（可选）")
    parser.add_argument("--out", default="讲稿.docx", help="输出 docx 路径（默认 讲稿.docx）")
    parser.add_argument("--md", nargs="?", const="", default=None,
                        help="同时输出 Markdown（不写路径时与 docx 同名 .md；传空字符串则不出）")
    parser.add_argument("--pptx", help="可选：对应 PPT，用来核对页数与备注字数")
    parser.add_argument("--duration", type=float, help="覆盖时长（分钟）")
    parser.add_argument("--speed", type=int, default=DEFAULT_SPEED, help="口播速度，字/分钟（默认 240）")
    parser.add_argument("--title-font", default="黑体", help="标题中文字体（默认 黑体）")
    parser.add_argument("--body-font", default="宋体", help="正文中文字体（默认 宋体）")
    parser.add_argument("--body-size", type=float, default=12, help="正文字号 pt（默认 12，即小四）")
    parser.add_argument("--line-spacing", type=float, default=1.5, help="行距（默认 1.5）")
    parser.add_argument("--accent", default="#00367F", help="标题颜色（默认 #00367F 深蓝）")
    args = parser.parse_args(argv)
    ARGS = args

    if not SPEED_RANGE[0] <= args.speed <= SPEED_RANGE[1]:
        print(
            f"[警告] 口播速度 {args.speed} 超出中文常规 220–260 字/分钟，时间估计会不准。",
            file=sys.stderr,
        )

    content_path = Path(args.content).expanduser()
    content = load_json(content_path)
    profile: dict = {}
    if args.profile:
        profile_path = Path(args.profile).expanduser()
        profile = load_json(profile_path)
    if isinstance(content.get("profile"), dict):
        merged = dict(content["profile"])
        merged.update(profile)
        profile = merged

    deck = build_deck(content, profile, args)

    # 生成各段内容
    deck.background = build_background(deck)
    deck.terms = build_terms(deck)
    deck.qa = build_qa(deck)
    deck.tips = build_tips(deck)
    deck.greeting = clean(deck.speech.get("greeting") or "") or "各位老师、同学好。"
    deck.greeting = clean(deck.greeting)

    allocate_time(deck)
    # 过渡句按页型轮换、全篇不重复；先算好，md 与 docx 用同一句
    deck.used_transitions = set()
    for idx, page in enumerate(deck.pages):
        next_page = deck.pages[idx + 1] if idx + 1 < len(deck.pages) else None
        page.transition = make_transition(page, next_page, deck)
    # 「时间紧可跳过」密度控制：每 4 页最多 1 条
    build_skip_lines(deck)
    speech_texts = [page_speech(page, deck) for page in deck.pages]
    speech_chars = sum(speak_len(t) for t in speech_texts)
    speech_minutes = speech_chars / deck.speed
    deck.stats = {
        "speech_chars": speech_chars,
        "speech_minutes": speech_minutes,
        "pages": len(deck.pages),
        "background": len(deck.background),
        "terms": len(deck.terms),
        "qa": len(deck.qa),
    }

    out_path = Path(args.out).expanduser()
    if not out_path.is_absolute():
        out_path = Path.cwd() / out_path

    # 输出 md
    md_path: Path | None = None
    if args.md is None:
        md_path = out_path.with_suffix(".md")
    elif args.md:
        md_path = Path(args.md).expanduser()
        if not md_path.is_absolute():
            md_path = Path.cwd() / md_path
    markdown = build_markdown(deck)
    if md_path:
        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text(markdown + "\n", encoding="utf-8")

    # 输出 docx
    build_docx(deck, out_path, args)

    print(f"build_speech v{VERSION}")
    print(f"  content：{content_path}")
    if args.profile:
        print(f"  profile：{Path(args.profile).expanduser()}")
    print(
        f"  讲稿：{len(deck.pages)} 页 / 正文 {speech_chars:,} 字 / "
        f"按 {deck.speed} 字/分钟约 {speech_minutes:.1f} 分钟（设定 {deck.duration_min} 分钟）"
    )
    print(
        f"  结构：背景补充 {len(deck.background)} 节 · 术语 {len(deck.terms)} 条 · "
        f"预设问答 {len(deck.qa)} 条 · 口头提示 {len(deck.tips)} 条"
    )
    print(f"  已写入：{out_path}")
    if md_path:
        print(f"  已写入：{md_path}")

    if speech_minutes < deck.duration_min * 0.75:
        print(
            f"⚠ [警告] 正文只有约 {speech_minutes:.1f} 分钟，低于设定 {deck.duration_min} 分钟；"
            "建议在 content.json 的页面 speech.text 里补细节，或在 speech.background 里补背景。",
            file=sys.stderr,
        )
    elif speech_minutes > deck.duration_min * 1.25:
        print(
            f"⚠ [警告] 正文约 {speech_minutes:.1f} 分钟，超过设定 {deck.duration_min} 分钟；"
            "删掉每页标「可跳过」的句子，或调低体积。",
            file=sys.stderr,
        )

    if args.pptx:
        pptx_path = Path(args.pptx).expanduser()
        info = read_pptx_info(pptx_path)
        if info is None:
            print("[提示] 没装 python-pptx，跳过 pptx 页数校验（不影响 docx 生成）。", file=sys.stderr)
        else:
            print(f"  pptx：{info['slides']} 页（讲稿 {len(deck.pages)} 页），备注共 {info['notes_chars']} 字")
            if info["slides"] != len(deck.pages):
                print(
                    f"[提示] PPT 页数（{info['slides']}）和讲稿页数（{len(deck.pages)}）不一致；"
                    "如果 PPT 多了致谢/备用页，记得在讲稿里补一句；如果少了，检查 content.json 是否漏页。",
                    file=sys.stderr,
                )

    return 0


ARGS: argparse.Namespace = argparse.Namespace(body_size=12)


if __name__ == "__main__":
    sys.exit(main())

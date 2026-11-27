#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""humanize_lint.py — 学院汇报 PPT / 讲稿的「说人话」检查器。

检查两种输入：
  1. content.json（format_spec v2，兼容 v1 的 blocks/bullets/lines 写法）
  2. Markdown / 文本 / .docx 讲稿

用法：
    python3 humanize_lint.py --content content.json
    python3 humanize_lint.py --content content.json --json report.json --strict
    python3 humanize_lint.py --content content.json --hints
    python3 humanize_lint.py --text 讲稿.md
    python3 humanize_lint.py --text 讲稿.docx --hints --json

规则来自 references/humanizer_cn.md：正则 + 词表，按 high / medium / low 分级。
high 是「几乎一定要改」的 AI 腔；medium 是「看上下文改」；low 是风格建议。
同一句话里被引号包住、或属于专业术语白名单（如「闭环控制」「显著性检验」）的命中会被放行。

退出码：
    0  检查完成（即使有命中；--strict 时另算）
    1  --strict 模式且存在 high 命中
    2  用法或文件错误
"""

from __future__ import annotations

import argparse
import json
import re
import statistics
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Iterator, Sequence
from xml.etree import ElementTree as ET

VERSION = "1.0"

SEVERITY_ORDER = {"high": 0, "medium": 1, "low": 2}
SEVERITY_LABEL = {"high": "high 必改", "medium": "medium 看语境", "low": "low 风格建议"}

# 句子切分：中文句末标点或换行
SENT_SPLIT_RE = re.compile(r"[^。！？；\n]+[。！？；]?")
# 判断「同一句」用的边界
SENT_BOUNDARY = "。！？；\n"
# 引号区间：规则命中若落在引号内视为引用，不判 AI 腔
QUOTE_SPAN_RE = re.compile(r"[「『“\"][^」』”\"]{0,120}[」』”\"]")
# 数字（用于「后面有具体数据就放行」的上下文豁免）
NUMBER_RE = re.compile(
    r"\d+(?:\.\d+)?\s*(?:%|％|倍|dBm?|bps|Mbps|Gbps|Hz|kHz|MHz|GHz|"
    r"ms|s|秒|分钟|小时|个|条|项|架|根|次|页|回合|万|轮|米|m|km|kg|dB)"
)
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
WORD_RE = re.compile(r"[A-Za-z]+|\d+(?:\.\d+)?")


# ---------------------------------------------------------------------------
# 规则表
# ---------------------------------------------------------------------------


@dataclass
class Rule:
    """一条检查规则。

    patterns    要抓的正则列表
    keep        命中句里出现这些正则就放行（上下文豁免 / 术语白名单）
    keep_prev   前置一句里出现这些也放行（如「值得注意的是」前文有数据）
    suggestion  一句话改写建议
    hint        给 --hints 的详细改法（可含正反例）
    kinds       适用对象：deck=PPT 文案，speech=讲稿/口语，text=普通文本
    """

    id: str
    name: str
    severity: str
    patterns: Sequence[str] = field(default_factory=tuple)
    keep: Sequence[str] = field(default_factory=tuple)
    keep_prev: Sequence[str] = field(default_factory=tuple)
    suggestion: str = ""
    hint: str = ""
    kinds: Sequence[str] = ("deck", "speech", "text")


def _r(*args, **kwargs) -> Rule:
    return Rule(*args, **kwargs)


# 术语 / 上下文白名单（与 humanizer_cn.md「专业术语保护」一节对应）
KEEP_CLOSED_LOOP = r"闭环(?:控制|系统|回路|反馈|增益|稳定性|跟踪|检测|调节|功率)|(?:控制|电流|电压|速度|功率|姿态)闭环|闭环(?:零|极)点"
KEEP_ECO = r"生态(?:系统|环境|保护|文明|位|修复|补偿)|产业生态系统"
KEEP_PARADIGM = r"(?:研究|科学|技术|计算|设计|工程|教育)范式|范式(?:转移|转换|革命)"
KEEP_SIGNIFICANCE = r"显著性|显著水平|显著差异|统计显著|显著相关|p\s*[<=＜＝]\s*0?\.\d+|置信区间|置信度"
KEEP_TARGET = r"(?:检测|识别|跟踪|感知)(?:概率|性能|指标)|概率(?:密度|分布)"
KEEP_INVEST = r"投资回报|收益率|回报率"

RULES: list[Rule] = [
    # ---- A. 先否定再肯定的「舞台腔」 ----
    _r(
        id="not_x_but_y",
        name="不是…而是（先否定再肯定）",
        severity="high",
        patterns=[
            r"不是[^。！？；\n]{0,40}?而是",
            r"并非[^。！？；\n]{0,40}?而是",
            r"不仅仅?是[^。！？；\n]{0,40}?(?:更是|而是)",
            r"不只是[^。！？；\n]{0,40}?(?:更是|而是)",
        ],
        suggestion="直接写后半句的结论。前半句的否定只在你确实要纠正听众的预期时才留。",
        hint=(
            "改法：删掉「不是 A，而是 B」里的 A，只留 B。\n"
            "反例：不是模型不行，而是参数没调好。\n"
            "正例：参数没调好，调完准确率从 61% 到 78%。"
        ),
    ),
    _r(
        id="not_only_but",
        name="不仅…而且 / 不但…还",
        severity="medium",
        patterns=[
            r"不仅[^。！？；\n]{0,30}?(?:而且|并且|还|更|也)",
            r"不但[^。！？；\n]{0,30}?(?:而且|并且|还|更|也)",
        ],
        suggestion="直接说第二层信息；「不仅」不提供事实，只放大语气。",
        hint="反例：不仅提升了速率，而且降低了时延。改：速率提上去了，时延也压下来了（给出数字更好）。",
    ),
    _r(
        id="rather_than",
        name="而不是（口头对比）",
        severity="medium",
        patterns=[r"而不是[^。！？；\n]{0,30}"],
        keep=[KEEP_CLOSED_LOOP],
        suggestion="如果对比两边都带信息可以留；只是换个说法就删掉否定的一半。",
        hint="反例：我们关心的是鲁棒性，而不是精度。改：我们主要看鲁棒性。",
    ),
    _r(
        id="worth_noting",
        name="值得注意的是 / 由此可见",
        severity="high",
        patterns=[r"值得注意的是", r"值得一提的是", r"不难发现", r"由此可见", r"可以看出"],
        keep=[KEEP_SIGNIFICANCE, r"\d+(?:\.\d+)?\s*(?:%|％|倍|dBm?|bps|Hz|个|条|项|架|根|次|页|回合|万|秒)"],
        keep_prev=[r"\d", r"图\s*\d", r"表\s*\d"],
        suggestion="删掉提示语，直接给数据或结论；数据自带说服力。",
        hint=(
            "反例：值得注意的是，MODE 的收敛速度比基线快 40%。\n"
            "正例：MODE 收敛快 40%（1.0 万回合到平台，基线 1.7 万回合）。\n"
            "豁免：前一句已经有具体数据时，本规则放行。"
        ),
    ),
    _r(
        id="attention_note",
        name="需要注意的是 / 应当指出",
        severity="medium",
        patterns=[r"需要注意的是", r"应当指出", r"需要指出的是", r"必须指出"],
        keep_prev=[r"\d"],
        suggestion="实验约束直接写进句子里，比如「这里只测了单基站」。",
        hint="反例：需要注意的是，本文只做了仿真。改：实验只有仿真，没有实测。",
    ),

    # ---- B. 总结腔 / 报告腔 ----
    _r(
        id="conclusion_smell",
        name="综上所述 / 总而言之",
        severity="high",
        patterns=[
            r"综上所述",
            r"总而言之",
            r"总的来说",
            r"一言以蔽之",
            r"概括而言",
            r"整体而言",
        ],
        suggestion="换成一句具体结论，或者直接删掉这四个字。",
        hint="反例：综上所述，本文方法性能更优。正例：三种偏好权重下，MODE 的和速率都比基线高 12% 以上。",
    ),
    _r(
        id="enum_triplet",
        name="首先…其次…最后（成套出现）",
        severity="high",
        patterns=[
            r"首先[^。！？；\n]{0,80}?其次[^。！？；\n]{0,80}?(?:最后|再次|第三)",
            r"其一[^。！？；\n]{0,60}?其二[^。！？；\n]{0,60}?(?:其三|最后)",
        ],
        kinds=("speech", "text"),
        suggestion="讲步骤就直接说「第一步…第二步…」；讲理由就一条一条说，不用序数词开场。",
        hint="反例：首先……其次……最后……。正例：第一步，把功率归一化；第二步，剩余时隙直飞终点。",
    ),
    _r(
        id="enum_single",
        name="首先 / 其次 / 最后（单用）",
        severity="medium",
        patterns=[r"(?:^|[。！？；\n])\s*(?:首先|其次|再次|最后)[，,]"],
        keep=[r"第一步|第二步|第三步|步骤"],
        kinds=("speech", "text"),
        suggestion="口语里不用序数词也能听出顺序；保留时后面接具体动作。",
        hint="反例：首先我们看背景。改：先看背景：低空经济到底缺什么。",
    ),
    _r(
        id="inflated_significance",
        name="拔高意义（奠定基础 / 填补空白 / 具有重大意义）",
        severity="high",
        patterns=[
            r"具有(?:重要|重大|深远)(?:的)?(?:理论|现实|实践|学术|指导)?意义",
            r"意义重大",
            r"填补(?:了)?[^。！？；\n]{0,8}空白",
            r"开辟(?:了)?[^。！？；\n]{0,8}(?:新|道路|方向|思路)",
            r"奠定(?:了|下)?[^。！？；\n]{0,10}基础",
            r"打下(?:了|下)?[^。！？；\n]{0,10}基础",
            r"提供(?:了)?(?:新|重要|有益|可靠)?的?(?:思路|参考|借鉴|启示|依据)",
            r"指明了?方向",
            r"(?:里程碑|划时代|革命性|前所未有)",
            r"具有(?:广阔|良好)的?(?:应用|发展)前景",
        ],
        suggestion="把「有什么意义」换成「解决了什么、好多少」；论文汇报里，数字比评价管用。",
        hint="反例：该工作填补了低空 ISAC 的空白，具有重要意义。正例：它把波束和轨迹放进同一次训练，和速率提升 18%。",
    ),
    _r(
        id="closer",
        name="一句话收尾（这说明 / 这告诉我们）",
        severity="medium",
        patterns=[
            r"这说明[^。！？；\n]{0,30}[。！？]",
            r"这告诉我们[^。！？；\n]{0,30}",
            r"这就是为什么[^。！？；\n]{0,30}",
            r"这就是[^。！？；\n]{0,12}的(?:意义|所在)",
        ],
        kinds=("speech", "text"),
        suggestion="收尾句如果只是重复上句，删掉；如果加了新事实，直接说事实。",
        hint="反例：这说明 MoE 很有效。正例：换 b 之后微调 2,000 回合就回到原水平，MODE-w 要重训 5 万回合。",
    ),
    _r(
        id="benwen",
        name="本文 / 本报告（讲稿里应说「这篇论文」）",
        severity="high",
        patterns=[r"本文(?:提出|针对|研究|主要|设计|构建|工作)", r"本报告(?:提出|主要|认为)"],
        kinds=("speech",),
        suggestion="口头讲稿里说「这篇论文」「作者」；「本文」是论文书面语。",
        hint="反例：本文提出了 MODE。正例：这篇论文提出了一套 MODE。",
    ),
    _r(
        id="aiming",
        name="旨在 / 致力于 / 聚焦于",
        severity="medium",
        patterns=[r"旨在", r"致力于", r"着眼于", r"聚焦于", r"着手"],
        suggestion="改成「要做的是…」，「旨在」是论文摘要腔。",
        hint="反例：本文旨在提升系统性能。正例：它想解决的是：同一个基站怎么既供网又感知。",
    ),
    _r(
        id="system_nominalization",
        name="体系化 / 形成体系",
        severity="medium",
        patterns=[
            r"构建(?:了|起)?[^。！？；\n]{0,8}体系",
            r"形成(?:了)?[^。！？；\n]{0,8}体系",
            r"体系化",
            r"系统化(?:地)?(?:推进|建设|研究)",
        ],
        suggestion="说清楚具体做了哪几件事，别用「体系」概括。",
        hint="反例：构建了完整的低空 ISAC 技术体系。正例：它做了三件事：问题建模、MoE 结构、整回合回放。",
    ),
    _r(
        id="staged_authority",
        name="众所周知 / 毋庸置疑 / 显而易见",
        severity="high",
        patterns=[r"众所周知", r"毋庸置疑", r"毫无疑问", r"显而易见", r"不言而喻"],
        suggestion="删掉。真正显然的事不用铺垫，需要铺垫的事说明它不显然。",
        hint="反例：众所周知，DRL 训练不稳定。正例：DRL 训练不稳定，这篇用整回合回放来压方差。",
    ),

    # ---- C. 互联网黑话 / 空泛形容词 ----
    _r(
        id="buzzword",
        name="赋能 / 抓手 / 闭环 / 生态 / 范式（黑话）",
        severity="high",
        patterns=[
            r"赋能",
            r"抓手",
            r"闭环",
            r"生态",
            r"范式",
            r"底层逻辑",
            r"顶层设计",
            r"组合拳",
            r"护城河",
        ],
        keep=[KEEP_CLOSED_LOOP, KEEP_ECO, KEEP_PARADIGM],
        suggestion="换成具体动作：谁给谁提供了什么、哪一步接哪一步。",
        hint=(
            "反例：通过技术赋能形成闭环生态，打造新范式。\n"
            "正例：把感知结果接回波束分配，每 40 个时隙更新一次。\n"
            "豁免：闭环控制、闭环系统、生态系统、研究范式等专业用法不判。"
        ),
    ),
    _r(
        id="consulting_jargon",
        name="落地 / 打通 / 对齐 / 颗粒度",
        severity="low",
        patterns=[r"打通", r"落地", r"对齐", r"颗粒度", r"发力", r"协同增效", r"提质增效"],
        keep=[r"对齐(?:到|误差|精度|方式|时间|矩阵|算法)|相位对齐|时序对齐|波束对齐|载波对齐|帧对齐"],
        suggestion="汇报里说具体动作：接上了、跑通了、每帧对齐一次。",
        hint="反例：打通算法与工程落地。正例：算法输出直接接进现有的功率分配模块。",
    ),
    _r(
        id="empty_adjective",
        name="空泛形容词（显著 / 大幅 / 优异 / 充分）",
        severity="low",
        patterns=[
            r"显著(?:提升|提高|改善|增强|优于|下降|降低)",
            r"大幅(?:提升|提高|改善|增强|降低|减少)",
            r"极大(?:地)?(?:提升|提高|改善|增强)",
            r"优异",
            r"卓越",
            r"亮眼",
            r"出色",
            r"强大",
            r"充分(?:证明|体现|说明|发挥)",
            r"有效(?:提升|提高|改善|解决|缓解)",
            r"至关重要",
            r"核心关键",
        ],
        keep=[KEEP_SIGNIFICANCE],
        suggestion="把形容词换成数字或对比：提升了多少、比谁好、在哪个设置下。",
        hint="反例：性能显著提升。正例：和速率从 1.8×10² 提到 2.1×10² bps/Hz。",
    ),
    _r(
        id="vague_quantity",
        name="模糊数量（大量 / 许多 / 一定程度上）",
        severity="low",
        patterns=[
            r"(?:大量|众多|许多|若干)(?:的)?(?:研究|工作|实验|数据|方法)",
            r"一定程度上",
            r"某种程度上",
            r"在某种程度",
            r"相对较好",
        ],
        suggestion="能给数字就给数字；给不了就删掉程度副词。",
        hint="反例：大量实验证明了有效性。正例：4 组仿真、5 万回合训练，4 个基线都跑了一遍。",
    ),
    _r(
        id="intensity_adverb",
        name="程度副词堆叠（非常 / 十分 / 极其）",
        severity="low",
        patterns=[r"非常(?:重要|明显|显著|关键)", r"十分(?:重要|明显|显著)", r"极其(?:重要|复杂|困难)"],
        keep=[KEEP_SIGNIFICANCE],
        suggestion="删掉程度词，让事实说话；口语里可以用一次，不要连用。",
        hint="反例：这个问题非常复杂。正例：这个问题的难点在非凸，加上目标会跑。",
    ),
    _r(
        id="buzz_abstract",
        name="抽象大词（多维度 / 全方位 / 深度融合）",
        severity="low",
        patterns=[r"多维度", r"全方位", r"深层次", r"立体化", r"深度融合", r"系统性(?:地)?(?:推进|提升|建设)"],
        suggestion="换成具体维度或具体做法。",
        hint="反例：多维度提升系统性能。正例：和速率、感知 SNR 两个指标一起优化。",
    ),

    # ---- D. 翻译腔 / 论文腔 ----
    _r(
        id="passive_translation",
        name="被广泛… / 被认为是（翻译腔被动）",
        severity="medium",
        patterns=[
            r"被(?:广泛|普遍|大量)?(?:研究|关注|提出|应用|采用|认为|证明|验证|讨论)",
            r"被广泛应用",
            r"被认为是",
        ],
        suggestion="补出主语：谁提出、谁在用，或者直接说「大家常用」。",
        hint="反例：该问题被广泛研究。正例：这个问题从 2019 年起有一批工作在做，[8]–[21] 都是。",
    ),
    _r(
        id="nominal_translation",
        name="进行…研究 / 通过…的方式 / 在…方面",
        severity="low",
        patterns=[
            r"进行(?:了)?[^。！？；\n]{0,6}(?:研究|分析|实验|测试|讨论|优化|验证|比较)",
            r"通过[^。！？；\n]{0,10}的?方式",
            r"在[^。！？；\n]{0,8}方面",
            r"对[^。！？；\n]{0,12}进行[^。！？；\n]{0,6}",
        ],
        suggestion="中文直接说动词：分析、测试、比较，别用「进行 + 名词」。",
        hint="反例：对系统性能进行了分析。正例：我们测了系统性能，重点看两个指标。",
    ),
    _r(
        id="chat_residue",
        name="聊天机器人残留（希望对你有帮助 / 以下是）",
        severity="high",
        patterns=[
            r"希望对(?:你|您|大家)(?:有)?(?:所)?帮助",
            r"以下是[^。！？；\n]{0,10}(?:内容|介绍|总结)",
            r"如需(?:要)?(?:了解|帮助)",
            r"感谢(?:大家|你的)?阅读",
            r"如有(?:任何)?(?:疑问|问题)(?:欢迎|随时)",
        ],
        suggestion="删掉服务式话术，讲稿只留内容。",
        hint="反例：以下是本文的方法介绍，希望对你有帮助。正例：（直接开始）先说方法要解决什么。",
    ),
    _r(
        id="truly",
        name="真正 / 本质上 / 核心在于（装饰性强调）",
        severity="low",
        patterns=[r"真正(?:的)?(?:问题|价值|难点|意义|优势|区别)", r"本质上", r"根本上", r"核心在于", r"关键在于"],
        suggestion="删掉强调词，直接说是什么；一篇讲稿里别超过两次。",
        hint="反例：真正的难点在于耦合。正例：难点是波束和轨迹耦合在一起。",
    ),

    # ---- E. 排比 / 破折号 / 标题 ----
    _r(
        id="parallel_triad",
        name="排比三连",
        severity="medium",
        patterns=[
            r"既要[^。！？；\n]{0,30}?又要[^。！？；\n]{0,30}?还要",
            r"从[^。！？；\n]{0,12}?到[^。！？；\n]{0,12}?再到",
            r"[，,]是[^。！？；，\n]{0,10}[，,]是[^。！？；，\n]{0,10}[，,]也是",
            r"不仅[^。！？；\n]{0,20}?还[^。！？；\n]{0,20}?更",
        ],
        suggestion="三连只保留信息量最大的两项；第三项往往是凑数。",
        hint="反例：既要低时延，又要高可靠，还要低成本。正例：主要约束是时延，成本排在后面。",
    ),
    _r(
        id="dash_chain",
        name="破折号串句（一段里出现两次以上）",
        severity="high",
        patterns=[],
        suggestion="把破折号换成句号、逗号或括号；一段最多留一个。",
        hint="反例：这个方法——不用重训——换 b 只微调一层——很快。正例：这个方法不用重训，换 b 时只微调一层，两千回合就收敛。",
        kinds=("deck", "speech", "text"),
    ),
    _r(
        id="dash_single",
        name="破折号（单次，风格提示）",
        severity="low",
        patterns=[],
        suggestion="确认这里必须用破折号；能用逗号或句号就用标点断开。",
        hint="PPT 和讲稿里破折号容易显得像翻译腔；一个页面最多一处。",
    ),
    _r(
        id="symmetric_heading",
        name="对称小标题（A 与 B）",
        severity="low",
        patterns=[],
        suggestion="标题改成一句结论，比如「三种方法各缺一块」。",
        hint="反例：研究背景与研究意义。正例：低空经济为什么需要通感一体。",
        kinds=("deck",),
    ),
    _r(
        id="boring_title",
        name="空标题（研究背景 / 实验设置）",
        severity="medium",
        patterns=[
            r"^\s*(?:研究背景|研究意义|研究方法|研究内容|技术路线|实验设置|实验结果|"
            r"结论与展望|总结与展望|现状分析|相关工作|系统模型|问题建模)\s*$"
        ],
        suggestion="标题写结论：读者只看标题就该知道这页说什么。",
        hint="反例：实验设置。正例：6 架无人机、8 根天线，一回合 40 个时隙。",
        kinds=("deck",),
    ),

    # ---- F. 问答 / 设问 / 夸张 ----
    _r(
        id="staged_question",
        name="先设问再自答（为什么…？原因在于）",
        severity="medium",
        patterns=[r"为什么[^。！？\n]{0,24}？[^。！？\n]{0,10}(?:原因|答案|关键)(?:在于|是)"],
        kinds=("speech", "text"),
        suggestion="问题可以直接问，但答案别用「原因在于」起头。",
        hint="反例：为什么用 MoE？原因在于偏好权重会变。正例：为什么用 MoE？因为 b 会变，重训一次要 5 万回合。",
    ),
    _r(
        id="overclaim",
        name="绝对化表述（完美 / 彻底 / 领先）",
        severity="high",
        patterns=[
            r"全面领先",
            r"遥遥领先",
            r"完美(?:解决|契合|复现)",
            r"彻底解决",
            r"一劳永逸",
            r"颠覆(?:性)?(?:的)?(?:技术|方法|方案)",
        ],
        suggestion="给出成立的条件和对比对象；论文汇报里少用「全面领先」。",
        hint="反例：MODE 全面领先所有基线。正例：两项指标上 MODE 和 MODE-w 并列最高，MODE 的优势在收敛速度。",
    ),
    _r(
        id="first_claim",
        name="首次提出（新颖性断言）",
        severity="medium",
        patterns=[r"首次(?:提出|实现|验证|将)"],
        suggestion="确认原文确实这么写；不确定就改成「较早提出」或删掉。",
        hint="反例：本文首次把 MoE 用到 ISAC。正例：作者把 MoE 用到了多目标 ISAC 上（原文没有用「首次」）。",
    ),
    _r(
        id="weak_claim",
        name="空泛判断（效果好 / 表现不错）",
        severity="low",
        patterns=[r"(?:效果|性能|表现)(?:很|较|比较)?(?:好|不错|理想|优异)"],
        suggestion="换成指标和数字。",
        hint="反例：效果很好。正例：感知 SNR 稳定在 0.9×10² 以上。",
    ),
    _r(
        id="hedge_stack",
        name="模糊叠加（可能…或许 / 大概可能）",
        severity="medium",
        patterns=[r"可能[^。！？；\n]{0,12}?或许", r"或许[^。！？；\n]{0,12}?可能", r"大概可能", r"似乎可能"],
        suggestion="留一个不确定词就够了，写清楚不确定在哪。",
        hint="反例：这可能是或许因为训练不稳定。正例：我猜是训练不稳定，论文没给消融验证。",
    ),
    _r(
        id="exclamation",
        name="感叹号（汇报语气过冲）",
        severity="low",
        patterns=[r"！"],
        suggestion="学术汇报用句号，最多保留一个感叹号在结尾致谢。",
        hint="反例：性能提升太明显了！正例：性能提升明显，见右图。",
    ),

    # ---- G. 排版 / 占位符 ----
    _r(
        id="placeholder",
        name="占位符残留（XX / xxx / 待补充 / TODO / 【】）",
        severity="high",
        patterns=[r"待补充", r"待完善", r"TODO", r"xxx+", r"XX+", r"【\s*】", r"\[\s*\]", r"……\s*…+"],
        suggestion="占位符不能出现在成稿里；没有数据就删掉这一条。",
        hint="成稿里出现「待补充」会直接被听众看到；要么补数据，要么删这一页。",
    ),
    _r(
        id="markdown_bold",
        name="正文里的 Markdown 加粗符号 **",
        severity="medium",
        patterns=[r"\*\*[^*\n]{1,30}\*\*"],
        suggestion="PPT 里加粗交给生成器；正文里不要写 Markdown 标记。",
        hint="反例：**核心结论**是……正例：核心结论是……",
        kinds=("deck",),
    ),
    _r(
        id="emoji",
        name="Emoji / 装饰符号",
        severity="high",
        patterns=[r"[\U0001F300-\U0001FAFF\u2600-\u27BF\u2B50\u2705\u274C]"],
        suggestion="学术汇报不用 emoji；要强调就换字体颜色或加粗。",
        hint="反例：🚀 实验结果。正例：实验结果。",
    ),
    # 弯引号不做单处判断：中文排版用弯引号很正常，只在密集出现时提示统一风格。
    _r(
        id="curly_quote",
        name="弯引号密集（格式提醒）",
        severity="low",
        patterns=[],
        suggestion="引号风格统一即可；中文正文可考虑用「」。",
        hint="引号本身不是 AI 腔，只是排版一致性提醒。",
    ),
    # ---- H. 写给自己/写手的工作备注漏给观众 ----
    _r(
        id="internal_note_leak",
        name="工作备注/内部提示上屏",
        severity="high",
        patterns=[
            r"请(?:再)?核对",
            r"引用前",
            r"待确认|待补充|待补|待定",
            r"TODO|FIXME|TBD",
            r"占位|草稿|自检|核对清单|勿外传",
            r"(?:本页|此处)(?:待|需|要)(?:补|改|确认)",
        ],
        suggestion="删掉这句：观众不需要看写手的工作备注。需要保留的提醒放进讲稿备注或交付说明。",
        hint="写给自己/写手的话（请再核对、待确认、占位、引用前核对）不能出现在页面上。\n反例：数据来源：LetPub（2026-11），引用前请再核对一次。\n正例：数据来源：LetPub（2026-11）。",
        kinds=("deck", "speech", "text"),
    ),
]

RULE_BY_ID = {rule.id: rule for rule in RULES}


# ---------------------------------------------------------------------------
# 命中结构
# ---------------------------------------------------------------------------


@dataclass
class Hit:
    file: str
    field: str
    context: str
    rule_id: str
    rule_name: str
    severity: str
    matched: str
    sentence: str
    suggestion: str
    hint: str
    offset: int = 0

    def to_dict(self) -> dict:
        return {
            "file": self.file,
            "field": self.field,
            "rule": self.rule_id,
            "rule_name": self.rule_name,
            "severity": self.severity,
            "matched": self.matched,
            "sentence": self.sentence,
            "suggestion": self.suggestion,
            "hint": self.hint,
            "offset": self.offset,
        }


# ---------------------------------------------------------------------------
# 读文本（md / txt / docx，docx 只用标准库）
# ---------------------------------------------------------------------------

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def read_docx_text(path: Path) -> str:
    """不依赖 python-docx，直接从 docx 里抽段落文本。"""
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise SystemExit(f"[错误] 读不了 docx：{path}（{exc}）") from exc
    root = ET.fromstring(xml)
    paras: list[str] = []
    for p in root.iter(W_NS + "p"):
        texts = [t.text or "" for t in p.iter(W_NS + "t")]
        if texts:
            paras.append("".join(texts))
    return "\n".join(paras)


def read_text_target(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return read_docx_text(path)
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------


def han_count(text: str) -> int:
    return len(CJK_RE.findall(text))


def speak_len(text: str) -> int:
    """口播字数：汉字 + 英文词/数字，不含标点空白。用于估算时长。"""
    return han_count(text) + len(WORD_RE.findall(text))


def split_sentences(text: str) -> list[str]:
    return [s.strip() for s in re.split(r"[。！？；\n]+", text) if s.strip()]


def sentence_span(text: str, pos: int) -> str:
    """取 pos 所在的句子（按中文句末标点或换行切）。"""
    start = pos
    while start > 0 and text[start - 1] not in SENT_BOUNDARY:
        start -= 1
    end = pos
    while end < len(text) and text[end] not in SENT_BOUNDARY:
        end += 1
    if end < len(text):
        end += 1
    return text[start:end].strip()


def prev_sentence(text: str, pos: int) -> str:
    sentence = sentence_span(text, pos)
    idx = text.find(sentence[:-1]) if sentence else -1
    if idx <= 0:
        return ""
    before = text[:idx]
    sentences = re.findall(r"[^。！？；\n]+[。！？；]?", before)
    return sentences[-1].strip() if sentences else ""


def in_quote(text: str, pos: int) -> bool:
    for match in QUOTE_SPAN_RE.finditer(text):
        if match.start() < pos < match.end():
            return True
    return False


def count_occurrences(text: str, needle: str) -> int:
    return text.count(needle)


def _matches(pattern: str, text: str) -> bool:
    return re.search(pattern, text) is not None


# ---------------------------------------------------------------------------
# 通用检查
# ---------------------------------------------------------------------------


def check_unit(
    units: Sequence[tuple[str, str, str]],
    file_label: str,
    hits: list[Hit],
    stats_text: list[str] | None = None,
) -> None:
    """units: [(field, kind, text), ...]；kind in {deck speech text}。"""
    seen: set[tuple[str, str, str]] = set()
    for field_path, kind, text in units:
        if not text or not text.strip():
            continue
        if stats_text is not None:
            stats_text.append(text)
        text = text.strip()

        for rule in RULES:
            if kind not in rule.kinds:
                continue

            # 计数型规则：破折号
            if rule.id in ("dash_chain", "dash_single"):
                n_dash = text.count("——")
                if rule.id == "dash_chain" and n_dash >= 2:
                    key = (field_path, rule.id, "dash")
                    if key not in seen:
                        seen.add(key)
                        hits.append(
                            Hit(
                                file=file_label,
                                field=field_path,
                                context=kind,
                                rule_id=rule.id,
                                rule_name=rule.name,
                                severity="high",
                                matched=f"—— × {n_dash}",
                                sentence=text[:80],
                                suggestion=rule.suggestion,
                                hint=rule.hint,
                            )
                        )
                elif rule.id == "dash_single" and n_dash == 1:
                    key = (field_path, rule.id, "dash")
                    if key not in seen:
                        seen.add(key)
                        hits.append(
                            Hit(
                                file=file_label,
                                field=field_path,
                                context=kind,
                                rule_id=rule.id,
                                rule_name=rule.name,
                                severity="low",
                                matched="——",
                                sentence=text[:80],
                                suggestion=rule.suggestion,
                                hint=rule.hint,
                            )
                        )
                continue

            if not rule.patterns:
                # 计数/聚合型规则在 check_prose_rules 里处理
                continue

            # 空标题 / 对称标题只检查「页面主标题」，章节名和标签不算
            if rule.id in ("boring_title", "symmetric_heading") and not re.search(
                r"pages\[\d+\]\(.*\)\.title$", field_path
            ):
                continue

            # 对称小标题：仅标题、仅较短的文本
            if rule.id == "symmetric_heading":
                stripped = re.sub(r"^\s*\d+(?:\.\d+)*[、.．]\s*", "", text).strip()
                m = re.fullmatch(r"([\u4e00-\u9fff]{2,8})(?:与|和|及)([\u4e00-\u9fff]{2,8})", stripped)
                if m and abs(len(m.group(1)) - len(m.group(2))) <= 1:
                    key = (field_path, rule.id, stripped)
                    if key not in seen:
                        seen.add(key)
                        hits.append(
                            Hit(
                                file=file_label,
                                field=field_path,
                                context=kind,
                                rule_id=rule.id,
                                rule_name=rule.name,
                                severity="low",
                                matched=stripped,
                                sentence=text,
                                suggestion=rule.suggestion,
                                hint=rule.hint,
                            )
                        )
                continue

            for pattern in rule.patterns:
                try:
                    finder: Iterable[re.Match[str]] = re.finditer(pattern, text)
                except re.error:  # 规则写错不阻断检查
                    continue
                for match in finder:
                    sentence = sentence_span(text, match.start())
                    key = (field_path, rule.id, sentence)
                    if key in seen:
                        continue
                    # 引号内是引用，不算
                    if in_quote(text, match.start()):
                        continue
                    # 上下文豁免（术语白名单等）
                    if any(_matches(keep, sentence) for keep in rule.keep):
                        continue
                    if rule.keep_prev and any(
                        _matches(keep, prev_sentence(text, match.start())) for keep in rule.keep_prev
                    ):
                        continue
                    seen.add(key)
                    hits.append(
                        Hit(
                            file=file_label,
                            field=field_path,
                            context=kind,
                            rule_id=rule.id,
                            rule_name=rule.name,
                            severity=rule.severity,
                            matched=match.group(0),
                            sentence=sentence[:120],
                            suggestion=rule.suggestion,
                            hint=rule.hint,
                            offset=match.start(),
                        )
                    )


def check_prose_rules(
    units: Sequence[tuple[str, str, str]],
    file_label: str,
    hits: list[Hit],
) -> None:
    """句长、句长均匀度、连接词密度、重复句首、强调词频次等整段规则。"""
    # 1) 长句：只对讲稿/文本
    for field_path, kind, text in units:
        if kind not in ("speech", "text") or not text:
            continue
        long_sentences: list[str] = []
        for sentence in split_sentences(text):
            n = han_count(sentence)
            if n > 80:
                hits.append(
                    Hit(
                        file=file_label,
                        field=field_path,
                        context=kind,
                        rule_id="very_long_sentence",
                        rule_name="超长句（80 字以上）",
                        severity="high",
                        matched=f"{n} 字",
                        sentence=sentence[:120],
                        suggestion="拆成两到三句；口语念不出来就是太长。",
                        hint="念一遍，换气点处直接断句。目标：80 字以上为 0。",
                    )
                )
            elif n > 60:
                long_sentences.append(sentence)
        if len(long_sentences) >= 6:
            hits.append(
                Hit(
                    file=file_label,
                    field=field_path,
                    context=kind,
                    rule_id="long_sentence",
                    rule_name="长句偏多（60–80 字）",
                    severity="medium",
                    matched=f"{len(long_sentences)} 句",
                    sentence=long_sentences[0][:100],
                    suggestion="优先拆这些句子；一句只留一个意思。",
                    hint="目标：60 字以上句子占全篇 5%–10%，单句不超过 80 字。",
                )
            )
            break

    # 2) 整篇统计：句长均匀度、连接词密度、重复句首
    speech_texts = [text for _, kind, text in units if kind in ("speech", "text") and text]
    if not speech_texts:
        return
    whole = "\n".join(speech_texts)
    sentences = split_sentences(whole)
    lengths = [han_count(s) for s in sentences]
    if len(lengths) >= 20:
        mean = statistics.mean(lengths)
        stdev = statistics.pstdev(lengths)
        # 口头稿天然短句多（参考真人讲稿：平均 16 字、短句 70%），
        # 真正像 AI 的是「句长居中且几乎没有起伏」：18–34 字 + 标准差 < 9。
        if 18 <= mean <= 34 and stdev < 9:
            hits.append(
                Hit(
                    file=file_label,
                    field="全文统计",
                    context="text",
                    rule_id="sentence_uniform",
                    rule_name="句子长度太匀（AI 节奏）",
                    severity="medium",
                    matched=f"平均 {mean:.1f} 字，标准差 {stdev:.1f}",
                    sentence=f"共 {len(lengths)} 句，短句占比 {sum(1 for x in lengths if x <= 20) / len(lengths):.0%}",
                    suggestion="把 2–3 个相邻短句合成中长句，再补一句解释或例子，让句长有起伏。",
                    hint="目标（zh-ai-rate-rewrite）：平均 26–32 字，标准差 14–18，20 字以内短句 35%–45%；口头稿可以更短，但不能全一样长。",
                )
            )
        elif mean > 38:
            hits.append(
                Hit(
                    file=file_label,
                    field="全文统计",
                    context="text",
                    rule_id="sentence_uniform",
                    rule_name="句子平均太长",
                    severity="medium",
                    matched=f"平均 {mean:.1f} 字",
                    sentence=f"共 {len(lengths)} 句",
                    suggestion="优先拆 80 字以上的句子，再把长句改成两句短句。",
                    hint="口语节奏：一句 20–40 字，最多 3 个逗号。",
                )
            )

    # 连接词密度（每 500 字超过 4 个算密）
    connectors = ["因此", "然而", "此外", "同时", "综上", "另外", "再者", "从而"]
    total_conn = sum(whole.count(w) for w in connectors)
    per_500 = total_conn / max(han_count(whole), 1) * 500
    if per_500 > 4:
        hits.append(
            Hit(
                file=file_label,
                field="全文统计",
                context="text",
                rule_id="connector_density",
                rule_name="连接词密度偏高",
                severity="medium",
                matched=f"{total_conn} 个 / {han_count(whole)} 字",
                sentence="、".join(f"{w}×{whole.count(w)}" for w in connectors if whole.count(w)),
                suggestion="删掉一半连接词，改用短句直接推进；口语里停顿比「因此」自然。",
                hint="每 500 字最多 4 个「因此/然而/此外/同时」。",
            )
        )

    # 重复句首
    openers = [
        s[:2] if re.fullmatch(r"[\u4e00-\u9fff]{2}", s[:2]) else None
        for s in sentences
        if len(s) >= 4
    ]
    for i in range(len(openers) - 2):
        trio = sentences[i : i + 3]
        # 封面的「标签：值」信息行不算重复句首
        if all("：" in s for s in trio):
            continue
        if openers[i] is not None and openers[i] == openers[i + 1] == openers[i + 2]:
            hits.append(
                Hit(
                    file=file_label,
                    field="全文统计",
                    context="text",
                    rule_id="repeat_opening",
                    rule_name="连续句子同样开头",
                    severity="low",
                    matched=f"「{openers[i]}」×3",
                    sentence=" / ".join(sentences[i : i + 3])[:100],
                    suggestion="合并句子或换主语，别让三句都是一个开头。",
                    hint="口语里连着三句同开头，听着像念稿。",
                )
            )
            break

    # 弯引号密集：全篇出现 10 次以上时提示统一排版
    n_curly = whole.count("“") + whole.count("”")
    if n_curly >= 10:
        hits.append(
            Hit(
                file=file_label,
                field="全文统计",
                context="text",
                rule_id="curly_quote",
                rule_name="弯引号密集（格式提醒）",
                severity="low",
                matched=f"“” × {n_curly}",
                sentence="中文排版允许弯引号，此处仅提示统一风格。",
                suggestion=RULE_BY_ID["curly_quote"].suggestion,
                hint=RULE_BY_ID["curly_quote"].hint,
            )
        )

    # 强调词频次（zh-ai-rate-rewrite 的目标区间）
    for word, limit in (("但是", 20), ("真正", 5), ("而是", 6)):
        n = whole.count(word)
        if n > limit:
            hits.append(
                Hit(
                    file=file_label,
                    field="全文统计",
                    context="text",
                    rule_id=f"freq_{word}",
                    rule_name=f"「{word}」出现过多",
                    severity="medium",
                    matched=f"{word} × {n}",
                    sentence=f"超过建议上限 {limit} 次",
                    suggestion=f"换直接陈述或删掉一部分「{word}」。",
                    hint="参考 zh-ai-rate-rewrite：但是 ≤20、真正 ≤5、而是 ≤6（按一篇 8,000 字稿算；短稿按比例收紧）。",
                )
            )


# ---------------------------------------------------------------------------
# content.json 抽取
# ---------------------------------------------------------------------------


def _page_label(index: int, page: dict) -> str:
    kind = page.get("type", "?")
    title = str(page.get("title") or page.get("q") or "").strip()
    short = title[:18] + ("…" if len(title) > 18 else "")
    return f"pages[{index}](第{index + 1}页·{kind}{('·' + short) if short else ''})"


def check_deck_rhythm(data: dict, file_label: str, hits: list[Hit]) -> None:
    """整篇节奏检查：数字对仗标题、导语覆盖率、标题/导语句式整齐度。

    单句看都没问题，合起来「整齐得不像人」才是 AI 味：这是 LLM 的默认输出节奏。
    """
    pages = [p for p in (data.get("pages") or []) if isinstance(p, dict)]
    if len(pages) < 6:
        return

    def _plain(x):
        import re as _re
        return _re.sub(r"<[^>]+>", "", str(x or "")).strip()

    titles = [_plain(p.get("title")) for p in pages if _plain(p.get("title"))]
    leads = [_plain(p.get("lead")) for p in pages if _plain(p.get("lead"))]
    n = len(pages)

    # 1) 导语覆盖率：每页都配一句「金句导语」是 LLM 的默认做法
    if len(leads) / max(n, 1) > 0.6:
        hits.append(
            Hit(
                file=file_label, field="全文统计", context="deck",
                rule_id="lead_every_page",
                rule_name="几乎每页都有导语（AI 默认节奏）",
                severity="medium",
                matched=f"{len(leads)}/{n} 页有 lead",
                sentence="、".join(leads[:3])[:100],
                suggestion="删掉一半页面的导语；导语只留给真正需要一句结论的页，其他页直接上内容。",
                hint="人写 30 页不会每页都配一句对仗导语。建议 lead 覆盖率 ≤ 40%。",
            )
        )

    # 2) 数字对仗标题：「两个期望，六条约束」「三句话 + 一组数字」「四个缺口……」
    # 只抓真正的模板句式：数量词打头的短对仗、或已知的高频套话
    strong_re = re.compile(
        r"^(?:[一二三四五六七八九十两]\s*[个条句种步块处点件篇])"
        r"|(?:[一二三四五六七八九十两]\s*[个条句种步块处点件篇][^，。；]{0,10}[，、][^，。；]{0,6}"
        r"[一二三四五六七八九十两]\s*[个条句种步块处点件篇])"
    )
    known = re.compile(
        r"各缺一块|各管一段|各去掉一块|正好是下一步|一组数字|三句话|三处可迁移|四个缺口|三条结论|"
        r"两个期望|六条约束|五个对比|几个要素|几条结论|几处可迁移"
    )
    para_titles = [t for t in titles if strong_re.search(t) or known.search(t)]
    para_leads = [l for l in leads if strong_re.search(l) or known.search(l)]
    ratio = (len(para_titles) + len(para_leads)) / max(len(titles) + len(leads), 1)
    if len(para_titles) + len(para_leads) >= 4 and ratio > 0.2:
        hits.append(
            Hit(
                file=file_label, field="全文统计", context="deck",
                rule_id="number_parallel_titles",
                rule_name="数字对仗标题/导语过多",
                severity="high",
                matched=f"{len(para_titles)} 个标题 + {len(para_leads)} 个导语",
                sentence="；".join((para_titles + para_leads)[:5])[:140],
                suggestion="换掉大部分数字对仗句：改成陈述事实、提问、条件或直接给结论；同类句式全篇最多留 2–3 处。",
                hint="人不会 30 页维持「X 个 Y，Z 个 W」的纪律。参考改法：\n反例：两个期望，六条约束 / 三句话记住结果 / 四个缺口 / 三条结论，四个缺口\n正例：问题写成双目标，约束一共六条 / 结果小结 / 还没解决的问题 / 下一步",
            )
        )

    # 3) 标题长度太匀（标准差过小）
    if len(titles) >= 10:
        tlen = [len(t) for t in titles]
        sd = statistics.pstdev(tlen)
        if sd < 3.5:
            hits.append(
                Hit(
                    file=file_label, field="全文统计", context="deck",
                    rule_id="title_uniform",
                    rule_name="标题长度太整齐",
                    severity="medium",
                    matched=f"平均 {statistics.mean(tlen):.1f} 字，标准差 {sd:.1f}",
                    sentence="；".join(titles[:3])[:100],
                    suggestion="让标题长短错开：有的 6 字，有的 20 字；长标题可以带冒号补一句。",
                    hint="人写标题长短不齐；标准差 <3.5 说明是按模板批量产的。",
                )
            )

    # 4) 导语句式单一：长度接近 + 都只有一个逗号
    if len(leads) >= 6:
        lens = [len(l) for l in leads]
        one_comma = sum(1 for l in leads if l.count("，") == 1)
        if statistics.pstdev(lens) < 5 and one_comma / len(leads) > 0.6:
            hits.append(
                Hit(
                    file=file_label, field="全文统计", context="deck",
                    rule_id="lead_uniform",
                    rule_name="导语句式太单一",
                    severity="medium",
                    matched=f"{len(leads)} 句，长度标准差 {statistics.pstdev(lens):.1f}，单逗号 {one_comma} 句",
                    sentence="；".join(leads[:3])[:120],
                    suggestion="有的导语写长一点（带个数字或条件），有的直接删；不要每句都是「A，B」两段式。",
                    hint="人写导语有的 8 字，有的 36 字，有的干脆没有。",
                )
            )

    # 5) 连续多页都有导语
    streak, best = 0, 0
    seq = []
    for p in pages:
        if _plain(p.get("lead")):
            streak += 1
            best = max(best, streak)
        else:
            streak = 0
    if best >= 6:
        hits.append(
            Hit(
                file=file_label, field="全文统计", context="deck",
                rule_id="lead_streak",
                rule_name="连续多页都配导语",
                severity="medium",
                matched=f"连续 {best} 页",
                sentence="连续的「标题 + 一句导语 + 要点」结构读起来像模板",
                suggestion="每 5–6 页至少打断一次：有的页直接上图，有的页只有要点，有的页只有一句话。",
                hint="节奏要有起伏：密页、疏页、图页交替，不要 30 页一个模式。",
            )
        )


def check_deck_structure(data: dict, file_label: str, hits: list[Hit]) -> None:
    """结构性指纹：词表查不出来、但人一眼能看出「机器排的」的重复结构。"""
    import re as _re
    pages = [p for p in (data.get("pages") or []) if isinstance(p, dict)]
    if len(pages) < 6:
        return

    def plain(x):
        return _re.sub(r"<[^>]+>", "", str(x or "")).strip()

    def shape(text):
        out = []
        for ch in text:
            if "\u4e00" <= ch <= "\u9fff":
                out.append("C")
            elif ch.isdigit():
                out.append("9")
            elif ch.isalpha() and ch.isascii():
                out.append("X")
            elif ch.strip():
                out.append(ch)
        return "".join(out)

    def visible(pg):
        parts = [plain(pg.get("title")), plain(pg.get("lead"))]
        for key in ("items", "bullets"):
            for it in pg.get(key) or []:
                parts += [plain(it.get("k")) if isinstance(it, dict) else "",
                          plain(it.get("t") or it.get("p")) if isinstance(it, dict) else plain(it)]
        for key in ("lines", "blocks", "steps", "takeaway", "note", "tnote", "sub"):
            val = pg.get(key)
            if isinstance(val, list):
                for it in val:
                    if isinstance(it, dict):
                        parts += [plain(it.get("h")), plain(it.get("p")), plain(it.get("t"))]
                    else:
                        parts.append(plain(it))
            elif isinstance(val, str):
                parts.append(plain(val))
        return " ".join(x for x in parts if x)

    body = [p for p in pages if p.get("type") not in ("cover", "toc", "thanks")]

    # 1) 同页同构/同词缀短标签（label_parallel_streak / label_affix_repeat）
    stop_char = set("的了在与和及或上个")
    for i, pg in enumerate(pages, 1):
        labels = []
        for key in ("items", "bullets", "blocks", "steps"):
            for it in pg.get(key) or []:
                if isinstance(it, dict):
                    v = plain(it.get("k") or it.get("h"))
                    if v and len(v) <= 8:
                        labels.append(v)
        if len(labels) >= 4:
            from collections import Counter
            cnt = Counter(shape(x) for x in labels)
            top, num = cnt.most_common(1)[0]
            if num >= 4:
                hits.append(Hit(file=file_label, field=_page_label(i - 1, pg), context="deck",
                                rule_id="label_parallel_streak", rule_name="同页同构短标签 ≥4",
                                severity="medium", matched=f"「{top}」×{num}",
                                sentence="、".join(labels[:6]),
                                suggestion="换掉一半标签的写法：有的用名词，有的用「动词+对象」，有的带数字或单位。",
                                hint="同页 4 个以上标签是同一句式（如全是「X的Y」或全是两字词），像模板批量填的。"))
        if len(labels) >= 3:
            from collections import Counter
            # 编号式标签（贡献一/步骤2/方案③）共享词干是正常写法，先排除
            enum_re = _re.compile(r"[一二三四五六七八九十\d①②③④⑤⑥⑦⑧⑨⑩]$")
            plain_labels = [v for v in labels if not enum_re.search(v)]
            # ≥3 个标签共用同一个含汉字的双字词缀：「建模/结构/回放 可复用」；
            # 纯 ASCII 前缀（MODE/MODE-w/MODE-o）是方法名，不算
            big_cnt = Counter()
            for v in plain_labels:
                cjk_bgs = {v[j:j + 2] for j in range(len(v) - 1)
                           if any("\u4e00" <= c <= "\u9fff" for c in v[j:j + 2])}
                for bg in cjk_bgs:
                    big_cnt[bg] += 1
            for bg, num in big_cnt.most_common(1):
                if num >= 3:
                    hits.append(Hit(file=file_label, field=_page_label(i - 1, pg), context="deck",
                                    rule_id="label_affix_repeat", rule_name="同页标签共用词缀",
                                    severity="medium", matched=f"「{bg}」出现在 {num} 个标签里",
                                    sentence="、".join(labels[:6]),
                                    suggestion="共用词缀的标签换掉一半：改成不同结构，或把共性放进正文说明。",
                                    hint="「建模可复用/结构可复用/回放可复用」是同一模板连填三次。"))
                    break
            # ≥4 个标签围着同一个字打转：「起点更高/收敛更快/换b更稳/部署更省」全带「更」
            char_cnt = Counter()
            for v in plain_labels:
                for c in {c for c in v if "\u4e00" <= c <= "\u9fff" and c not in stop_char}:
                    char_cnt[c] += 1
            for c, num in char_cnt.most_common(1):
                if num >= 4:
                    hits.append(Hit(file=file_label, field=_page_label(i - 1, pg), context="deck",
                                    rule_id="label_affix_repeat", rule_name="同页标签共用一字",
                                    severity="medium", matched=f"「{c}」出现在 {num} 个标签里",
                                    sentence="、".join(labels[:6]),
                                    suggestion="这组标签是同一个句式套出来的，保留一处，其余改成具体判断。",
                                    hint="「起点更高/收敛更快/部署更省」全带「更」，是模板对仗。"))
                    break

    # 2) 计数词密度（count_word_density）：每页 visible 文本里量词 ≥3 次
    num_re = _re.compile(r"[一二三四五六七八九十两0-9]+\s*[个条句种步块处点类项篇次组张页所台套]")
    for i, pg in enumerate(pages, 1):
        heads = [plain(pg.get("title")), plain(pg.get("lead")), plain(pg.get("sub"))]
        for key in ("items", "bullets", "steps", "blocks"):
            for it in pg.get(key) or []:
                if isinstance(it, dict):
                    heads.append(plain(it.get("k") or it.get("h")))
        txt = " ".join(x for x in heads if x)
        n = len(num_re.findall(txt))
        if n >= 3:
            hits.append(Hit(file=file_label, field=_page_label(i - 1, pg), context="deck",
                            rule_id="count_word_density", rule_name="单页计数词 ≥3",
                            severity="medium", matched=f"{n} 个量词",
                            sentence=txt[:80],
                            suggestion="把「三个/五个/两种」改成直接说对象或数字；一页里计数词最多留 2 个。",
                            hint="「五个对比方案，各去掉一块」这类句式的指纹就是量词密度高。"))

    # 3) 过渡页导语同构（section_lead_template）
    chapter_subs = [(i, plain(pg.get("sub"))) for i, pg in enumerate(pages, 1)
                    if pg.get("type") in ("chapter", "section") and plain(pg.get("sub"))]
    if len(chapter_subs) >= 2:
        shp = [shape(x) for _, x in chapter_subs]
        from collections import Counter
        dup = [k for k, v in Counter(shp).items() if v >= 2]
        if dup:
            hits.append(Hit(file=file_label, field="章节过渡页", context="deck",
                            rule_id="section_lead_template", rule_name="过渡页导语同构",
                            severity="medium", matched=f"{len(dup)} 组同句式",
                            sentence="；".join(x for _, x in chapter_subs[:3])[:120],
                            suggestion="过渡页导语换写法：有的用一句话，有的只放名词短语，有的干脆不要副题。",
                            hint="每章导语都是「A → B → C」或「从 A 到 B」，四章连看就是模板。"))

    # 4) 同一句出现在 ≥2 页（motif_repeat）。注意逐字段分句：
    #    visible() 用空格拼接，直接整页 split 会把全页粘成一个巨型子句，规则永远不触发。
    def _parts_of(pg):
        out = [plain(pg.get("title")), plain(pg.get("lead")), plain(pg.get("sub"))]
        for key in ("items", "bullets"):
            for it in pg.get(key) or []:
                if isinstance(it, dict):
                    out += [plain(it.get("k")), plain(it.get("t") or it.get("p"))]
                else:
                    out.append(plain(it))
        for key in ("lines", "blocks", "steps", "takeaway", "note", "tnote"):
            v = pg.get(key)
            if isinstance(v, list):
                for it in v:
                    if isinstance(it, dict):
                        out += [plain(it.get("h")), plain(it.get("p") or it.get("t"))]
                    else:
                        out.append(plain(it))
            elif isinstance(v, str):
                out.append(plain(v))
        return out

    clause_pages: dict[str, set] = {}
    gram_pages: dict[str, set] = {}
    for i, pg in enumerate(body, 1):
        for part in _parts_of(pg):
            for clause in _re.split(r"[，。；、：\n]", part):
                clause = _re.sub(r"\s+", "", clause)
                if len(clause) >= 8:
                    clause_pages.setdefault(clause, set()).add(i)
                for j in range(len(clause) - 5):
                    gram_pages.setdefault(clause[j:j + 6], set()).add(i)
    repeats = [(c, ps) for c, ps in clause_pages.items() if len(ps) >= 2]
    if repeats:
        repeats.sort(key=lambda x: -len(x[1]))
        hits.append(Hit(file=file_label, field="全文统计", context="deck",
                        rule_id="motif_repeat", rule_name="同一句出现在多页",
                        severity="medium", matched=f"{len(repeats)} 处重复",
                        sentence="；".join(c for c, _ in repeats[:3])[:120],
                        suggestion="同一句只留一处；其余页换说法或引用回前一页，不要复读。",
                        hint="人写多页会有呼应，但不会整句照搬；整句重复是模板复制粘贴的痕迹。"))
    # 近似复读：换几个词再说一遍（「一条策略覆盖不同偏好」vs「一条策略覆盖多种偏好」）
    near = [(g, ps) for g, ps in gram_pages.items() if len(ps) >= 2
            and not any(g in c for c, _ in repeats)]
    if near:
        near.sort(key=lambda x: (-len(x[1]), -len(x[0])))
        seen, shown = set(), []
        for g, _ps in near:
            if not any(g in s for s in seen):
                seen.add(g)
                shown.append(g)
            if len(shown) >= 3:
                break
        hits.append(Hit(file=file_label, field="全文统计", context="deck",
                        rule_id="motif_repeat", rule_name="金句近似复读（换词再说一遍）",
                        severity="low", matched=f"{len(near)} 个共享片段",
                        sentence="；".join(shown)[:120],
                        suggestion="同一个说法全篇只展开一次；其他页要么引用前一页，要么换更具体的说法。",
                        hint="封面副题的句子又出现在正文页和小结里，是「呼应」用力过猛。"))

    # 5) 同后缀标题 ≥3（title_suffix_repeat）
    titles = [plain(p.get("title")) for p in body if plain(p.get("title"))]
    from collections import Counter
    suffixes = Counter()
    for t in titles:
        m = _re.search(r"(的[\u4e00-\u9fff]{1,3})$", t)
        if m:
            suffixes[m.group(1)] += 1
    bad = [(k, v) for k, v in suffixes.items() if v >= 3]
    if bad:
        hits.append(Hit(file=file_label, field="全文统计", context="deck",
                        rule_id="title_suffix_repeat", rule_name="同后缀标题 ≥3",
                        severity="medium", matched="、".join(f"{k}×{v}" for k, v in bad),
                        sentence="；".join(t for t in titles if any(t.endswith(k) for k, _ in bad))[:120],
                        suggestion="换掉一部分后缀：改成「扫描」「敏感性」「对比」等名词，或直接去掉「的影响」。",
                        hint="连续三页「……的影响」是模板批量命名。"))

    # 6) 导语与要点名词重叠 ≥3（lead_previews_bullets）
    for i, pg in enumerate(pages, 1):
        lead = plain(pg.get("lead"))
        if not lead:
            continue
        item_txt = " ".join(
            plain(it.get("k")) + plain(it.get("t") or it.get("p")) if isinstance(it, dict) else plain(it)
            for it in (pg.get("items") or pg.get("bullets") or []))
        if len(item_txt) < 10:
            continue
        STOP = {"我们", "可以", "这个", "一个", "的是", "不是", "就是", "以及", "同时", "因此"}
        lb = {lead[j:j + 2] for j in range(len(lead) - 1)} - STOP
        ib = {item_txt[j:j + 2] for j in range(len(item_txt) - 1)}
        overlap = lb & ib
        if len(overlap) >= 5:
            hits.append(Hit(file=file_label, field=_page_label(i - 1, pg), context="deck",
                            rule_id="lead_previews_bullets", rule_name="导语在重复要点",
                            severity="medium", matched=f"重叠 {len(overlap)} 个双字词",
                            sentence=lead[:80],
                            suggestion="导语不要预告下面的要点，改成给结论、给数字或干脆删掉。",
                            hint="导语把要点名词又说一遍，等于同页复读。"))

    # 7) 收束句页数 ≥4（summary_prefix）：build_deck 会把每条 takeaway 渲染成底部收束行，
    #    所以按 takeaway 非空的页数算，而不是在 content 里找「小结」前缀。
    taken = [(i, plain(pg.get("takeaway"))) for i, pg in enumerate(body, 1) if plain(pg.get("takeaway"))]
    if len(taken) >= 4:
        hits.append(Hit(file=file_label, field="全文统计", context="deck",
                        rule_id="summary_prefix", rule_name="收束句（takeaway）≥4 页",
                        severity="medium", matched=f"{len(taken)} 页带收束句",
                        sentence="；".join(t for _, t in taken[:3])[:120],
                        suggestion="收束句只留给真正要收束的 2–3 页；其余页要点里已经说清，直接删。",
                        hint="每页都在底部收一句是模板动作；人只在关键页收。"))

    # 8) 糊墙词尾巴（vague_tail）
    VAGUE = ("进一步", "有效地", "显著地", "更好地", "更加", "充分", "深入", "全面", "合理", "良好", "大量")
    for i, pg in enumerate(pages, 1):
        txt = visible(pg)
        for v in VAGUE:
            for m in _re.finditer(v + r"(?:[，。；]|$)", txt):
                hits.append(Hit(file=file_label, field=_page_label(i - 1, pg), context="deck",
                                rule_id="vague_tail", rule_name="糊墙词收尾",
                                severity="low", matched=v,
                                sentence=txt[max(0, m.start() - 20):m.end() + 2],
                                suggestion="换成可核查的量或动作：「进一步降低」→「低 12%」；删掉空形容词。",
                                hint="进一步/有效/充分/深入 这类词放在句尾，等于什么都没说。"))

    # 9) 页内要点相似（near_duplicate_bullet）
    for i, pg in enumerate(body, 1):
        items = []
        for it in pg.get("items") or pg.get("bullets") or []:
            items.append(plain(it.get("k")) + " " + plain(it.get("t") or it.get("p"))
                         if isinstance(it, dict) else plain(it))
        if len(items) < 3:
            continue
        bg = [{s[j:j + 2] for j in range(len(s) - 1)} for s in items]
        worst = 0.0
        pair = ""
        for a in range(len(items)):
            for b in range(a + 1, len(items)):
                inter = len(bg[a] & bg[b])
                union = len(bg[a] | bg[b]) or 1
                jac = inter / union
                if jac > worst:
                    worst, pair = jac, f"{items[a][:18]} / {items[b][:18]}"
        if worst >= 0.55:
            hits.append(Hit(file=file_label, field=_page_label(i - 1, pg), context="deck",
                            rule_id="near_duplicate_bullet", rule_name="页内要点相似",
                            severity="medium", matched=f"相似度 {worst:.2f}",
                            sentence=pair,
                            suggestion="合并或删掉一条；两条要点说同一件事时，留信息量大的那条。",
                            hint="页内两条要点字面或语义高度重合，是「为了凑满四条」的典型痕迹。"))


def extract_content_units(data: dict) -> list[tuple[str, str, str]]:
    """把 content.json 摊平成 [(字段路径, kind, 文本)]。"""
    units: list[tuple[str, str, str]] = []

    def add(field_path: str, kind: str, value) -> None:
        if isinstance(value, str) and value.strip():
            units.append((field_path, kind, value))

    add("title", "deck", data.get("title"))
    add("subtitle", "deck", data.get("subtitle"))
    add("citation", "deck", data.get("citation"))

    paper = data.get("paper")
    if isinstance(paper, dict):
        add("paper.summary", "deck", paper.get("summary"))
        add("paper.title_cn", "deck", paper.get("title_cn"))
        for i, item in enumerate(paper.get("contributions") or []):
            if isinstance(item, dict):
                add(f"paper.contributions[{i}].k", "deck", item.get("k"))
                add(f"paper.contributions[{i}].t", "deck", item.get("t"))
        for i, item in enumerate(paper.get("limits") or []):
            if isinstance(item, dict):
                add(f"paper.limits[{i}].k", "deck", item.get("k"))
                add(f"paper.limits[{i}].t", "deck", item.get("t"))
        for i, fig in enumerate(paper.get("figures") or []):
            if isinstance(fig, dict):
                add(f"paper.figures[{i}].caption", "deck", fig.get("caption"))

    for si, sec in enumerate(data.get("sections") or []):
        if isinstance(sec, dict):
            add(f"sections[{si}].title", "deck", sec.get("title"))
            add(f"sections[{si}].summary", "deck", sec.get("summary"))

    for pi, page in enumerate(data.get("pages") or []):
        if not isinstance(page, dict):
            continue
        label = _page_label(pi, page)
        for key in ("title", "lead", "notes", "tag", "tnote", "note", "caption", "caption2", "sub", "text", "from"):
            add(f"{label}.{key}", "deck", page.get(key))
        for key in ("items", "blocks", "bullets", "steps", "lines", "figs", "formulas"):
            for ii, item in enumerate(page.get(key) or []):
                if isinstance(item, str):
                    add(f"{label}.{key}[{ii}]", "deck", item)
                elif isinstance(item, dict):
                    for sub in ("k", "t", "h", "p", "d", "t2", "caption", "text", "no"):
                        add(f"{label}.{key}[{ii}].{sub}", "deck", item.get(sub))
        for ri, row in enumerate(page.get("rows") or []):
            if isinstance(row, list):
                for ci, cell in enumerate(row):
                    add(f"{label}.rows[{ri}][{ci}]", "deck", cell)
            elif isinstance(row, str):
                add(f"{label}.rows[{ri}]", "deck", row)
        cols = page.get("cols")
        if isinstance(cols, list):
            for ci, col in enumerate(cols):
                if isinstance(col, dict):
                    add(f"{label}.cols[{ci}].h", "deck", col.get("h"))
                    for ii, item in enumerate(col.get("items") or []):
                        add(f"{label}.cols[{ci}].items[{ii}]", "deck", item)
        side = page.get("side")
        if isinstance(side, dict):
            add(f"{label}.side.h", "deck", side.get("h"))
            for ii, item in enumerate(side.get("items") or []):
                if isinstance(item, dict):
                    add(f"{label}.side.items[{ii}].k", "deck", item.get("k"))
                    add(f"{label}.side.items[{ii}].t", "deck", item.get("t"))
                    add(f"{label}.side.items[{ii}].h", "deck", item.get("h"))
                    add(f"{label}.side.items[{ii}].d", "deck", item.get("d"))
        # 页面级自定义讲稿
        speech = page.get("speech")
        if isinstance(speech, dict):
            for key in ("text", "open", "bg", "skip", "transition"):
                add(f"{label}.speech.{key}", "speech", speech.get(key))
        elif isinstance(speech, str):
            add(f"{label}.speech", "speech", speech)

    speech = data.get("speech")
    if isinstance(speech, dict):
        add("speech.greeting", "speech", speech.get("greeting"))
        add("speech.opening", "speech", speech.get("opening"))
        auto_pages = speech.get("auto_pages")
        if isinstance(auto_pages, dict):
            for key, val in auto_pages.items():
                add(f"speech.auto_pages.{key}", "speech", val)
        for i, item in enumerate(speech.get("background") or []):
            if isinstance(item, dict):
                add(f"speech.background[{i}].h", "deck", item.get("h"))
                add(f"speech.background[{i}].p", "speech", item.get("p"))
        for i, item in enumerate(speech.get("terms") or []):
            if isinstance(item, dict):
                add(f"speech.terms[{i}].cn", "deck", item.get("cn"))
                add(f"speech.terms[{i}].en", "deck", item.get("en"))
                add(f"speech.terms[{i}].plain", "speech", item.get("plain"))
        for i, item in enumerate(speech.get("qa") or []):
            if isinstance(item, dict):
                add(f"speech.qa[{i}].q", "speech", item.get("q"))
                add(f"speech.qa[{i}].a", "speech", item.get("a"))
        for i, item in enumerate(speech.get("tips") or []):
            if isinstance(item, dict):
                add(f"speech.tips[{i}].h", "deck", item.get("h"))
                add(f"speech.tips[{i}].p", "speech", item.get("p"))

    return units


def extract_text_units(text: str) -> list[tuple[str, str, str]]:
    """md/txt/docx：每行一个单元，行号 + 最近的标题。"""
    units: list[tuple[str, str, str]] = []
    heading = ""
    for idx, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()
        if not line:
            continue
        if line.startswith("#"):
            heading = line.lstrip("#").strip()
            units.append((f"第{idx}行（标题）", "deck", heading))
            continue
        label = f"第{idx}行" + (f"（{heading[:14]}）" if heading else "")
        # Markdown 表格行 / 列表行不算口语正文，避免误判句长与重复句首
        kind = "deck" if line.startswith("|") else "speech"
        units.append((label, kind, line))
    return units


# ---------------------------------------------------------------------------
# 内容契约长度检查（low，属于版式约束，不算 AI 腔）
# ---------------------------------------------------------------------------


def check_contract_limits(units: Sequence[tuple[str, str, str]], hits: list[Hit]) -> None:
    limits = {"title": 22, "lead": 48, "notes": 140}
    for field_path, kind, text in units:
        if kind != "deck":
            continue
        tail = field_path.rsplit(".", 1)[-1]
        limit = None
        if tail in limits:
            limit = limits[tail]
        elif re.search(r"\.items\[\d+\]\.t$", field_path):
            limit = 62
        elif re.search(r"\.items\[\d+\]\.k$", field_path):
            limit = 10
        if not limit:
            continue
        n = len(text.strip())
        if n > limit:
            hits.append(
                Hit(
                    file="",
                    field=field_path,
                    context="deck",
                    rule_id="over_contract",
                    rule_name="超出 format_spec 文案上限",
                    severity="low",
                    matched=f"{n} 字 > {limit} 字",
                    sentence=text[:80],
                    suggestion=f"压到 {limit} 字以内；超了生成器会截断或挤版。",
                    hint="上限见 references/format_spec.md「文案纪律」表。",
                )
            )


# ---------------------------------------------------------------------------
# 报告输出
# ---------------------------------------------------------------------------


def sort_hits(hits: list[Hit]) -> list[Hit]:
    return sorted(hits, key=lambda h: (SEVERITY_ORDER.get(h.severity, 9), h.file, h.field, h.offset))


def summarize(hits: list[Hit]) -> dict:
    summary = {"high": 0, "medium": 0, "low": 0, "total": len(hits)}
    for hit in hits:
        summary[hit.severity] = summary.get(hit.severity, 0) + 1
    return summary


def print_text_report(
    target: Path,
    kind_label: str,
    hits: list[Hit],
    stats: dict | None,
    hints: bool,
    json_stdout: bool,
) -> None:
    summary = summarize(hits)
    if json_stdout:
        return
    print(f"humanize_lint v{VERSION} · {kind_label}：{target}")
    if stats:
        bits = "，".join(f"{k} {v}" for k, v in stats.items() if v not in (None, ""))
        if bits:
            print(f"对象概况：{bits}")
    print(
        f"命中 {summary['total']} 处："
        f"high {summary['high']} / medium {summary['medium']} / low {summary['low']}"
    )
    if not hits:
        print("没有命中「说人话」规则。仍建议朗读一遍，检查是否像念稿。")
        return

    current_file = None
    current_sev = None
    for hit in hits:
        if hit.file != current_file or hit.severity != current_sev:
            label = SEVERITY_LABEL.get(hit.severity, hit.severity)
            print(f"\n── {label} ──")
            current_file = hit.file
            current_sev = hit.severity
        print(f"[{hit.severity}] {hit.field}  「{hit.rule_name}」")
        print(f"    命中：{hit.matched.strip()[:70]}")
        if hit.sentence:
            print(f"    原句：{hit.sentence.strip()[:90]}")
        print(f"    建议：{hit.suggestion}")
        if hints and hit.hint:
            for line in hit.hint.splitlines():
                print(f"      {line}")

    # 按规则统计
    counts: dict[str, int] = {}
    for hit in hits:
        counts[hit.rule_id] = counts.get(hit.rule_id, 0) + 1
    top = sorted(counts.items(), key=lambda kv: -kv[1])
    print("\n规则命中：" + "，".join(f"{RULE_BY_ID[k].name if k in RULE_BY_ID else k}×{v}" for k, v in top))
    if summary["high"]:
        print("提示：high 命中先改；medium 看语境；low 是风格建议（不影响 strict）。")
    if not hints:
        print("加 --hints 看每条的具体改写示范。")


def build_json_report(target: Path, kind: str, hits: list[Hit], stats: dict | None) -> dict:
    return {
        "tool": "humanize_lint",
        "version": VERSION,
        "target": str(target),
        "kind": kind,
        "summary": summarize(hits),
        "stats": stats or {},
        "hits": [hit.to_dict() for hit in sort_hits(hits)],
    }


def prose_stats(units: Sequence[tuple[str, str, str]]) -> dict:
    texts = [text for _, kind, text in units if kind in ("speech", "text") and text]
    whole = "\n".join(texts)
    sentences = split_sentences(whole)
    lengths = [han_count(s) for s in sentences]
    if not lengths:
        return {"字数": han_count(whole)}
    short = sum(1 for x in lengths if x <= 20)
    return {
        "字数": han_count(whole),
        "句数": len(lengths),
        "平均句长": f"{statistics.mean(lengths):.1f}",
        "句长标准差": f"{statistics.pstdev(lengths):.1f}",
        "短句占比": f"{short / len(lengths):.0%}",
        "60字以上": sum(1 for x in lengths if x > 60),
        "80字以上": sum(1 for x in lengths if x > 80),
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="检查 PPT 文案 / 讲稿的 AI 腔（humanizer 中文规则）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--content", metavar="content.json", help="content.json 路径（format_spec v1/v2）")
    source.add_argument("--text", metavar="文件", help="讲稿 .md / .txt / .docx 路径")
    parser.add_argument("--json", nargs="?", const="-", metavar="report.json",
                        help="输出 JSON 报告；不带路径时打到标准输出")
    parser.add_argument("--strict", action="store_true", help="有 high 命中时退出码 1")
    parser.add_argument("--hints", action="store_true", help="打印每条的改写建议与示例")
    parser.add_argument("--quiet", action="store_true", help="只输出汇总与退出码")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)

    if args.content:
        path = Path(args.content).expanduser()
        if not path.exists():
            print(f"[错误] 找不到文件：{path}", file=sys.stderr)
            return 2
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            print(f"[错误] JSON 解析失败：{exc}", file=sys.stderr)
            return 2
        if not isinstance(data, dict):
            print("[错误] content.json 顶层必须是对象", file=sys.stderr)
            return 2
        target = path
        kind = "content"
        kind_label = "content.json"
        units = extract_content_units(data)
        stats = prose_stats(units)
        if not stats.get("句数"):
            # v1 内容没有页面级讲稿，退而统计全部文案字数
            stats = {"文案字数": sum(len(text) for _, _, text in units)}
        stats["页数"] = len(data.get("pages") or [])
        if isinstance(data.get("paper"), dict) and data["paper"].get("title_cn"):
            stats["论文"] = str(data["paper"]["title_cn"])[:16]
    else:
        path = Path(args.text).expanduser()
        if not path.exists():
            print(f"[错误] 找不到文件：{path}", file=sys.stderr)
            return 2
        raw = read_text_target(path)
        target = path
        kind = "text"
        kind_label = "讲稿" if path.suffix.lower() in (".md", ".txt", ".docx") else "文本"
        units = extract_text_units(raw)
        stats = prose_stats(units)

    hits: list[Hit] = []
    check_unit(units, target.name, hits)
    check_prose_rules(units, target.name, hits)
    if args.content:
        # 整篇节奏检查：数字对仗、导语覆盖率、标题/导语句式
        check_deck_rhythm(data, target.name, hits)
        # 结构性指纹：同构标签、量词密度、复读句、标题后缀等
        check_deck_structure(data, target.name, hits)
        # 占位符和契约长度检查只在 content 上有意义
        check_contract_limits(units, hits)
        for hit in hits:
            if hit.file == "":
                hit.file = target.name

    hits = sort_hits(hits)
    report = build_json_report(target, kind, hits, stats)

    if args.json:
        payload = json.dumps(report, ensure_ascii=False, indent=2)
        if args.json == "-":
            print(payload)
        else:
            out = Path(args.json).expanduser()
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(payload + "\n", encoding="utf-8")
            print(f"JSON 报告已写入：{out}", file=sys.stderr)

    if not args.quiet and args.json != "-":
        print_text_report(target, kind_label, hits, stats, args.hints, json_stdout=False)
    elif args.quiet and args.json != "-":
        summary = summarize(hits)
        print(f"{target.name}: high {summary['high']} / medium {summary['medium']} / low {summary['low']}")

    if args.strict and report["summary"]["high"] > 0:
        print(f"[strict] 存在 {report['summary']['high']} 处 high 命中，返回 1", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

# format_spec v2 — 内容 JSON 与页型

生成器读入 `content.json` + 可选 `profile.json`，输出一个 PPTD 工程（再导出 pptx）。
本文件是唯一的内容契约：`verify_claims.py`、`build_speech.py`、`humanize_lint.py`、`build_deck.py` 共用。

## 顶层结构

```json
{
  "schema_version": 2,
  "profile": { "...": "见 profile_schema.md，可省略，由 --profile 提供" },
  "paper": { "...": "文献汇报必填，课程汇报可省" },
  "sections": [{"num": "1", "title": "研究背景", "summary": "这一节回答什么"}],
  "pages": [{"type": "...", "...": "..."}],
  "meta": {
    "source_pdf": "论文 PDF 路径（可选）",
    "extracted_from": "论文标题或材料名（可选）",
    "humanized": true,
    "verified": true
  }
}
```

- `pages` 只需写正文页：封面、论文简介、目录由生成器按 `profile` + `paper` + `sections` 自动补。
- 想手写封面/目录，在 `pages` 里加 `{"type":"cover"}` / `{"type":"toc"}` 覆盖自动生成。
- 页面顺序 = `pages` 顺序；页码自动算。
- 所有页面都可带 `notes`（讲稿要点，写给人念的短句，不是 Markdown 长文），生成 pptx 的演讲者备注，同时作为 `build_speech.py` 的底稿。

## 文案纪律（Generator 会校验长度）

| 位置 | 上限 | 要求 |
|---|---|---|
| 页面 `title` | 22 字 | 一句结论，读者只看标题就知道这页在说什么；不要「研究方法」这种空标题 |
| `eyebrow` | 16 字 | 编号式小标题，如 `2.1 系统模型`；`title` 已是结论时用它保留编号 |
| `lead` | 48 字 | 这页的一句话结论（高亮条），可省 |
| `items[].k` | 12 字 | 要点核心词，加粗显示 |
| `items[].t` | 62 字（约两行） | 一句话说明；分条列点，不写长段落 |
| `notes` | 140 字 | 口头讲稿，1–3 句 |
| `table` 单元格 | 24 字 | 超了换行，不要塞长句 |
| `captions` | 40 字 | 图注只写图是什么，不下结论 |

语言要求见 `references/humanizer_cn.md`：说人话、短句、具体数字、不堆排比、不用破折号串句。

## paper 字段（文献汇报）

```json
"paper": {
  "title_cn": "中文题目",
  "title_en": "English title",
  "authors": "作者1、作者2、…",
  "affiliations": "单位1；单位2",
  "keywords": "关键词1；关键词2",
  "summary": "一句话说清这篇论文做了什么（≤48 字）",
  "venue": "IEEE Transactions on Mobile Computing",
  "volume": "2026 年 10 月 · 第 25 卷第 10 期 · 第 17472–17488 页",
  "doi": "10.1109/TMC.2026.3693366",
  "ranks": ["CCF A 类（计算机网络）", "中科院 1 区 Top"],
  "metrics": "影响因子 8.8（2025 JCR）",
  "metrics_note": "CiteScore 11.90，H-index 122",
  "funding": "国家自然科学基金 U25A20398 等",
  "source_note": "分区数据来源：LetPub / 期刊官网（查询日期 2026-11）",
  "contributions": [{"k": "多目标 MDP", "t": "把双目标权衡写进同一个奖励向量"}],
  "limits": [{"k": "单基站", "t": "没有验证多基站协作的场景"}],
  "figures": [
    {"file": "fig01_scene.png", "caption": "低空多目标 ISAC 场景", "page": 2, "kind": "scene"},
    {"file": "fig02_framework.png", "caption": "MODE 框架", "page": 4, "kind": "framework"}
  ]
}
```

`kind` 取值：`scene` 场景图 / `framework` 框架图 / `result` 实验结果 / `table` 表格 / `concept` 概念图。生成器按 kind 推荐页型。

## 通用页面字段

| 字段 | 说明 |
|---|---|
| `type` | 页型，见下 |
| `eyebrow` | 编号式眉标，如 `1.1`；academic 皮肤显示在标题上方 |
| `title` | 页面主标题；结论式 |
| `lead` | 一句话结论，显示为浅色底/左侧色条的高亮条 |
| `notes` | 讲稿备注 |
| `motion` | 覆盖该页动画：`none` / `gentle` / `lively` |
| `fig` / `figs` / `fig_w` / `fig_h` / `caption` / `caption2` | 图片相关，见各页型 |
| `source` | 论文图来源脚注，只在该图首次出现时写 |

## 页型清单

### 1. `cover`（自动）
字段取自 profile + paper：校徽、校名、汇报题目、副标题、汇报人、学院、日期、页脚。无需手写；需要时加 `{"type":"cover","tagline":"…"}`。

### 2. `paper_card`（自动）
论文简介：题目（中英）、作者、单位、关键词、一句话 summary、期刊/卷期/DOI、分区与影响因子。`paper` 有值时自动插入在封面后。

### 3. `toc`（自动）
按 `sections` 与正文标题编号自动生成目录（章节名、描述、页码）。

### 4. `chapter`（章节过渡页）

```json
{"type": "chapter", "num": "2", "title": "这篇论文怎么做",
 "sub": "多任务 MoE + 整回合回放 + 在线微调",
 "bullets": ["MODE 整体框架", "多任务 MoE 结构", "混合经验回放"], "notes": "…"}
```

- `num` 大数字显示；`bullets` 最多 4 条，用于预告本节内容。
- academic/现代皮肤：白底大数字；campus/bold：品牌色底白字。

### 5. `bullets`（核心：分条列点）

```json
{"type": "bullets", "eyebrow": "1.2", "title": "三种主流方法，各缺一块",
 "lead": "传统优化、DQN、DDPG 都解决不了连续控制 + 多目标",
 "items": [
   {"k": "传统优化", "t": "非凸又强耦合，迭代代价高，常停在次优"},
   {"k": "DQN", "t": "只能选离散动作，连续量离散化后维度会爆"},
   {"k": "DDPG", "t": "能出连续动作，但一次只学一个目标"}
 ],
 "layout": "rows", "tag": "方法对比", "notes": "…"}
```

- `items` 3–5 条，`k` 加粗、`t` 正常；`layout`：`rows`（上下行）/ `cards`（横向卡片）/ `split`（左要点右图，配 `fig`）。
- 可选 `num` 字段（`"18%"`）显示为大字数字徽标；可选 `items[].tag`。
- academic 皮肤：行首品牌色小方块；campus：编号圆点；modern：细线条目；bold：左侧粗色条。

### 6. `keynums`（关键数字页）

```json
{"type": "keynums", "title": "这篇论文的四个关键数字",
 "items": [
   {"num": "8.8", "unit": "IF", "t": "TMC 2025 影响因子"},
   {"num": "40", "unit": "slot", "t": "一回合的时隙数"},
   {"num": "5", "unit": "万", "t": "训练回合数"}
 ]}
```

2–4 个数字；`num` ≤6 字符，`unit` ≤6 字符。

### 7. `text`（纯文字页）

```json
{"type": "text", "eyebrow": "1.1", "title": "低空经济需要通感一体",
 "blocks": [{"h": "小标题", "p": "正文一段，1–2 行。"}], "notes": "…"}
```

`h` 可省；最多 4 块，每块正文 ≤2 行。

### 8. `left_fig` / `fig_bullets` / `big_fig` / `two_figs`

```json
{"type": "left_fig", "eyebrow": "2.1", "title": "一个基站，同时供网和盯目标",
 "blocks": [{"h": "地面基站", "p": "N 根天线，通信感知共用波束"}],
 "fig": "fig01_scene.png", "fig_w": 444, "fig_h": 250,
 "caption": "低空多目标 ISAC 场景（论文原图）",
 "source": "X. Ye et al., IEEE Trans. Mobile Comput., 2026"}

{"type": "fig_bullets", "title": "…", "fig": "fig07.png", "fig_w": 400, "fig_h": 340,
 "caption": "…", "items": [{"k": "MODE", "t": "…"}]}

{"type": "big_fig", "title": "…", "fig": "fig02.png", "fig_w": 760, "fig_h": 284,
 "caption": "…", "lines": ["第一行说明。", "第二行说明。"]}

{"type": "two_figs", "title": "…",
 "figs": [{"file": "fig05.png", "caption": "图 2.6（a）和速率随回合变化"},
          {"file": "fig06.png", "caption": "图 2.6（b）感知 SNR 随回合变化"}],
 "fig_w": 372, "fig_h": 300, "lines": ["最多两行说明"]}
```

图片文件放 `--media-dir`，生成器会拷进工程 `media/` 并重新编号。

### 9. `table`（三线表）

```json
{"type": "table", "title": "实验设置",
 "columns": [0.34, 0.66], "tbl_w": 520, "tbl_h": 246,
 "rows": [["参数", "设置"], ["天线数 N", "8"]],
 "tnote": "论文没有写 Δt 与 P_max 的取值。",
 "side": {"h": "四个对比方案", "items": [{"k": "MODE-w", "t": "去掉多任务 MoE"}]}}
```

### 10. `formula`

```json
{"type": "formula", "title": "问题建模",
 "formulas": [{"t": "max<sub>{W<sub>c</sub>,W<sub>s</sub>,u}</sub> ( E[ΣR<sub>total</sub>] , E[ΣΓ<sub>tar</sub>] )", "no": "(6a)"}],
 "note": "公式下面的一句话说明。"}
```

### 11. `compare`（对比页）

```json
{"type": "compare", "title": "传统方法 vs 本文方法",
 "cols": [
   {"h": "传统优化", "items": ["非凸、迭代慢", "换偏好权重要重解"]},
   {"h": "本文 MODE", "items": ["一次训练覆盖多个偏好", "换 b 只微调一层"], "highlight": true}
 ]}
```

2–3 列；`highlight` 列用品牌色描边。

### 12. `timeline`（时间线）

```json
{"type": "timeline", "title": "研究计划",
 "steps": [{"h": "2026.11–12", "t": "完成场景建模"}, {"h": "2027.01–03", "t": "算法实现"}]}
```

3–5 步。

### 13. `pipeline`（方法流程）

```json
{"type": "pipeline", "title": "MODE 训练流程",
 "steps": [{"h": "约束动作选择", "t": "每时隙生成合法动作"}, {"h": "整回合回放", "t": "按完整回合取经验"}],
 "caption": "图 2.2 流程（自绘示意）"}
```

3–5 步，箭头连接。

### 14. `quote`（金句页）

```json
{"type": "quote", "text": "与其给每个偏好权重训练一个策略，不如让它们共享一套底层表示。", "from": "论文的核心直觉"}
```

大字号一句话，最多 2 行。

### 15. `summary`

```json
{"type": "summary", "title": "结论与不足",
 "cols": [{"h": "三条结论", "items": ["把约束做进动作里，波束和轨迹能一起优化", "MoE 让偏好权重可迁移"]},
          {"h": "不足与下一步", "items": ["只有单基站场景", "感知只用了 SNR"]}]}
```

### 16. `roadmap`

```json
{"type": "roadmap", "title": "下一步",
 "items": [{"when": "近期", "t": "补多基站实验"}, {"when": "中期", "t": "加入真实轨迹数据"}]}
```

### 17. `qa`（预设问答，答辩/组会用）

```json
{"type": "qa", "title": "可能的提问",
 "items": [{"q": "为什么不用 PPO？", "a": "动作空间连续且预算敏感，DDPG 系更省样本。"}]}
```

### 18. `thanks`

```json
{"type": "thanks", "title": "谢谢", "subtitle": "请老师和同学批评指正"}
```

## 自动生成与目录规则

- 封面 = 校徽 + 校名 + `title`（content 顶层没有 `title` 时用 paper.title_cn）+ `subtitle` + 汇报人/学院/日期 + 底部 motto。
- 论文简介只在 `paper.title_cn` 或 `paper.title_en` 存在时生成；页数 +1。
- 目录页只在 `sections` 非空且正文页 ≥6 时生成；章节描述取该节前三页标题。
- `title` 字段命名：封面标题取 `content.title`（可选），正文页各用自己的 `title`。
- 页码：按皮肤决定写法，生成器自动算。

## 内容精度要求（`verify_claims.py` 会查）

1. 每个数字、单位、专有名词必须能在论文原文找到；论文没写的参数标「论文没有写」，不替作者补。
2. 每条要点可回溯到原文位置（`meta.paper_evidence` 可选：`[{"page":9,"quote":"…"}]`）。
3. `paper.ranks` / `metrics` 必须现查（LetPub、X-MOL、期刊官网），并写 `source_note`；查不到就省略，不猜。
4. 图注编号用论文原图号（如「图 2.6（a）」），不要自己重编。

## 节奏要求（防「整齐得不像人」）

- **不要每页都配导语**：`lead` 覆盖率控制在 40% 以内；图页、表页、公式页通常不需要 lead；导语不要把本页要点的名词再列一遍。
- **数字对仗标题全篇最多 2–3 处**：「X 个 Y」「三句话」「各缺一块」这类句式不要页页复用；其余换成陈述、提问、条件或名词短语。
- **标题长短错开**：允许 5 字标题和 20 字标题并存，允许带冒号补一句；全篇标题长度标准差不要过小；同后缀标题（「……的影响」）全篇 ≤2 个。
- **语气有波动**：同一份 PPT 里既要有偏书面的技术描述（「受式 (6f) 约束」），也要有口语解释；同一个口语判词全篇最多出现一次。
- **计数词克制**：一页最多 2 个「三条/五个/两种」式量词；结构性计数（章节号、方案个数）除外。
- **标签长短错开**：同页 ≥4 个同句式短标签（全是四字、全带「更」、共用「可复用」词缀）会被判为模板；名词、动宾、问句混着写。
- **收束句只给重点页**：`takeaway` 全篇 ≤3 页；生成器不加前缀，文案里也不要写「小结：」。
- **金句只说一次**：同一句结论不要同时出现在封面、正文页和小结。
- **工作备注不上屏**：`请再核对 / 待确认 / 占位 / 引用前` 这类写手提示不能出现在页面文案里；生成器会自动剔除并转存到讲稿备注。

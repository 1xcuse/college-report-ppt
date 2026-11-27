---
name: college-report-ppt
description: 生成全国本科院校通用的中文汇报 PPT（组会文献/进展、个人课程汇报、开题/中期、毕业答辩、项目竞赛）。开始前用本地网页向导弹窗收集学校/姓名/身份/汇报目的并匹配学校模板；从论文或材料提炼分条列点内容，带页面转场与入场动画，自动生成详细演讲稿 .docx，并做「说人话」检查。Use when the user asks for 组会汇报、文献汇报、组会进展、课程汇报、开题报告、中期检查、毕业答辩、项目汇报，或要求套学校模板、把论文做成汇报、写演讲稿、PPT 说得让人能看懂。
metadata:
  license: MIT
  version: "2.1.1"
  bundle: "PPT 技能包 v2.0.0"
  displayName: "学院汇报 PPT v2（全国本科院校 · 论文提炼 · 演讲稿）"
  requires:
    python: ["pyyaml", "python-pptx", "Pillow", "python-docx"]
    pythonOptional: ["PyMuPDF"]
    dshPlugin: "dsh-ppt-maker（仅模式 A 需要；模式 B 不需要）"
---

# 学院汇报 PPT v2（全国本科院校通用 · 论文提炼 · 演讲稿）

把一篇论文、一份课程作业或一段研究进展，做成中文汇报：**全国本科院校模板匹配 + 分条列点 + 页面动画 + 详细演讲稿 + 说人话**。

适用：组会文献汇报、组会进展汇报、个人课程汇报、开题/中期、毕业答辩、项目竞赛汇报。
不适用：需要复杂图表编辑的正式出版级排版、纯英文对外展示。

## 先读：这个技能包含什么

| 能力 | 实现 | 入口 |
|---|---|---|
| 开始前弹窗收信息、匹配学校模板 | 本地网页向导（浏览器打开，离线） | `scripts/profile_wizard.py` |
| 全国本科院校模板库（配色/校徽/校训/页脚） | 双一流 147 所有品牌包与校徽，其余自动降级为文字标识 | `assets/schools/schools.json`、`brand_packs.json` |
| 论文结构化提取 + 抽图 | PDF → 全文/论文卡/图表清单 | `scripts/paper_extract.py`、`extract_figures.py` |
| 内容回查（不编数据） | 数字/术语/图注逐一回原文 | `scripts/verify_claims.py` |
| 分条列点排版 + 4 套皮肤 + 18 种页型 | 生成 PPTD 工程再导出 pptx | `scripts/build_deck.py` |
| 页面转场 + 元素入场动画 | PPTD `animations` / 本地兜底转场 | 生成器自动 / `scripts/add_motion.py` |
| 演讲稿 .docx（补充 PPT 没写的背景） | 逐页讲稿 + 术语表 + 预设问答 + 时间分配 | `scripts/build_speech.py` |
| 「说人话」检查与改写 | 中文 AI 腔词表 + 改写建议 | `scripts/humanize_lint.py`、`references/humanizer_cn.md` |
| 封面装饰图 / 自定义校徽取色 | 确定性装饰图、logo 取色、文字校徽 | `scripts/cover_art.py`、`brand_from_logo.py` |

契约文件（写内容 JSON 前必读）：
- `references/profile_schema.md` — profile.json（向导输出）
- `references/format_spec.md` — content.json v2（页型、字段、文案上限）
- `references/brand_schema.md` — 学校名单 / 品牌包规范
- `references/design_language.md` — 4 套皮肤与动画
- `references/paper_extraction.md` — 论文怎么提炼成页
- `references/speech_script_guide.md` — 讲稿写作
- `references/humanizer_cn.md` — 中文说人话规则

## 安装位置与解释器（开工前先跑）

```bash
# SKILL_DIR：优先读 <skill_content> 里的 Base directory；否则按序探测
for d in "${SKILL_DIR:-}" "$PWD/.dsh/skills/college-report-ppt" "$HOME/.dsh/skills/college-report-ppt" \
         "$PWD/.agents/skills/college-report-ppt" "$HOME/.agents/skills/college-report-ppt"; do
  [ -n "$d" ] && [ -f "$d/SKILL.md" ] && SKILL_DIR="$d" && break
done

# PY：优先 DSH 自带 Python（内置 yaml/pptx/docx/PIL），否则系统 python3
PY="${DSH_PYTHON:-}"
if [ -z "$PY" ]; then
  for c in "${DSH_HOME:-$HOME/.dsh}"/dsh-runtimes/*/dependencies/python/bin/python3; do
    [ -x "$c" ] && PY="$c" && break
  done
fi
[ -n "$PY" ] || PY="$(command -v python3)"

"$PY" -c "import yaml, pptx, PIL, docx" 2>/dev/null \
  || echo "缺依赖：$PY -m pip install -r \"$SKILL_DIR/requirements.txt\""
# 论文提取需要 PyMuPDF（fitz）；没有时：优先换装了 fitz 的 python3，或
"$PY" -c "import fitz" 2>/dev/null || echo "PDF 提取：改用系统 python3（若已装 PyMuPDF）或 $PY -m pip install PyMuPDF"
```

工作目录约定：所有中间产物放用户工作区的 `work/<任务名>/`（内容 JSON、论文卡、图、讲稿），**不要写进技能目录**。

## 两种运行模式

| | 模式 A · DSH 工具链（推荐） | 模式 B · 本地 python-pptx |
|---|---|---|
| 前提 | 本机有 `dsh-ppt-maker` 插件（`ppt_create/ppt_export/ppt_render`） | 只要 Python |
| 导出 | `ppt_export`：字体子集嵌入、演讲者备注、页面转场、元素动画都保留 | `pptd_to_pptx.py`：备注/图表/元素动画会丢，转场用 `add_motion.py` 补 |
| 何时用 | 默认 | 对方机器没有 DSH 插件时；或要一份不依赖插件的成品 |

## 标准工作流（按顺序做，别跳步）

### 第 0 步 · 弹窗收档案（profile.json）

**先问，不要猜**。能开浏览器就跑网页向导（这就是"弹窗"）：

```bash
"$PY" "$SKILL_DIR/scripts/profile_wizard.py" --skill-dir "$SKILL_DIR" --out work/<任务名>/profile.json
```

- 页面里选：学校（可搜中文/拼音/首字母）、学院、姓名、身份、汇报目的、时长、日期、皮肤、「说人话」开关，可上传校徽/院徽；提交后自动写 `profile.json`。
- 浏览器打不开：`--plain` 把问题打到 stdout，按 `references/profile_schema.md` 手写 profile.json。
- 完全没有向导时，用 `ask_user_question` 逐项问（一次一张卡）：
  1. 学校（给 6 个候选+可自填，选项来自 `build_deck.py --list-schools`）
  2. 学院 / 姓名
  3. 身份：本科生 / 研究生 / 博士生 / 教师
  4. 目的：个人课程汇报 / 组会文献汇报 / 组会进展汇报 / 开题中期 / 毕业答辩 / 项目竞赛
  5. 时长（分钟）与页数偏好
  6. 皮肤与动画：学术稳重 / 校园课程 / 现代简约 / 撞色项目；不动画 / 轻动画 / 活跃
  7. 有没有校徽/院徽文件（有就让用户给路径；没有就用模板库或文字标识）
- 说明隐私：所有信息只写本地文件，不上传。

### 第 1 步 · 读材料（论文类必做）

```bash
"$PY" "$SKILL_DIR/scripts/paper_extract.py" --pdf <论文.pdf> --out work/<任务名>/paper
"$PY" "$SKILL_DIR/scripts/extract_figures.py" --pdf <论文.pdf> --out work/<任务名>/figures --max-figures 8
```

得到 `paper.md`（带页码标记）、`paper_card.json`、`figures.json`、`figures/*.png`。
先 `read` 一遍 `paper.md` 和 `paper_card.json`；论文图用 `read_image` 逐张确认（有没有截断、白边、图注错位）。
课程汇报/进展汇报没有 PDF 时，让用户给材料（文档、截图、实验记录），先整理成材料清单再进第 2 步。

### 第 2 步 · 提炼内容，写 content.json

按 `references/format_spec.md` 写；**先定页面大纲再写字**。提炼方法见 `references/paper_extraction.md`：
三遍法（结构 → 骨架 → 回查）、标题结论化（≤22 字）、每页 3–4 条要点（k≤10 字 / t≤46 字）、
每节 1 个 keynums 或 quote、图选 2–4 张核心图、公式只留 1–2 个。

结构按目的：

| purpose | 结构 | 默认页数 |
|---|---|---|
| `group_paper` 组会文献 | 论文卡 → 背景 → 方法（模型/框架/流程）→ 实验（设置/结果/对比）→ 结论与不足 | 15–20 |
| `group_progress` 组会进展 | 上次结论 → 本期工作 → 数据/结果 → 卡点 → 下一步 | 8–12 |
| `course_report` 课程汇报 | 选题背景 → 做了什么 → 结果/作品 → 收获与反思 | 10–14 |
| `proposal` 开题/中期 | 背景与意义 → 现状与不足 → 研究内容 → 技术路线 → 计划与基础 | 12–18 |
| `defense` 答辩 | 背景 → 方法 → 实验 → 结论 → 创新点 → 展望 | 18–25 |
| `project` 项目/竞赛 | 痛点 → 方案 → 亮点 → 数据 → 团队与计划 | 12–18 |

写完先自查：标题是不是结论、要点是不是 3–4 条、有没有长段落、数字带没带条件、图注只描述不下结论。

### 第 3 步 · 回查（文献汇报必跑）

```bash
"$PY" "$SKILL_DIR/scripts/verify_claims.py" \
  --content work/<任务名>/content.json --paper work/<任务名>/paper/paper.md \
  --figures work/<任务名>/figures/figures.json --json work/<任务名>/verify.json --strict
```

- 每个数字/单位/术语都能在原文找到；论文没给的参数写「论文没有写」，不替作者补。
- 期刊分区/影响因子必须现查（LetPub、X-MOL、官网）并写 `paper.source_note`，查不到就删掉，不猜。
- `--strict` 有 high issue 时退出码 1：逐条改到通过，再往下走。

### 第 4 步 · 说人话检查

```bash
"$PY" "$SKILL_DIR/scripts/humanize_lint.py" --content work/<任务名>/content.json --hints
```

按 `references/humanizer_cn.md` 改：短句、动词开头、具体数字；删掉「不仅…而且」「值得注意的是」「综上所述」「赋能/闭环/抓手」、破折号串句、排比三连、对称小标题。保留专业术语。
`--strict` 有 high 命中就退出 1；改完再跑，直到只剩可解释的 medium。

### 第 5 步 · 建工程、生成页面、导出

**模式 A（DSH）：**

```bash
# 1) 建 PPTD 工程（必须先建，生成器只写 pages/media/deck.pptd）
#    工具调用：ppt_create({ dir: "work/<任务名>/deck", size: [960, 540], title: "..." })

# 2) 生成页面（校徽/院徽可由 profile 带入，也可显式给 --crest/--logo）
"$PY" "$SKILL_DIR/scripts/build_deck.py" \
  --content work/<任务名>/content.json --profile work/<任务名>/profile.json \
  --out work/<任务名>/deck --media-dir work/<任务名>/figures \
  [--crest 校徽.png] [--logo 院徽.png] [--skin academic] [--motion gentle]

# 3) 导出（中文文字按墨迹框判定会误报溢出，skipCheck: true 后以渲染目检为准）
#    ppt_export({ dir: "work/<任务名>/deck", out: "work/<任务名>/汇报.pptx", skipCheck: true })
```

**模式 B（无插件）：**

```bash
"$PY" "$SKILL_DIR/scripts/build_report.py" --content work/<任务名>/content.json \
  --profile work/<任务名>/profile.json --out work/<任务名>/deck \
  --pptx work/<任务名>/汇报.pptx --media-dir work/<任务名>/figures
```

生成器会打印 `[build_deck][警告]`：未匹配到学校、文案超长、图片缺失、页型未知——逐条处理，不要忽略。

### 第 6 步 · 目检（必做）

- DSH：`ppt_render({ dir: ... })` 出每页 PNG，逐页 `read_image`。
- 若 `ppt_render` 返回空白页（部分环境 LibreOfficeKit 光栅化异常）：改用 LibreOffice 把 pptx 转 PDF（`soffice --headless --convert-to pdf` 或 DSH 的 `office-export-pdf`），再用 PyMuPDF 把 PDF 页面渲染成 PNG：

```bash
python3 - <<'PY'
import pymupdf
doc = pymupdf.open('汇报.pdf')
for i, p in enumerate(doc):
    p.get_pixmap(dpi=110).save(f'qa/p{i+1:02d}.png')
PY
```

逐页看：标题有没有压到校徽、要点有没有出界/重叠、图有没有压字、图注在图下、页脚和页码在不在、留白是否均衡。改完重跑直到干净。

### 第 7 步 · 生成演讲稿（.docx）

```bash
"$PY" "$SKILL_DIR/scripts/build_speech.py" \
  --content work/<任务名>/content.json --profile work/<任务名>/profile.json \
  --pptx work/<任务名>/汇报.pptx --out work/<任务名>/讲稿.docx --md work/<任务名>/讲稿.md
```

- 内容要求：逐页可直接念的讲稿 + 过渡句 + 论文背景补充（研究脉络、概念白话解释、方法直觉、实验细节）+ 术语表 + 预设问答 8–12 条 + 口头表达提示。
- 时长按 220–260 字/分钟配平；20 分钟约 4500–6000 字。
- 讲稿也要过 humanizer：`humanize_lint.py --text work/<任务名>/讲稿.md --hints`，按提示改完重新生成。

### 第 8 步 · 交付

交付清单：`汇报.pptx`（可编辑）、`讲稿.docx`、`content.json`、`profile.json`、PPTD 工程目录、论文图目录；需要时加 PDF。
给路径用绝对路径；在对话里说明页数、时长、学校/模板、动画档位、哪些地方需要人工确认（如期刊分区、学校配色）。

## 文案节奏（防「LLM 味」）

单句写得对不够，全篇节奏才露馅——词都合规、但每页「形状」一样，照样一眼假。硬规则：

1. **导语不要每页都有**：`lead` 覆盖率 ≤40%，图页/表页/公式页通常不需要；30 页里 8–12 页有一句就够。导语不要把本页要点的名词再列一遍——要么给结论，要么给数字，要么删。
2. **数字对仗标题全篇最多 2–3 处**：「一个基站，两种任务」「三句话 + 一组数字」「四个缺口，正好是下一步」这类句式不要页页复用；换陈述、提问、条件或纯名词短语，标题长短错开；连续三页不要都用「……的影响」式同后缀标题。
3. **语气要有波动**：技术描述可以偏书面（「受式 (6f) 约束」），解释再口语；同一个口语判词（垫底/各缺一块/看住）全篇最多一次。
4. **计数词克制**：「三条结论，四个缺口，三处可迁移」一页最多 2 个量词；结构性计数（章节号、对比方案个数）不算。
5. **标签长短错开**：同页不要 4 个以上同一句式的短标签（全是四字、全是「XX 可复用」、全带「更」字都算）；名词、动宾、问句混着来。
6. **收束句只给重点页**：`takeaway`（页面底部收束行）全篇最多 3 页；生成器不加任何前缀，文案里也不要写「小结：」。
7. **金句全篇只说一次**：封面副题那句结论不要再搬进正文页、小结或致谢。
8. **工作备注不上屏**：`请再核对 / 待确认 / 占位 / 引用前` 这类写手提示由 `build_deck.py` 自动剔到讲稿备注，不要把内部核对清单留在页面上。

`humanize_lint.py` 查两层：
- 词句层：数字对仗过多（high）、导语覆盖率 >60%（medium）、连续 6 页都有导语（medium）、标题长度太匀（medium）、导语句式太单一（medium）、内部备注上屏（high）。
- 页面骨架层（`check_deck_structure`）：同页 ≥4 个同构短标签、标签共用词缀（「更×4」「可复用×3」）、单页计数词 ≥3、过渡页副题同模板（箭头链/数字串）、同一句跨页复读、≥3 个「…的影响」式标题、导语复述要点名词、≥4 页收束句、页内要点高度相似、糊墙词收尾。`build_deck.py --check-only` 与 `make_deck.py --verify` 也会做同源检查并告警。

## 页型速查（详见 format_spec.md）

`cover`（自动）/ `paper_card`（自动）/ `toc`（自动）/ `chapter` / `bullets`（核心：分条列点，layout=rows|cards|split）/ `keynums` / `text` / `left_fig` / `fig_bullets` / `big_fig` / `two_figs` / `table` / `formula` / `compare` / `timeline` / `pipeline` / `quote` / `art`（装饰图页）/ `summary` / `roadmap` / `qa` / `thanks`

文案上限：标题 ≤22 字；要点核心词 ≤10 字、说明 ≤46 字；`lead` ≤40 字；备注 ≤120 字。超了生成器会警告。

## 皮肤与动画

| skin | 观感 | 默认用于 |
|---|---|---|
| `academic` | 白底 + 校徽页眉 + 编号标题 + 细线 + 留白 | 组会文献、开题、答辩 |
| `campus` | 品牌色顶条/色块 + 圆点编号 + 卡片 | 课程汇报 |
| `modern` | 无框、结论式大标题、细线条目 | 组会进展 |
| `bold` | 深色封面 + 撞色块 + 粗色条 | 项目/竞赛 |

- `motion`：`none` 无动画 / `gentle` 转场 fade + 要点依次淡入 / `lively` 转场混合 + fly-in。
- 动画只通过生成器的页面 `animations` 与 `export` 配置写入，**不要手工往 deck.pptd 里塞 page.transition**（该版本引擎会判 SchemaBlocker）。本地轨道的转场用 `add_motion.py` 补。

## 学校模板库

```bash
"$PY" "$SKILL_DIR/scripts/build_deck.py" --list-schools 东北       # 查学校（关键字可省）
"$PY" "$SKILL_DIR/scripts/build_deck.py" --school-info 电子科技大学  # 看品牌包与校徽
```

- 优先用 profile 里的 `school` 匹配；匹配不到用中性配色 + 校名文字标识（生成器自动做）。
- 用户自带校徽/院徽：向导里上传，或生成时传 `--crest` / `--logo`；有校徽时可用 `brand_from_logo.py palette --image 校徽.png --json` 自动取色，把结果写进 profile 的 `brand_colors`。
- 校徽素材仅用于个人/教学汇报；对外发布前确认学校素材使用规定。

## 配图与图片生成

1. 优先级：论文原图 → 用户材料 → 检索到的公开图 → 生成图。
2. 论文图必须来自 `extract_figures.py` 抽出的原图（不拉伸、不裁掉坐标轴）；图注写"图 x.x 是什么"。
3. 需要封面/章节装饰：`cover_art.py --style arcs|grid|waves|blocks ... --out media/bg.png`（离线、确定性）；装饰图形放在右侧/底部，不压标题。
4. 需要 AI 概念图：`cover_art.py --ai-prompt ...` 打印英文提示词，再用 DSH 的 `generate_image` 生成，存 `media/` 后作为图片元素插入；风格要克制、与学校主色一致。

## 常见问题

- **规则质检大量"文字垂直溢出"**：中文按墨迹框判定，属误报；`ppt_export` 用 `skipCheck: true`，以渲染目检为准。
- **`ppt_render` 全白**：环境里 LibreOfficeKit 光栅化异常；按第 6 步用 LibreOffice→PDF→PyMuPDF 自建预览。
- **字体**：默认 `Microsoft YaHei`；macOS 没装时 PowerPoint 会回落，排版不变；目标机器是 Windows/Office 时保持 YaHei。
- **模式 B 丢备注/动画**：本地轨道限制；要完整体验用模式 A，或把备注单独写进讲稿。
- **旧版 content.json**（无 `schema_version`）仍能读，但旧页型 `moe_fig/replay_fig/action_fig` 已移除，请改成 `pipeline`/`big_fig`。
- **生成器报"图片找不到"**：`--media-dir` 要给对；content.json 里写文件名即可，绝对路径也支持。

## 命令速查

```bash
"$PY" "$SKILL_DIR/scripts/profile_wizard.py" --skill-dir "$SKILL_DIR" --out profile.json
"$PY" "$SKILL_DIR/scripts/paper_extract.py" --pdf a.pdf --out work/paper
"$PY" "$SKILL_DIR/scripts/extract_figures.py" --pdf a.pdf --out work/figures --max-figures 8
"$PY" "$SKILL_DIR/scripts/verify_claims.py" --content content.json --paper work/paper/paper.md --strict
"$PY" "$SKILL_DIR/scripts/humanize_lint.py" --content content.json --hints
"$PY" "$SKILL_DIR/scripts/build_deck.py" --content content.json --profile profile.json --out deck --media-dir work/figures
"$PY" "$SKILL_DIR/scripts/build_report.py" --content content.json --profile profile.json --out deck --pptx 汇报.pptx --media-dir work/figures
"$PY" "$SKILL_DIR/scripts/make_deck.py" --content content.json --profile profile.json --out deck --pptx 汇报.pptx --speech 讲稿.docx --paper work/paper/paper.md --verify --humanize
"$PY" "$SKILL_DIR/scripts/build_speech.py" --content content.json --profile profile.json --out 讲稿.docx --md 讲稿.md
"$PY" "$SKILL_DIR/scripts/brand_from_logo.py" palette --image 校徽.png --json
"$PY" "$SKILL_DIR/scripts/cover_art.py" --style arcs --primary "#00367F" --out media/bg.png
```

## 示例

`assets/example_content.json` 是 15 页的 v2 示例（低空 ISAC 论文汇报，`schema_version: 2`），改文字即可复用。

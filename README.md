# college-report-ppt

> 全国本科院校通用的中文汇报 PPT 生成器 · 组会文献 / 课程汇报 / 开题中期 / 毕业答辩 / 项目竞赛

输入论文或材料，输出一份能讲的 PPT：学校模板匹配、分条列点提炼、页面转场与入场动画、详细演讲稿 `.docx`，并做「说人话」与结构指纹检查。


## 功能

- **学校模板库**：1416 所本科院校（教育部口径），双一流 147 所有配色/校徽/校训品牌包，其他学校自动降级为文字标识 + 中性配色；支持自带校徽、院徽并自动取色。
- **开始前向导**：本地离线网页，选学校/学院/姓名/身份/目的/时长/皮肤/动画，上传校徽院徽，输出 `profile.json`。
- **论文提炼 + 回查**：PDF 抽正文、抽图、生成论文卡；`verify_claims.py` 把数字/术语/图注回查原文，不编数据。
- **分条列点排版**：19 种页型、4 套皮肤（学术/校园/现代/撞色），结论式标题、每页 4–5 条要点、留白优先。
- **公式与图表**：`render_formula.py` 用 matplotlib 渲染 LaTeX 公式成图，编号统一右对齐；论文图自动拆 a/b 子图。
- **动画**：页面转场 + 元素入场（依赖 PPT 工具链；本地兜底导出补转场）。
- **演讲稿 .docx**：逐页讲稿 + 背景补充 + 术语表 + 预设问答 + 口头提示，20 分钟约 4000–4500 字。
- **防 AI 味**：词句层 + 页面骨架层两层检查（数字对仗、同构标签、跨页复句、导语复述、页内重复等），可自动生成改写建议。

## 安装

```bash
cp -R college-report-ppt ~/.dsh/skills/
```

（把本仓库根目录放进 `~/.dsh/skills/` 下、命名 `college-report-ppt` 亦可。）装完重开会话，技能列表里会出现 `college-report-ppt`。

依赖：Python 3.9+；`pyyaml`、`python-pptx`、`Pillow`、`python-docx`；可选 `PyMuPDF`（PDF 提图/预览）、`matplotlib`（公式渲染）。模式 A 直接生成 `.pptx` 需要 `dsh-ppt-maker` 插件，模式 B（先写 PPTD 工程）不需要。

## 快速开始

```bash
# 1) 向导：生成 profile.json（不弹窗也可 --plain / --answers）
python3 scripts/profile_wizard.py --out profile.json

# 2) 内容：参考 assets/example_content.json 写 content.json

# 3) 一条命令：构建 → 导出 pptx → 讲稿 → 回查 → 说人话检查
python3 scripts/make_deck.py \
  --content content.json --profile profile.json --out deck_project \
  --media-dir media --pptx 汇报.pptx --speech 讲稿.docx \
  --verify --paper 论文.md --figures-json media/figures.json --humanize

# 只检查文案：
python3 scripts/build_deck.py --content content.json --check-only
python3 scripts/humanize_lint.py --content content.json
```

## 目录

```
college-report-ppt/
├── SKILL.md              # 技能入口（完整工作流）
├── 安装说明.md
├── requirements.txt
├── scripts/              # 构建/讲稿/检查/向导/公式/提图等
├── references/           # 格式契约、品牌 schema、设计语言、论文提取、讲稿指南、humanizer 规则
├── assets/
│   ├── example_content.json
│   ├── wizard/
│   └── schools/
│       ├── schools.json      # 1416 所本科院校
│       ├── brand_packs.json  # 147 个学校品牌包
│       └── crests/           # 145 个校徽 PNG
```

## 设计约定（摘要）

- 标题默认平实学术式，金句只留给章节过渡页和核心页；导语覆盖率 ≤40%。
- 正文页只保留标题下一条细线；过渡页无花纹、无大数字水印。
- 页面骨架检查：同页 ≥4 个同构标签、单页 ≥3 个计数词、≥3 个同后缀标题、同一句跨页复读、导语复述要点、≥4 页「小结：」等。
- 数字以论文原文为准；论文没写的参数标「论文没有写」，不替作者补。

## License

MIT

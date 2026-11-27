# 模板匹配向导（profile_wizard.py）· Agent 使用说明

`profile_wizard.py` 负责在开始做 PPT 前收集**汇报档案**：学校 / 学院 / 姓名 / 身份 / 目的 / 时长 / 日期 / 皮肤 / 动画档位 / 「说人话」开关 / 校徽院徽，输出严格符合 `references/profile_schema.md` 的 `profile.json`。本文件说明 agent 什么时候用它、怎么用、坏了怎么退。

## 1. 什么时候用

用：

- 用户说「帮我做个汇报 PPT」「组会汇报」「开题/答辩 PPT」，但**还没给**学校、姓名、身份、目的、时长等信息；
- 用户第一次用这个技能，需要把「学院格式」配置下来；
- 用户换了学校/学院，要重新生成 profile。

不用：

- 用户在聊天里已经把这些信息说清楚了 → 直接按 §3 手写 `profile.json` 更快；
- 用户只是让你改**已存在**的 PPT 或只补一页内容 → 不需要向导；
- 无图形界面 / 用户明确不想开浏览器 → 用 `--plain` / `--answers`（见 §2）。

> 向导是**可选加速器**，不是硬依赖。`build_deck.py` 只读 profile.json，手写的同样能用。

## 2. 三条使用路径

解释器统一用 DSH 自带 Python（已内置 Pillow / python-pptx / lxml）：

```bash
PY="${DSH_PYTHON:-}"
if [ -z "$PY" ]; then
  for c in "${DSH_HOME:-$HOME/.dsh}"/dsh-runtimes/*/dependencies/python/bin/python3; do
    [ -x "$c" ] && PY="$c" && break
  done
fi
[ -n "$PY" ] || PY="$(command -v python3)"
SKILL_DIR="<技能安装目录>"
```

### 2.1 网页向导（用户自己在电脑前操作，推荐）

```bash
"$PY" "$SKILL_DIR/scripts/profile_wizard.py" \
  --skill-dir "$SKILL_DIR" \
  --out "<项目目录>/profile.json"
```

- 自动在 `127.0.0.1` 起一个本地网页（默认随机端口），并尝试用系统浏览器打开；打不开时终端会打印 URL，让用户手动打开。
- 页面全中文、完全离线（无 CDN、无第三方字体），三步：① 选学校（搜索中文名/拼音/首字母/别名，显示省份+层次+校徽缩略图）+ 学院；② 姓名、身份、目的（6 张卡片，带默认时长/页数）、时长滑块、日期；③ 皮肤（4 套 + auto）、动画档位、「说人话」开关、上传校徽/院徽（可选，自动取色）。
- 右侧实时预览：选中学校后用它自己的 primary/accent 渲染预览条、校徽、校名、页脚。
- 提交后写 `profile.json`，页面显示配置摘要与文件路径，服务端 2 秒后自动退出；`--timeout`（默认 900 秒）内没人操作也会退出。

常用参数：

| 参数 | 说明 |
|---|---|
| `--skill-dir` | 技能根目录（默认按脚本位置自动推断，一般不用写） |
| `--out` | 输出 profile.json 路径（**必填**） |
| `--port 0` | 0 = 系统随机端口；指定端口被占用会自动顺延重试，最后退回随机 |
| `--no-open` | 不自动开浏览器，只打印 URL（无头环境 / 远程协助时用） |
| `--timeout 900` | 等待超时秒数 |

**给用户的话术**：告诉用户「马上会打开一个本地网页，三步填完点『生成汇报档案』，然后回聊天说一声」。不要替用户点。

### 2.2 纯文本模式（agent 把问题读给用户）

```bash
"$PY" "$SKILL_DIR/scripts/profile_wizard.py" --out "<项目目录>/profile.json" --plain
```

- 不开浏览器；把 8 个问题**逐条打印到 stdout**（可直接读给用户，或转成聊天里的提问）。
- 如果 stdin 是交互终端，会逐项提问；在 agent 环境里通常不是终端，所以它只打印问题并提示改用 `--answers`。

### 2.3 answers.json 模式（agent 在聊天里问完直接生成，最省事）

先按 §4 用 `ask_user_question` 问用户，再写一个 answers.json：

```json
{
  "school": "示例大学",
  "school_slug": "neu",
  "college": "示例学院",
  "college_en": "Sydney Smart Technology College",
  "reporter": "张三",
  "role": "研究生",
  "purpose": "group_paper",
  "duration_min": 20,
  "date": "2026 年 11 月",
  "skin": "auto",
  "motion": "gentle",
  "humanize": true,
  "crest": "<你的目录>/校徽.png",
  "logo": "<你的目录>/院徽.png"
}
```

```bash
"$PY" "$SKILL_DIR/scripts/profile_wizard.py" \
  --out "<项目目录>/profile.json" \
  --answers "<项目目录>/answers.json"
```

- 只有 `school`、`reporter`、`role`、`purpose` 影响必填校验；`purpose` 也可以写中文名（「组会文献汇报」）或序号（`2`），`role` 可以写「硕士」「博士研究生」等常见说法。
- `crest` / `logo` 是本机图片路径，会**复制**到 `<out 目录>/../user_assets/` 并写入 profile 的 `crest_override` / `logo_override`；校徽还能自动取色填 `brand_colors`。
- `brand_colors` 可显式指定：`{"primary": "#00367F", "accent": "#0055AA"}`（给了就不再从图片取色）。
- 生成后会打印摘要与 warnings（比如「没匹配到学校，生成器会用中性配色」）。

## 3. 不走向导：手写 profile.json（推荐用于快速改配置）

按 `references/profile_schema.md` 写最小集即可：

```json
{
  "schema_version": 2,
  "school": "电子科技大学",
  "college": "信息与通信工程学院",
  "reporter": "张三",
  "role": "研究生",
  "purpose": "group_paper",
  "duration_min": 20,
  "skin": "auto",
  "motion": "gentle",
  "humanize": true
}
```

其余字段缺省规则见契约。向导额外会填：`school_slug`、`audience`（按目的推导）、`motto`（品牌包校训）、`brand_colors`（从校徽取色时）、`created_at`、`lang`。

## 4. 失败回退：ask_user_question 问 6 题

向导起不来（无 Python / 用户不想开网页 / 服务异常）时，直接问用户，然后手写 profile.json。用 DSH 的 `ask_user_question` 分 6 个问题问：

1. **学校 + 学院**：学校中文全称（不确定就多写别名），学院/系名可留空。
2. **汇报人 + 身份**：姓名；身份选 `本科生 / 研究生 / 博士生 / 教师 / 其他`。
3. **汇报目的**（六选一，决定默认结构、时长与皮肤）：
   - `course_report` 个人课程汇报（默认 9 分钟 / 12 页 / campus）
   - `group_paper` 组会文献汇报（默认 22 分钟 / 18 页 / academic）
   - `group_progress` 组会进展汇报（默认 12 分钟 / 10 页 / modern）
   - `proposal` 开题 / 中期检查（默认 18 分钟 / 15 页 / academic）
   - `defense` 毕业答辩（默认 26 分钟 / 22 页 / academic）
   - `project` 项目 / 竞赛汇报（默认 15 分钟 / 15 页 / bold）
4. **时长 + 日期**：计划讲多少分钟；封面日期（默认当前年月，如「2026 年 11 月」）。
5. **皮肤 + 动画 + 说人话**：皮肤 `auto / academic / campus / modern / bold`（默认 auto）；动画 `none / gentle / lively`（默认 gentle）；「说人话」默认开。
6. **校徽 / 院徽**（可选）：有没有学校校徽、院徽/实验室 logo 图片，有就给本机路径（PNG/SVG/JPG）。

拿到答案后：

- 直接手写 `profile.json`（§3）；或
- 写成 `answers.json` 跑 `--answers`（§2.3），让工具统一处理取色/素材落盘。

要点：**必填只有 学校、姓名、身份、目的**；其余给不出一律用默认值，不要为了凑字段反复追问用户。

## 5. 校徽 / 院徽 上传处理

- 支持 PNG / SVG / JPG / WebP / GIF；单张 ≤ 6 MB（页面上限 4 MB）。
- 落盘位置：**`<out 目录>/../user_assets/`**（例如 `--out work/我的PPT/profile.json` → `work/user_assets/`），文件名 `crest_<slug>_<hash>.png` / `logo_<slug>_<hash>.<ext>`。
- profile 里只存**本机绝对路径**：
  - 校徽 → `crest_override`（优先于模板库校徽）；
  - 院徽/实验室 logo → `logo_override`（封面右下 96×96、页脚右侧 22×22 以内）。
- 取色：
  - 上传后页面立即调用 `brand_from_logo.py` 的取色逻辑，显示主色/辅色/深色/浅色色板；
  - 提交时若从校徽取到色，写入 `brand_colors`（含 `primary/accent/deep/soft`）；优先校徽，其次院徽；
  - **SVG 可以保存，但无法自动取色**（Pillow 不解 SVG）。需要配色时让用户给 PNG，或显式给 `brand_colors`。
- 用户没上传时：`crest_override`/`logo_override` 留空字符串，生成器用品牌包校徽或调 `brand_from_logo.py text-crest` 生成文字校徽。

## 6. 隐私说明（可以原样告诉用户）

- 向导只监听 `127.0.0.1`，**不联网、不上传**任何数据；页面资源全部内嵌，无 CDN。
- 姓名、学校、校徽图片只写在本机：`profile.json` + `user_assets/`。
- 生成结束后服务自动退出；不会常驻、不会留后台进程。

## 7. 故障排查

| 现象 | 原因 / 处理 |
|---|---|
| 浏览器没自动打开 | 终端里会打印 `http://127.0.0.1:端口/`，复制到浏览器即可；远程/无头环境用 `--no-open` |
| 提示「端口被占用，已改用 …」 | 正常自动顺延；也可以直接 `--port 0` |
| 页面显示「未找到 schools.json，用内置 15 所样例」 | 学校库还没构建好 / 放错位置；把 `schools.json` 放到 `<技能根>/assets/schools/` 后重启向导 |
| 搜索到学校但预览是中性色 | 该校没有品牌包（`brand_packs.json` 缺该项），生成时也会用 `#1F3A5F` 中性色 |
| 上传后提示「SVG 无法自动取色」 | 正常限制；换 PNG 或在 profile 里显式写 `brand_colors` |
| 提示缺少 Pillow | 换 DSH 自带 Python（内置 Pillow），或 `python3 -m pip install Pillow` |
| 900 秒后自己退出 | 超时保护；重新跑一次即可，浏览器刷新不会恢复 |
| 提交后报 400 | 服务端校验：学校 / 姓名 / 目的必填；看页面底部红色提示 |
| profile 写不进去 | `--out` 目录不可写；换到项目目录下再试 |

## 8. 与后续脚本的衔接

```bash
# 1) 向导产出 profile.json（本文件 §2）
# 2) 生成 PPTD 工程 / pptx
"$PY" "$SKILL_DIR/scripts/build_deck.py" content.json --profile profile.json --out-dir ...
# 3) 本地兜底导出链路补转场
"$PY" "$SKILL_DIR/scripts/add_motion.py" --pptx out.pptx --out out-motion.pptx --mode fade
```

- `cover_art.py` 可生成封面装饰图（`--style arcs|grid|waves|blocks`，`--ai-prompt` 给 DSH `generate_image` 拼英文提示词）。
- `brand_from_logo.py` 独立可用：`palette` 取色、`text-crest` 生成文字校徽、`make-pack` 拼品牌包。

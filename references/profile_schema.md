# profile.json — 汇报档案（向导输出）

`profile_wizard.py` 生成，`build_deck.py` 读取。用户没走向导时，agent 按本规范手写一份。

```json
{
  "schema_version": 2,
  "school": "示例大学",
  "school_slug": "neu",
  "college": "示例学院",
  "college_en": "Sydney Smart Technology College",
  "reporter": "张三",
  "role": "研究生",
  "purpose": "group_paper",
  "audience": "课题组组会",
  "duration_min": 20,
  "date": "2026 年 11 月",
  "skin": "auto",
  "motion": "gentle",
  "humanize": true,
  "footer_left": "",
  "footer_right": "",
  "motto": "",
  "crest_override": "",
  "logo_override": "",
  "brand_colors": null,
  "lang": "zh",
  "created_at": "2026-11-27T12:00:00+08:00"
}
```

## 字段

| 字段 | 必填 | 取值 / 说明 |
|---|---|---|
| `school` | ✔ | 学校中文全称。生成器会在 `assets/schools/schools.json` 里做精确→别名→模糊匹配 |
| `school_slug` |  | 学校 slug（如 `neu`）。知道就直接给，省一次匹配 |
| `college` |  | 学院 / 系名，显示在封面与页脚 |
| `college_en` |  | 学院英文名，显示在页脚右侧 |
| `reporter` | ✔ | 汇报人姓名 |
| `role` | ✔ | `本科生` / `研究生` / `博士生` / `教师` / `其他` |
| `purpose` | ✔ | 见下表 |
| `audience` |  | 听众，如「课题组组会」「课程班」「答辩委员会」 |
| `duration_min` |  | 计划时长（分钟），用于演讲稿时间分配；缺省按目的给默认值 |
| `date` |  | 封面日期，如「2026 年 11 月」；缺省当前年月 |
| `skin` |  | `auto` / `academic` / `campus` / `modern` / `bold`；`auto` 按目的挑 |
| `motion` |  | `none` / `gentle` / `lively`；缺省 `gentle` |
| `humanize` |  | `true` 时正文与讲稿都过「说人话」检查（默认 true） |
| `footer_left` |  | 页脚左文字覆盖（缺省用品牌 motto） |
| `footer_right` |  | 页脚右文字覆盖（缺省用 college_en 或品牌英文名） |
| `motto` |  | 校训覆盖 |
| `crest_override` |  | 用户提供的校徽图片绝对路径（优先于模板库校徽） |
| `logo_override` |  | 用户提供的院徽 / 实验室 logo 绝对路径（小尺寸，放封面角落与页脚） |
| `brand_colors` |  | 用户提供的配色，`{"primary":"#00367F","accent":"#0055AA"}`；从 logo 取色时由 `brand_from_logo.py` 填 |
| `lang` |  | `zh`（默认）/ `en` |

## purpose 取值

| 值 | 中文 | 默认页数 | 默认时长 | 默认 skin | 结构要点 |
|---|---|---|---|---|---|
| `course_report` | 个人课程汇报 | 10–14 | 8–10 min | `campus` | 选题背景 → 做了什么 → 结果/作品 → 收获与反思 |
| `group_paper` | 组会文献汇报 | 15–20 | 18–25 min | `academic` | 论文简介 → 背景 → 方法 → 实验 → 结论与启发 |
| `group_progress` | 组会进展汇报 | 8–12 | 10–15 min | `modern` | 上次结论 → 本期工作 → 数据/结果 → 卡点 → 下一步 |
| `proposal` | 开题 / 中期检查 | 12–18 | 15–20 min | `academic` | 背景与意义 → 现状与不足 → 研究内容 → 技术路线 → 计划与基础 |
| `defense` | 毕业答辩 | 18–25 | 20–30 min | `academic` | 背景 → 方法 → 实验 → 结论 → 创新点 → 展望 |
| `project` | 项目 / 竞赛汇报 | 12–18 | 10–20 min | `bold` | 痛点 → 方案 → 亮点 → 数据 → 团队与计划 |

## 默认值与自动规则

- 缺 `school_slug`：生成器按 `school` 匹配，匹配不到用中性配色（中性深蓝 `#1F3A5F`）并打印警告，不阻断。
- 缺 `reporter`：封面留空下划线占位，并在 stderr 提示。
- `skin: auto`：按 `purpose` 取上表默认；`purpose` 不在表里按 `academic`。
- `motion: gentle`：转场 fade；封面标题 fadeIn；正文要点块 afterPrevious 依次淡入，每页 ≤3 组。
- `motion: lively`：转场用引擎 `auto`（引擎自带 fade / push / wipe 混合），要点块 flyIn 轻位移，每页 ≤3 组。
- `motion: none`：不加动画、转场 fade（保持可放映的基本观感）；打印类 deck 用这个。

## 手写示例（没走向导时）

```json
{
  "schema_version": 2,
  "school": "电子科技大学",
  "college": "信息与通信工程学院",
  "reporter": "张三",
  "role": "研究生",
  "purpose": "group_paper",
  "duration_min": 20,
  "skin": "academic",
  "motion": "gentle",
  "humanize": true
}
```

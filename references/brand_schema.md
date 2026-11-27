# 学校模板库数据规范

模板库分两份数据，按 `slug` 关联：

1. `assets/schools/schools.json` — 身份库（全国本科院校名单）。
2. `assets/schools/brand_packs.json` — 品牌包（配色 / 校徽 / 页脚 / 版式倾向）。

生成器 `build_deck.py` 用 `school` 或 `slug` 查身份库，再取品牌包；品牌包缺的字段用默认规则推导。

## 1. schools.json

```json
[
  {
    "slug": "neu",
    "name": "示例大学",
    "name_en": "Northeastern University",
    "aliases": ["东大", "NEU"],
    "province": "示例省",
    "city": "示例市",
    "level": ["985", "211", "双一流"],
    "type": "理工",
    "pinyin": "dongbeidaxue",
    "abbr": "dbdx",
    "website": "https://www.neu.edu.cn"
  }
]
```

| 字段 | 必填 | 说明 |
|---|---|---|
| `slug` | ✔ | 小写字母/数字/连字符，全局唯一；品牌包与校徽文件名都用它 |
| `name` | ✔ | 中文全称，官方名 |
| `name_en` |  | 英文名（有则填） |
| `aliases` |  | 简称、旧名、俗名；搜索与匹配用 |
| `province` | ✔ | 省级行政区（不带「省」「市」，如 `示例省`、`北京`、`内蒙古`） |
| `city` |  | 地级市；直辖市填与 province 相同的值 |
| `level` | ✔ | 数组，可含 `985`、`211`、`双一流`、`部属`、`省属`；普通本科为 `[]` |
| `type` | ✔ | `综合` / `理工` / `师范` / `农林` / `医药` / `财经` / `政法` / `语言` / `艺术` / `体育` / `民族` / `军事` / `其他` |
| `pinyin` | ✔ | 全拼小写无空格；搜索用 |
| `abbr` |  | 拼音首字母（如 `dbdx`），用户可输入 `dbdx` 直接命中 |
| `website` |  | 官网根地址 |

要求：覆盖全国**所有本科院校**（教育部最新全国高等学校名单口径，约 1300+ 所），逐条校验字段完整；同校异名合并为 aliases。

## 2. brand_packs.json

```json
{
  "neu": {
    "school": "示例大学",
    "school_en": "Northeastern University",
    "colors": {
      "primary": "#00367F",
      "accent": "#0055AA",
      "deep": "#002B5C",
      "soft": "#EAF1FA",
      "ink": "#1A1A1A",
      "gold": "#C9A063"
    },
    "crest": "crests/neu.png",
    "crest_kind": "image",
    "crest_source": "official-website",
    "crest_verified": true,
    "motto": "自强不息 知行合一",
    "footer_left": "MAKE GLOBAL CHANGES",
    "footer_right": "Sydney Smart Technology College",
    "decor": "rule",
    "skin_hint": "academic",
    "college_default": "示例学院",
    "source": "VI 标准色 / 官网",
    "verified": true
  }
}
```

| 字段 | 必填 | 说明 |
|---|---|---|
| `school` / `school_en` | ✔ | 与身份库一致 |
| `colors.primary` | ✔ | 主色，校徽/校名标准色；用于标题、页眉、章节页 |
| `colors.accent` | ✔ | 辅助色，比主色亮一档；用于强调、图表 |
| `colors.deep` |  | 深色，用于深底页/正文重点；缺省由 primary 压暗 25% |
| `colors.soft` |  | 浅色底，用于卡片/表头；缺省由 primary 提亮 92% |
| `colors.ink` |  | 正文色，缺省 `#1A1A1A` |
| `colors.gold` |  | 点缀色（校徽金/银），可缺省 |
| `crest` |  | 相对 `assets/schools/` 的路径；没有校徽就填 `null` + `crest_kind: "text"` |
| `crest_kind` |  | `image` / `text`（文字标识，由 `brand_from_logo.py text-crest` 生成） |
| `crest_source` |  | `official-website` / `wikimedia` / `dataset` / `user` / `estimated` |
| `crest_verified` |  | 是否人工看过（像素、透明底、无白边） |
| `motto` |  | 校训；页脚左或封面底 |
| `footer_left` / `footer_right` |  | 页脚文字；缺省 `footer_left` 用 motto、`footer_right` 用 school_en |
| `decor` |  | `rule`（细线）/ `block`（色块）/ `arc`（弧线）/ `minimal`；缺省 `rule` |
| `skin_hint` |  | 该校最合适的默认皮肤；缺省 `academic` |
| `verified` |  | 配色是否经人工核对（VI 标准色） |
| `source` |  | 数据来源说明 |

## 3. 缺校徽 / 缺品牌包时的降级规则（生成器必须实现）

1. 有品牌包但 `crest: null` → 调 `brand_from_logo.py text-crest` 生成「校名文字标识」PNG（主色底 + 白字），缓存到 `assets/schools/crests/<slug>-text.png`。
2. 完全没有品牌包 → 用中性配色（`primary #1F3A5F`），校徽位置放文字标识；stderr 打印「未找到 X 的品牌包，已用中性配色」。
3. 用户提供 `crest_override` → 优先；同时跑 `brand_from_logo.py palette` 从校徽取主色，覆盖 primary/accent（除非 profile 里显式给了 brand_colors）。
4. `logo_override`（院徽/实验室 logo）→ 放封面右下与页脚右侧，尺寸不超过 22×22 px（页脚）/ 96×96 px（封面）。

## 4. 颜色对比约束（生成器与数据都要满足）

- `primary` 与白色对比度 ≥ 4.5（放白底当文字/深色底白字都够用）。
- `soft` 作为背景时，其上正文（`ink`）对比度 ≥ 7。
- `accent` 只做点缀，不做大面积正文色。

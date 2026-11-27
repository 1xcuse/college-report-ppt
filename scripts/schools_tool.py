#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""全国本科院校身份库查询工具（只依赖标准库）。

数据：assets/schools/schools.json（教育部 2026 全国普通高校名单·本科口径）。

用法：
    python3 scripts/schools_tool.py stats
    python3 scripts/schools_tool.py list --province 示例省 --level 985
    python3 scripts/schools_tool.py search 东北
    python3 scripts/schools_tool.py resolve "示例大学"
    python3 scripts/schools_tool.py validate
    python3 scripts/schools_tool.py export-csv --out schools.csv

供 build_deck.py 复用的接口：
    from schools_tool import resolve_school, load_schools, get_by_slug
    rec = resolve_school("示例大学")   # -> dict（含 _score/_match）或 None
"""
import argparse
import csv
import difflib
import json
import os
import re
import sys
import unicodedata

DATA_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'assets', 'schools', 'schools.json',
)

PROVINCES = [
    '北京', '天津', '河北', '山西', '内蒙古', '示例省', '吉林', '黑龙江', '上海', '江苏',
    '浙江', '安徽', '福建', '江西', '山东', '河南', '湖北', '湖南', '广东', '广西',
    '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃', '青海', '宁夏', '新疆',
]
PROVINCE_ALIASES = {
    '北京市': '北京', '天津市': '天津', '上海市': '上海', '重庆市': '重庆',
    '河北省': '河北', '山西省': '山西', '内蒙古自治区': '内蒙古', '示例省省': '示例省',
    '吉林省': '吉林', '黑龙江省': '黑龙江', '江苏省': '江苏', '浙江省': '浙江',
    '安徽省': '安徽', '福建省': '福建', '江西省': '江西', '山东省': '山东',
    '河南省': '河南', '湖北省': '湖北', '湖南省': '湖南', '广东省': '广东',
    '广西壮族自治区': '广西', '广西省': '广西', '海南省': '海南', '四川省': '四川',
    '贵州省': '贵州', '云南省': '云南', '西藏自治区': '西藏', '陕西省': '陕西',
    '甘肃省': '甘肃', '青海省': '青海', '宁夏回族自治区': '宁夏',
    '新疆维吾尔自治区': '新疆',
}
LEVELS = ['985', '211', '双一流', '部属', '省属']
LEVEL_ORDER = {v: i for i, v in enumerate(LEVELS)}
TYPES = ['综合', '理工', '师范', '农林', '医药', '财经', '政法', '语言', '艺术',
         '体育', '民族', '军事', '其他']
REQUIRED_FIELDS = ['slug', 'name', 'province', 'level', 'type', 'pinyin']
FIELDS = ['slug', 'name', 'name_en', 'aliases', 'province', 'city', 'level',
          'type', 'pinyin', 'abbr', 'website']

_SCHOOLS = None
_ALIAS_INDEX = None


# ------------------------------------------------------------------ 基础
def _norm(text):
    """匹配用归一化：全角括号转半角、去空白、转小写。"""
    if text is None:
        return ''
    text = str(text).strip()
    text = text.replace('（', '(').replace('）', ')')
    text = re.sub(r'\s+', '', text)
    return text.lower()


def _norm_province(value):
    if not value:
        return ''
    v = str(value).strip()
    if v in PROVINCES:
        return v
    if v in PROVINCE_ALIASES:
        return PROVINCE_ALIASES[v]
    for suf in ('壮族自治区', '回族自治区', '维吾尔自治区', '自治区', '省', '市'):
        if v.endswith(suf):
            return PROVINCE_ALIASES.get(v, v[:-len(suf)])
    return v


def load_schools(path=None, force=False):
    """读取 schools.json（带缓存）。文件缺失或损坏时抛 SystemExit 并打印中文提示。"""
    global _SCHOOLS
    if _SCHOOLS is not None and not force and path is None:
        return _SCHOOLS
    p = path or DATA_PATH
    try:
        with open(p, encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        sys.stderr.write('错误：找不到学校数据文件 %s\n' % p)
        sys.stderr.write('请确认技能目录结构完整（assets/schools/schools.json）。\n')
        raise SystemExit(2)
    except json.JSONDecodeError as e:
        sys.stderr.write('错误：学校数据文件不是合法 JSON（%s）：%s\n' % (p, e))
        raise SystemExit(2)
    if not isinstance(data, list) or not data:
        sys.stderr.write('错误：学校数据文件应为非空数组：%s\n' % p)
        raise SystemExit(2)
    if path is None:
        _SCHOOLS = data
    return data


def _alias_index(schools):
    global _ALIAS_INDEX
    if _ALIAS_INDEX is not None:
        return _ALIAS_INDEX
    idx = {}
    for rec in schools:
        for a in rec.get('aliases') or []:
            idx.setdefault(_norm(a), []).append(rec)
    _ALIAS_INDEX = idx
    return idx


def get_by_slug(slug):
    """按 slug 精确查一条；找不到返回 None。"""
    for rec in load_schools():
        if rec.get('slug') == slug:
            return rec
    return None


def get_by_name(name):
    """按官方名精确查一条（全/半角括号、空白不敏感）；找不到返回 None。"""
    want = _norm(name)
    for rec in load_schools():
        if _norm(rec.get('name')) == want:
            return rec
    return None


def _level_rank(rec):
    """层次优先级：985 > 211 > 双一流 > 普通；同分用于打破平局。"""
    lv = rec.get('level') or []
    if '985' in lv:
        return 4
    if '211' in lv:
        return 3
    if '双一流' in lv:
        return 2
    return 1 if ('部属' in lv or '省属' in lv) else 0


def _search_scored(query):
    """对全部学校打分，返回 [(score, match_kind, rec)]，按分数降序。"""
    q = _norm(query)
    if not q:
        return []
    schools = load_schools()
    aliases = _alias_index(schools)
    out = []
    for rec in schools:
        name = _norm(rec.get('name'))
        py = _norm(rec.get('pinyin'))
        ab = _norm(rec.get('abbr'))
        score, kind = 0, ''
        if q == name:
            score, kind = 100, '名称精确'
        elif q in aliases and rec in aliases[q]:
            score, kind = 95, '别名精确'
        elif py and q == py:
            score, kind = 92, '拼音精确'
        elif ab and q == ab:
            score, kind = 90, '首字母精确'
        else:
            cands = [name, py, ab]
            kinds = ['名称前缀', '拼音前缀', '首字母前缀']
            for value, k in zip(cands, kinds):
                if value and value.startswith(q):
                    score, kind = 84, k
                    break
            if not score:
                for value, k in ((name, '名称包含'), (py, '拼音包含'), (ab, '首字母包含')):
                    if value and q in value:
                        score, kind = 78, k
                        break
            if not score:
                alias_hit = [a for a in (_norm(x) for x in (rec.get('aliases') or []))
                             if q and (a.startswith(q) or q in a)]
                if alias_hit:
                    score, kind = 76, '别名包含'
            if not score and len(q) >= 2:
                ratio = max(difflib.SequenceMatcher(None, q, name).ratio(),
                            difflib.SequenceMatcher(None, q, py).ratio() if py else 0)
                if ratio >= 0.55:
                    score, kind = 60 * ratio, '模糊匹配'
        if score > 0:
            out.append((score, kind, rec))
    out.sort(key=lambda t: (-t[0], -_level_rank(t[2]), len(t[2].get('name', '')),
                            t[2].get('slug', '')))
    return out


# ------------------------------------------------------------------ 公共 API
def resolve_school(query, threshold=50.0):
    """把用户输入解析成一条学校记录。

    匹配顺序：官方名精确 → 别名精确 → 拼音/首字母精确 → 前缀/子串 → difflib 模糊。
    返回匹配记录 dict（额外带 _score / _match），分数低于 threshold 时返回 None。
    找不到时可用 resolve_school_candidates(query, 3) 取候选。
    """
    scored = _search_scored(query)
    if not scored:
        return None
    score, kind, rec = scored[0]
    if score < threshold:
        return None
    hit = dict(rec)
    hit['_score'] = round(score, 1)
    hit['_match'] = kind
    return hit


def resolve_school_candidates(query, n=3):
    """返回最接近的 n 个候选（list of dict，带 _score / _match）。

    没有启发式命中时，仍用 difflib 硬给出最相似的 n 条（方便用户纠错）。
    """
    scored = _search_scored(query)
    out = []
    for score, kind, rec in scored[:n]:
        hit = dict(rec)
        hit['_score'] = round(score, 1)
        hit['_match'] = kind
        out.append(hit)
    if len(out) < n:
        q = _norm(query)
        seen = {r['slug'] for r in out}
        pool = []
        for rec in load_schools():
            if rec.get('slug') in seen:
                continue
            ratio = difflib.SequenceMatcher(None, q, _norm(rec.get('name'))).ratio()
            py = _norm(rec.get('pinyin'))
            if py:
                ratio = max(ratio, difflib.SequenceMatcher(None, q, py).ratio() * 0.9)
            pool.append((ratio, rec))
        pool.sort(key=lambda t: (-t[0], len(t[1].get('name', '')), t[1].get('slug', '')))
        for ratio, rec in pool[:n - len(out)]:
            hit = dict(rec)
            hit['_score'] = round(60 * ratio, 1)
            hit['_match'] = '模糊匹配'
            out.append(hit)
    return out


# ------------------------------------------------------------------ 格式化
def _level_text(rec):
    lv = rec.get('level') or []
    return '/'.join(lv) if lv else '普通本科'


def _disp_len(text):
    """终端显示宽度：中日韩全角字符按 2 列计。"""
    width = 0
    for ch in str(text):
        width += 2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1
    return width


def _pad(text, width):
    text = str(text)
    return text + ' ' * max(0, width - _disp_len(text))


def _print_table(headers, rows):
    widths = [_disp_len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], _disp_len(cell))
    print('  '.join(_pad(h, widths[i]) for i, h in enumerate(headers)))
    print('  '.join('-' * widths[i] for i in range(len(headers))))
    for row in rows:
        print('  '.join(_pad(cell, widths[i]) for i, cell in enumerate(row)))


# ------------------------------------------------------------------ 子命令
def cmd_stats(args):
    schools = load_schools()
    total = len(schools)
    level_counts = {}
    for rec in schools:
        for lv in rec.get('level') or []:
            level_counts[lv] = level_counts.get(lv, 0) + 1
    plain = sum(1 for r in schools if not (r.get('level') or []))
    type_counts = {}
    for rec in schools:
        t = rec.get('type') or '其他'
        type_counts[t] = type_counts.get(t, 0) + 1
    prov_counts = {}
    for rec in schools:
        p = rec.get('province') or ''
        prov_counts[p] = prov_counts.get(p, 0) + 1
    print('学校总数：%d 所（本科）' % total)
    print('口径：教育部《全国高等学校名单》截至 2026-06-17 的普通高校本科 1412 所，'
          '另补 4 所军校（该名单不含军队院校），合计 1416 所。')
    print('')
    print('分层次（标签可重叠，普通本科为无标签）：')
    _print_table(['层次', '学校数'],
                 [[lv, level_counts.get(lv, 0)] for lv in LEVELS] + [['普通本科', plain]])
    print('  注：211 官方「116 所」口径 = 112 所高校 + 4 所两地办学（中国地质大学、中国石油大学、')
    print('      中国矿业大学、华北电力大学 各按两地计）；本库按「学校」建条，华电保定校区')
    print('      并入华北电力大学 aliases，故 211 标签为 %d 条（116 - 1）。' % level_counts.get('211', 0))
    print('')
    print('分类型：')
    _print_table(['类型', '学校数'],
                 [[t, type_counts[t]] for t in TYPES if type_counts.get(t)])
    print('')
    print('分省：')
    rows = [[p, prov_counts.get(p, 0)] for p in PROVINCES if prov_counts.get(p)]
    rest = sorted((p, c) for p, c in prov_counts.items() if p not in PROVINCES)
    _print_table(['省级行政区', '学校数'], rows + [list(x) for x in rest])
    print('')
    print('合计：%d 所，覆盖 %d 个省级行政区' % (total, len([p for p in prov_counts if prov_counts[p]])))


def cmd_list(args):
    schools = load_schools()
    prov = _norm_province(args.province) if args.province else ''
    rows = []
    for rec in schools:
        if prov and rec.get('province') != prov:
            continue
        if args.level and args.level not in (rec.get('level') or []):
            continue
        if args.type and rec.get('type') != args.type:
            continue
        rows.append([rec.get('slug'), rec.get('name'), rec.get('province'),
                     (rec.get('city') or ''), _level_text(rec), rec.get('type', '')])
    if not rows:
        sys.stderr.write('没有符合条件的学校。\n')
        return 1
    _print_table(['slug', '名称', '省份', '城市', '层次', '类型'], rows)
    print('')
    print('共 %d 所' % len(rows))
    return 0


def cmd_search(args):
    q = ' '.join(args.keyword).strip()
    scored = _search_scored(q)
    if not scored:
        sys.stderr.write('未找到与「%s」相关的学校。\n' % q)
        return 1
    rows = []
    for score, kind, rec in scored[:args.limit]:
        rows.append([rec.get('slug'), rec.get('name'), rec.get('province'),
                     _level_text(rec), int(round(score))])
    _print_table(['slug', '名称', '省份', '层次', '得分'], rows)
    print('')
    print('共 %d 条匹配（显示前 %d 条）' % (len(scored), min(len(scored), args.limit)))
    return 0


def cmd_resolve(args):
    q = ' '.join(args.query).strip()
    hit = resolve_school(q)
    if hit:
        print('最佳匹配：')
        print('  slug : %s' % hit['slug'])
        print('  名称 : %s' % hit['name'])
        print('  省份 : %s %s' % (hit.get('province', ''), hit.get('city', '')))
        print('  层次 : %s' % _level_text(hit))
        print('  得分 : %s（%s）' % (hit['_score'], hit['_match']))
        return 0
    print('未找到「%s」，候选如下：' % q)
    cands = resolve_school_candidates(q, 3)
    if not cands:
        print('  （没有任何相似条目，请检查输入）')
        return 1
    for i, rec in enumerate(cands, 1):
        print('  %d. %s  %s  %s  （得分 %s，%s）'
              % (i, rec['slug'], rec['name'], rec.get('province', ''),
                 rec['_score'], rec['_match']))
    return 1


def cmd_validate(args):
    path = args.path or DATA_PATH
    fails = []
    warns = []
    try:
        with open(path, encoding='utf-8') as f:
            schools = json.load(f)
    except Exception as e:
        print('FAIL 无法读取或解析 JSON：%s（%s）' % (path, e))
        return 1
    if not isinstance(schools, list):
        print('FAIL JSON 顶层不是数组')
        return 1
    passed = 0

    def check(cond, msg):
        nonlocal passed
        if cond:
            passed += 1
        else:
            fails.append(msg)

    check(len(schools) > 0, '学校列表为空')
    seen_slug, seen_name = {}, {}
    required = set(FIELDS)
    for i, rec in enumerate(schools):
        tag = '第 %d 条(%s)' % (i + 1, rec.get('name', '?') if isinstance(rec, dict) else '?')
        if not isinstance(rec, dict):
            fails.append('%s：不是对象' % tag)
            continue
        for f in REQUIRED_FIELDS:
            check(bool(str(rec.get(f, '')).strip()), '%s：缺少必填字段 %s' % (tag, f))
        extra = set(rec) - required
        check(not extra, '%s：存在规范外字段 %s' % (tag, sorted(extra)))
        slug = str(rec.get('slug', ''))
        check(bool(re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*', slug)),
              '%s：slug 不合法（%s）' % (tag, slug))
        if slug in seen_slug:
            fails.append('%s：slug 重复（%s 与 %s）' % (tag, slug, seen_slug[slug]))
        else:
            seen_slug[slug] = rec.get('name', '?')
        name = str(rec.get('name', '')).strip()
        if name in seen_name:
            fails.append('%s：学校名重复（%s）' % (tag, name))
        else:
            seen_name[name] = slug
        prov = str(rec.get('province', '')).strip()
        check(prov in PROVINCES, '%s：省份不合法（%s）' % (tag, prov))
        city = rec.get('city', '')
        check(isinstance(city, str), '%s：city 应为字符串' % tag)
        lv = rec.get('level')
        check(isinstance(lv, list), '%s：level 应为数组' % tag)
        if isinstance(lv, list):
            bad = [x for x in lv if x not in LEVELS]
            check(not bad, '%s：level 含非法值 %s' % (tag, bad))
            check(len(lv) == len(set(map(str, lv))), '%s：level 有重复项' % tag)
        ty = rec.get('type')
        check(ty in TYPES, '%s：type 不合法（%s）' % (tag, ty))
        py = str(rec.get('pinyin', ''))
        check(bool(re.fullmatch(r'[a-z0-9]+', py)), '%s：pinyin 不合法（%s）' % (tag, py))
        ab = str(rec.get('abbr', ''))
        check(bool(re.fullmatch(r'[a-z0-9]*', ab)), '%s：abbr 不合法（%s）' % (tag, ab))
        al = rec.get('aliases')
        check(isinstance(al, list) and all(isinstance(x, str) and x.strip() for x in al),
              '%s：aliases 应为非空字符串数组' % tag)
        ne = rec.get('name_en', '')
        check(isinstance(ne, str), '%s：name_en 应为字符串' % tag)
        web = rec.get('website', '')
        check(isinstance(web, str) and (not web or web.startswith('http')),
              '%s：website 应为 http(s) 地址或空串' % tag)
        # 双一流/985/211 至少有别名
        if isinstance(lv, list) and set(lv) & {'985', '211', '双一流'}:
            check(isinstance(al, list) and len(al) >= 1, '%s：重点校缺少至少 1 个别名' % tag)

    # 覆盖率类提醒（不阻断）
    no_en = [r.get('name') for r in schools if isinstance(r, dict) and not r.get('name_en')]
    no_web = [r.get('name') for r in schools if isinstance(r, dict) and not r.get('website')]
    print('校验文件：%s' % path)
    print('条目数：%d；PASS：%d；FAIL：%d' % (len(schools), passed, len(fails)))
    if fails:
        print('')
        for msg in fails[:50]:
            print('  [FAIL] %s' % msg)
        if len(fails) > 50:
            print('  ...（共 %d 条，只显示前 50 条）' % len(fails))
    if no_en:
        print('提示：%d 条缺英文名（规范允许为空）' % len(no_en))
    if no_web:
        print('提示：%d 条缺官网（规范允许为空）' % len(no_web))
    print('')
    if fails:
        print('结果：FAIL（%d 项不合格）' % len(fails))
        return 1
    print('结果：PASS（全部检查通过）')
    return 0


def cmd_export_csv(args):
    schools = load_schools()
    out = args.out
    with open(out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(FIELDS)
        for rec in schools:
            row = []
            for k in FIELDS:
                v = rec.get(k, '')
                if isinstance(v, list):
                    v = '、'.join(str(x) for x in v)
                row.append(v)
            w.writerow(row)
    print('已导出 %d 条到 %s' % (len(schools), out))
    return 0


# ------------------------------------------------------------------ CLI
def build_parser():
    p = argparse.ArgumentParser(
        prog='schools_tool.py',
        description='全国本科院校身份库查询工具（数据：assets/schools/schools.json）',
    )
    sub = p.add_subparsers(dest='command', metavar='子命令')

    sp = sub.add_parser('stats', help='总数 / 分省 / 分层次 / 分类型统计')
    sp.set_defaults(func=cmd_stats)

    sp = sub.add_parser('list', help='按省份 / 层次 / 类型筛选')
    sp.add_argument('--province', help='省级行政区，如 示例省（可写 示例省省）')
    sp.add_argument('--level', choices=LEVELS, help='层次：985/211/双一流/部属/省属')
    sp.add_argument('--type', choices=TYPES, help='类型：综合/理工/师范/…')
    sp.set_defaults(func=cmd_list)

    sp = sub.add_parser('search', help='按名称/别名/拼音/首字母模糊搜索（前 10 条）')
    sp.add_argument('keyword', nargs='+', help='关键词，可多个词')
    sp.add_argument('--limit', type=int, default=10, help='显示条数，默认 10')
    sp.set_defaults(func=cmd_search)

    sp = sub.add_parser('resolve', help='把自由输入解析成唯一学校')
    sp.add_argument('query', nargs='+', help='学校名/简称/拼音/首字母')
    sp.set_defaults(func=cmd_resolve)

    sp = sub.add_parser('validate', help='校验数据文件（必填/唯一/省份/JSON）')
    sp.add_argument('--path', help='指定要校验的 JSON（默认 assets/schools/schools.json）')
    sp.set_defaults(func=cmd_validate)

    sp = sub.add_parser('export-csv', help='导出 CSV')
    sp.add_argument('--out', default='schools.csv', help='输出路径，默认 schools.csv')
    sp.set_defaults(func=cmd_export_csv)
    return p


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, 'func', None):
        parser.print_help()
        return 1
    return args.func(args) or 0


if __name__ == '__main__':
    sys.exit(main())

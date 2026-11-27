#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""extract_figures.py — 从论文 PDF 抽图（位图 + 矢量图）。

用法：
    python3 extract_figures.py --pdf x.pdf --out work/xxx/figures [--pages 1-17]
                              [--max-figures 20] [--min-width 600]

输出：
    figures/fig01.png …   抽出来的图（PNG，宽 ≥600px，已裁白边）
    figures.json          [{file,page,caption,bbox,kind,width,height,paper_no,method,dpi}]

策略：
  1) 先扫每页的「Fig. n.」图注（排版基准）；
  2) 把页内嵌位图 bbox 与矢量绘图 bbox（get_drawings）聚成图区簇；
  3) 按「图注上方最近的图形簇」分配，跨栏图自动合并（如首页跨两栏的 Fig.1）；
  4) 单张位图占比 ≥85% 且无旋转时直接抽原图（分辨率更高），否则按 300dpi 渲染；
  5) 裁掉白边；跳过校徽/头像等无图注的小图；每张 ≥600px 宽（不够则自适应升 dpi）。

依赖：PyMuPDF（pymupdf / fitz）+ Pillow。缺依赖时中文报错。
"""
from __future__ import annotations

import argparse
import datetime as _dt
import io
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from paper_extract import (  # noqa: E402
        CAP_FIG_RE,
        build_block_records,
        body_font_size,
        clean_inline,
        detect_columns,
        extract_page_lines,
        require_pymupdf,
    )
except Exception as exc:  # pragma: no cover
    raise SystemExit(
        "无法导入同目录的 paper_extract.py（%s）。\n"
        "请确认 paper_extract.py 与 extract_figures.py 在同一个 scripts/ 目录下。" % exc
    )

try:
    from PIL import Image, ImageChops  # noqa: F401
except Exception:
    raise SystemExit(
        "缺少 Pillow，无法处理抽出的图片。\n"
        "请用带 Pillow 的解释器运行（如 /opt/anaconda3/bin/python3），"
        "或安装：/opt/anaconda3/bin/python3 -m pip install pillow"
    )


# ---------------------------------------------------------------------------
# 图形簇检测
# ---------------------------------------------------------------------------

def _clip_rect(r, page_w, page_h, tol=5.0):
    """把 bbox 夹到页面内；超出页面太多的（剪裁路径 bug）直接丢弃。"""
    x0, y0, x1, y1 = r
    if x1 <= 0 or y1 <= 0 or x0 >= page_w or y0 >= page_h:
        return None
    if x0 < -tol or y0 < -tol or x1 > page_w + tol or y1 > page_h + tol:
        # 明显出界的绘图多半是剪裁路径，丢掉以免把整页并成一簇
        if x1 > page_w + tol or y1 > page_h + tol:
            return None
    x0 = max(0.0, x0)
    y0 = max(0.0, y0)
    x1 = min(page_w, x1)
    y1 = min(page_h, y1)
    if x1 - x0 < 0.6 or y1 - y0 < 0.6:
        return None
    return (x0, y0, x1, y1)


def collect_graphics(page, doc, body_size):
    """返回 (clusters, images)。

    clusters: [{'bbox':(x0,y0,x1,y1), 'items':[...], 'has_image':bool}]
    items 里 kind=image/draw。
    """
    page_w, page_h = page.rect.width, page.rect.height
    items = []

    # 1) 内嵌位图
    images = []
    try:
        infos = page.get_image_info(xrefs=True)
    except Exception:
        infos = []
    seen_xref_rect = set()
    for info in infos:
        xref = info.get("xref", 0)
        bbox = info.get("bbox")
        if not bbox:
            continue
        bx = _clip_rect(tuple(bbox), page_w, page_h)
        if not bx:
            continue
        w, h = bx[2] - bx[0], bx[3] - bx[1]
        area = w * h
        if area < 120:  # 校徽/项目图标一类的小图
            continue
        if area > page_w * page_h * 0.85:  # 整页扫描底图
            continue
        if w < 130 and 1.05 <= (h / max(w, 1e-6)) <= 1.55:
            # 作者头像（300×375 的竖图）
            continue
        transform = info.get("transform")
        images.append(
            {
                "xref": xref,
                "bbox": bx,
                "area": area,
                "width": info.get("width", 0),
                "height": info.get("height", 0),
                "transform": transform,
            }
        )
        items.append({"kind": "image", "bbox": bx, "area": area})

    # 2) 矢量绘图
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    for d in drawings:
        r = d.get("rect")
        if r is None:
            continue
        bx = _clip_rect((r.x0, r.y0, r.x1, r.y1), page_w, page_h)
        if not bx:
            continue
        area = (bx[2] - bx[0]) * (bx[3] - bx[1])
        if area < 0.5:
            continue
        items.append({"kind": "draw", "bbox": bx, "area": area})

    # 3) 用 bbox 邻近度聚簇
    items.sort(key=lambda it: (it["bbox"][1], it["bbox"][0]))
    gap = 5.0
    clusters = []
    for it in items:
        x0, y0, x1, y1 = it["bbox"]
        cand = None
        for cl in clusters:
            cx0, cy0, cx1, cy1 = cl["bbox"]
            if x0 <= cx1 + gap and x1 >= cx0 - gap and y0 <= cy1 + gap and y1 >= cy0 - gap:
                cand = cl
                break
        if cand is None:
            clusters.append({"bbox": [x0, y0, x1, y1], "items": [it], "has_image": it["kind"] == "image"})
            continue
        cand["items"].append(it)
        cand["bbox"][0] = min(cand["bbox"][0], x0)
        cand["bbox"][1] = min(cand["bbox"][1], y0)
        cand["bbox"][2] = max(cand["bbox"][2], x1)
        cand["bbox"][3] = max(cand["bbox"][3], y1)
        if it["kind"] == "image":
            cand["has_image"] = True
        # 簇变大后可能吞并其他簇：反复检查
        changed = True
        while changed:
            changed = False
            for other in list(clusters):
                if other is cand:
                    continue
                ox0, oy0, ox1, oy1 = other["bbox"]
                if (
                    ox0 <= cand["bbox"][2] + gap
                    and ox1 >= cand["bbox"][0] - gap
                    and oy0 <= cand["bbox"][3] + gap
                    and oy1 >= cand["bbox"][1] - gap
                ):
                    cand["items"].extend(other["items"])
                    cand["bbox"][0] = min(cand["bbox"][0], ox0)
                    cand["bbox"][1] = min(cand["bbox"][1], oy0)
                    cand["bbox"][2] = max(cand["bbox"][2], ox1)
                    cand["bbox"][3] = max(cand["bbox"][3], oy1)
                    cand["has_image"] = cand["has_image"] or other.get("has_image", False)
                    clusters.remove(other)
                    changed = True
    return clusters, images


def cluster_bbox(cl):
    return tuple(round(v, 2) for v in cl["bbox"])


# ---------------------------------------------------------------------------
# 图注 → 图区分配
# ---------------------------------------------------------------------------

def collect_tab_captions(recs, page_no):
    """表注（IEEE 表格题注在表格上方），用于排除表格区域，避免把表并进图。"""
    tabs = []
    for r in recs:
        flat = clean_inline(r["text"])
        if r.get("kind") != "caption_tab":
            continue
        tabs.append({"page": page_no, "caption": flat, "bbox": r["bbox"]})
    return tabs


def collect_fig_captions(recs, page_no):
    caps = []
    for r in recs:
        flat = clean_inline(r["text"])
        m = CAP_FIG_RE.match(flat)
        if not m:
            continue
        # 跳过正文里 "Fig. 3(a) and (b) present..." 这类引用句（图注要求编号后是句点/冒号）
        caps.append(
            {
                "page": page_no,
                "no": "Fig. " + m.group(1),
                "caption": flat,
                "bbox": r["bbox"],
            }
        )
    return caps


def _column_bounds_for_x(x, columns):
    if not columns:
        return (0.0, 1e9)
    best = (0.0, 1e9)
    bestd = 1e9
    for (x0, x1) in columns:
        d = 0 if x0 <= x <= x1 else min(abs(x - x0), abs(x - x1))
        if d < bestd:
            best, bestd = (x0, x1), d
    return best


def _table_owned(cl, table_caps, columns):
    """判断簇是否属于某个表格（表注在上方 + 同栏 + 垂直窗口内）。

    表注文字往往比表格窄（居中），所以按表注所在栏的宽度算横向重叠。
    """
    if not table_caps:
        return False
    x0, y0, x1, _y1 = cl["bbox"]
    for tc in table_caps:
        tx0, ty0, tx1, ty1 = tc["bbox"]
        if y0 - ty1 < -2 or y0 - ty1 > 430:
            continue
        cx0, cx1 = _column_bounds_for_x((tx0 + tx1) / 2.0, columns)
        ex0, ex1 = min(tx0, cx0), max(tx1, cx1)
        if min(x1, ex1) - max(x0, ex0) > 15:
            return True
    return False


def _col_of_x(x, columns):
    if not columns:
        return 0
    best, bestd = 0, 1e9
    for i, (x0, x1) in enumerate(columns):
        d = 0 if x0 <= x <= x1 else min(abs(x - x0), abs(x - x1))
        if d < bestd:
            best, bestd = i, d
    return best


def _caption_col(cap, columns):
    return _col_of_x((cap["bbox"][0] + cap["bbox"][2]) / 2.0, columns)


def assign_clusters(caps, clusters, search_gap=50.0, stack_gap=45.0, overlap_tol=8.0,
                    table_caps=None, columns=None):
    """把图形簇分配给图注。

    规则：
      1) 图注在图下方：簇底边贴着图注上沿（允许少量重叠）→ 候选；
      2) 每个簇归「gap 最小」的图注，避免左右相邻的两张图被并成一张；
      3) 再向上「堆叠合并」：同一栏里正上方、紧贴的面板归下面那张图的图注
         （如 Fig. 7 有两个上下子图，上面那个离 Fig. 6 的图注更近，要抢回来）；
      4) 兜底：若某图注没分到簇，再找它正下方的簇（图注在图上方的情况）。
    """
    groups = {ci: [] for ci in range(len(caps))}
    best_for_cluster = {}
    for ki, cl in enumerate(clusters):
        by1 = cl["bbox"][3]
        cands = []
        for ci, cap in enumerate(caps):
            gap = cap["bbox"][1] - by1
            if -overlap_tol <= gap <= search_gap:
                # 允许轻微重叠，但重叠的要重罚，避免抢到下一张图的图注
                score = gap if gap >= 0 else gap + 40.0
                cands.append((score, ci))
        if cands:
            cands.sort()
            best_for_cluster[ki] = cands[0][1]
    for ki, ci in best_for_cluster.items():
        groups[ci].append(clusters[ki])
    owner = {}
    for ci, g in groups.items():
        for cl in g:
            owner[id(cl)] = ci

    # 向上的堆叠面板合并：从页面最下方的图注开始处理
    cap_order = sorted(range(len(caps)), key=lambda ci: caps[ci]["bbox"][1], reverse=True)
    for ci in cap_order:
        grp = groups[ci]
        if not grp:
            continue
        changed = True
        while changed:
            changed = False
            top = min(c["bbox"][1] for c in grp)
            gx0 = min(c["bbox"][0] for c in grp)
            gx1 = max(c["bbox"][2] for c in grp)
            for cl in clusters:
                if any(cl is x for x in grp):
                    continue
                gap = top - cl["bbox"][3]
                if not (0 <= gap <= stack_gap):
                    continue
                ox = min(gx1, cl["bbox"][2]) - max(gx0, cl["bbox"][0])
                if ox <= 20:
                    continue
                # 跨栏宽图（如整页顶部的通栏图）不能当成同一栏的堆叠面板
                if cl["bbox"][0] < gx0 - 60 or cl["bbox"][2] > gx1 + 60:
                    continue
                if _table_owned(cl, table_caps, columns):
                    continue
                own = owner.get(id(cl))
                if own is not None:
                    own_col = _caption_col(caps[own], columns)
                    cl_col = _col_of_x((cl["bbox"][0] + cl["bbox"][2]) / 2.0, columns)
                    # 簇和它原来的图注在同一栏 -> 它是那张图的主体面板，不抢
                    if own_col == cl_col:
                        continue
                for cj, g2 in groups.items():
                    if cj != ci:
                        g2[:] = [x for x in g2 if x is not cl]
                grp.append(cl)
                owner[id(cl)] = ci
                changed = True

    # 兜底：图注在图上方（少数期刊）
    used = set()
    for g in groups.values():
        for cl in g:
            used.add(id(cl))
    for ci, cap in enumerate(caps):
        if groups[ci]:
            continue
        cy1 = cap["bbox"][3]
        for cl in clusters:
            if id(cl) in used:
                continue
            gap = cl["bbox"][1] - cy1
            if 0 <= gap <= search_gap:
                groups[ci].append(cl)
                used.add(id(cl))
                break
    return groups


# ---------------------------------------------------------------------------
# 裁剪 / 渲染 / 抽位图
# ---------------------------------------------------------------------------

def _img_white_bg(img):
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        bg.alpha_composite(img)
        return bg.convert("RGB")
    return img.convert("RGB")


def trim_white(img, pad=6, threshold=246):
    """裁掉四周纯白边（保留浅色底/网格），留 pad 像素。"""
    gray = img.convert("L")
    mask = gray.point(lambda v: 0 if v >= threshold else 255)
    bbox = mask.getbbox()
    if not bbox:
        return img
    x0, y0, x1, y1 = bbox
    x0 = max(0, x0 - pad)
    y0 = max(0, y0 - pad)
    x1 = min(img.width, x1 + pad)
    y1 = min(img.height, y1 + pad)
    if x1 - x0 < 8 or y1 - y0 < 8:
        return img
    return img.crop((x0, y0, x1, y1))


def _pick_direct_image(doc, page, images, union_bbox, drawings_inside):
    """判断能否直接抽内嵌位图（分辨率比渲染高）。"""
    if len(images) != 1:
        return None
    img = images[0]
    bx = img["bbox"]
    area = (bx[2] - bx[0]) * (bx[3] - bx[1])
    ua = (union_bbox[2] - union_bbox[0]) * (union_bbox[3] - union_bbox[1])
    if ua <= 0 or area / ua < 0.85:
        return None
    if drawings_inside > 2:  # 图里有大量矢量标注/曲线时，渲染更完整
        return None
    tr = img.get("transform")
    if tr:
        try:
            a, b, c, d = tr[0], tr[1], tr[2], tr[3]
            if abs(b) > 1e-3 or abs(c) > 1e-3 or a <= 0 or d <= 0:
                return None  # 旋转/翻转的位图，直接抽会方向不对
        except Exception:
            return None
    if not img.get("xref"):
        return None
    try:
        raw = doc.extract_image(img["xref"])
    except Exception:
        return None
    if not raw or not raw.get("image"):
        return None
    try:
        pil = Image.open(io.BytesIO(raw["image"]))
        pil = _img_white_bg(pil)
        return pil
    except Exception:
        return None


def render_region(page, rect, dpi, min_width):
    """渲染图区；若宽度不足 min_width 则自动升 dpi。"""
    x0, y0, x1, y1 = rect
    w_pt = max(1.0, x1 - x0)
    need_dpi = int(min_width / (w_pt / 72.0)) + 1
    use_dpi = min(max(dpi, need_dpi), 600)
    pix = page.get_pixmap(clip=page.rect.__class__(x0, y0, x1, y1), dpi=use_dpi, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return img, use_dpi


def expand_crop(union_bbox, cap_bbox, columns, page_rect, pad_top=8.0, pad_bottom=6.0,
                text_bounds=None):
    """把图形簇 bbox 扩到图区：横向按栏宽，纵向留白但避开图注与页眉。"""
    x0, y0, x1, y1 = union_bbox
    # 覆盖了哪些栏
    covered = []
    for ci, (cx0, cx1) in enumerate(columns):
        overlap = min(x1, cx1) - max(x0, cx0)
        if overlap > 25:
            covered.append(ci)
    if len(covered) >= 2:
        crop_x0 = min(columns[c][0] for c in covered)
        crop_x1 = max(columns[c][1] for c in covered)
    elif len(covered) == 1:
        ci = covered[0]
        crop_x0, crop_x1 = columns[ci][0], columns[ci][1]
    else:
        # 整栏都是图、正文列检测不到图所在栏时，用页面文字区宽度兜底
        tb = text_bounds or (0.0, page_rect.width)
        crop_x0 = max(tb[0], x0 - 3.0)
        crop_x1 = min(tb[1], x1 + 3.0)
    # 纵向：上留 pad，下不进入图注
    crop_y0 = max(y0 - pad_top, 46.0)
    if cap_bbox:
        crop_y1 = min(y1 + pad_bottom, cap_bbox[1] - 2.5)
    else:
        crop_y1 = min(y1 + pad_bottom, page_rect.height - 40)
    crop_y0 = max(crop_y0, 0.0)
    crop_y1 = min(crop_y1, page_rect.height)
    if crop_x1 - crop_x0 < 30 or crop_y1 - crop_y0 < 20:
        return None
    return (crop_x0, crop_y0, crop_x1, crop_y1)


def _row_ink_profile(image):
    """逐行统计非白像素比例（用于找子图之间的分隔带）。"""
    g = image.convert("L")
    w, h = g.size
    if h <= 0 or w <= 0:
        return []
    px = g.load()
    step = max(1, w // 500)
    rows = []
    for y in range(h):
        cnt = 0
        n = 0
        for x in range(0, w, step):
            if px[x, y] < 235:
                cnt += 1
            n += 1
        rows.append(cnt / max(1, n))
    return rows


def analyze_panels(page, crop, split_dpi=150):
    """识别上下堆叠的 (a)/(b)/(c) 子图。

    返回 {"splits": [pt], "bounds": [(y0,y1)...], "labels": [(y0,y1) or None, ...]}；
    识别失败返回 {"splits": [], ...}。

    做法：把图区按 150dpi 渲染，逐行统计墨水比例；
    - 「面板段」= 连续有墨水、高度 >= 图区 12% 的段；
    - 两个面板之间的空白带里取最宽的一条切分（子图注在面板下方）；
    - 「标签段」= 切分点之前最后一个文字段（"(a) Sum-rate." 这类子图注），
      供 --strip-labels 时裁掉（PPT 图注会再写一次图号）。
    """
    fitz = require_pymupdf()
    empty = {"splits": [], "bounds": [], "labels": []}
    x0, y0, x1, y1 = crop
    if x1 - x0 < 40 or y1 - y0 < 60:
        return empty
    try:
        pix = page.get_pixmap(clip=fitz.Rect(x0, y0, x1, y1), dpi=split_dpi, alpha=False)
        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    except Exception:
        return empty
    rows = _row_ink_profile(image)
    h = len(rows)
    if h < 100:
        return empty
    white = [r < 0.004 for r in rows]
    segs = []
    i = 0
    while i < h:
        if not white[i]:
            j = i
            while j < h and not white[j]:
                j += 1
            segs.append((i, j))
            i = j
        else:
            i += 1
    if len(segs) < 3:
        return empty
    min_panel = max(8, int(h * 0.12))
    panels = [k for k, (a, b) in enumerate(segs) if (b - a) >= min_panel]
    if len(panels) < 2 or len(panels) > 3:
        return empty
    splits_px = []
    label_px = []
    for k in range(len(panels) - 1):
        p1, p2 = panels[k], panels[k + 1]
        # 两面板之间的空白带里，取最宽的一条（子图注与下一面板之间）
        lo = segs[p1][1]
        hi = segs[p2][0]
        best_len = 0
        best_mid = None
        y = lo
        while y < hi:
            if white[y]:
                z = y
                while z < hi and white[z]:
                    z += 1
                if (z - y) > best_len:
                    best_len = z - y
                    best_mid = (y + z) // 2
                y = z
            else:
                y += 1
        if best_mid is None:
            between = segs[p1 + 1 : p2]
            best_mid = between[-1][1] if between else lo
        splits_px.append(best_mid)
        # 切分点之前最后一个（较矮的）文字段 = 本面板的子图注
        lab = None
        for (a, b) in segs[p1 + 1 : p2]:
            if b <= best_mid and (b - a) <= min_panel:
                lab = (a, b)
        label_px.append(lab)
    # 最后一个面板的子图注：面板之后最后一个小段
    last_lab = None
    for (a, b) in segs[panels[-1] + 1 :]:
        if (b - a) <= min_panel:
            last_lab = (a, b)
    label_px.append(last_lab)

    splits_pt = [y0 + sp * (y1 - y0) / float(h) for sp in splits_px]
    bounds = [y0] + splits_pt + [y1]
    if any(b - a < 30 for a, b in zip(bounds[:-1], bounds[1:])):
        return empty
    total_h = y1 - y0
    total_w = x1 - x0
    ratios = [(b - a) / total_h for a, b in zip(bounds[:-1], bounds[1:])]
    if len(ratios) == 2:
        if not all(0.30 <= r <= 0.70 for r in ratios):
            return empty
    elif len(ratios) == 3:
        if not all(0.22 <= r <= 0.45 for r in ratios):
            return empty
    else:
        return empty
    for a, b in zip(bounds[:-1], bounds[1:]):
        wh = total_w / max(1e-6, (b - a))
        if not (0.75 <= wh <= 3.2):
            return empty
    labels = []
    for (a, b) in label_px:
        if a is None:
            labels.append(None)
        else:
            labels.append((y0 + a * total_h / float(h), y0 + b * total_h / float(h)))
    return {"splits": splits_pt, "bounds": list(zip(bounds[:-1], bounds[1:])), "labels": labels}


def detect_panel_splits(page, crop, split_dpi=150):
    """兼容旧接口：只返回切分点。"""
    return analyze_panels(page, crop, split_dpi)["splits"]


def classify_figure_kind(caption):
    t = (caption or "").lower()
    framework_kw = ["framework", "architecture", "structure", "overview", "block diagram",
                    "flowchart", "pipeline", "schematic", "process of", "procedure", "framework"]
    result_kw = ["sum-rate", "sum rate", "snr", "nmse", "pareto", "trajectory", "convergence",
                 "ablation", "comparison", "versus", "vs.", "performance", "results", "achieved",
                 "error", "trade-off", "tradeoff", "throughput", "accuracy", "robustness",
                 "evaluation", "coverage", "kpi", "denoising", "loss"]
    scene_kw = ["system model", "scenario", "deployment", "systems", "setup", "illustration",
                "low-altitude", "network", "process"]
    if any(k in t for k in framework_kw):
        return "framework"
    if any(k in t for k in result_kw):
        return "result"
    if any(k in t for k in scene_kw):
        return "scene"
    return "concept"


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def extract_figures(pdf_path, out_dir, pages_spec=None, max_figures=20, min_width=600,
                    dpi=300, debug=False, split_panels=False, strip_labels=False):
    fitz = require_pymupdf()
    doc = fitz.open(pdf_path)
    page_count = doc.page_count
    if pages_spec:
        pages = []
        for part in str(pages_spec).split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                a = int(a) if a.strip() else 1
                b = int(b) if b.strip() else page_count
                pages.extend(range(max(1, a), min(page_count, b) + 1))
            else:
                p = int(part)
                if 1 <= p <= page_count:
                    pages.append(p)
        pages = sorted(set(pages))
    else:
        pages = list(range(1, page_count + 1))

    # 第一遍：估算正文字号（全页，稳定）
    all_lines = {}
    for pno in range(1, page_count + 1):
        lines, _, _ = extract_page_lines(doc[pno - 1])
        all_lines[pno] = lines
    body_size = body_font_size([all_lines[p] for p in all_lines])

    fig_dir = out_dir  # figNN.png 直接放在 --out 目录下（与 figures.json 同级）
    os.makedirs(fig_dir, exist_ok=True)

    records = []
    skipped = []
    warnings = []
    saved = 0
    for pno in pages:
        page = doc[pno - 1]
        lines = all_lines[pno]
        recs, _images, _dc, _lines = build_block_records(page, body_size)
        caps = collect_fig_captions(recs, pno)
        if not caps:
            continue
        good_lines = [ln for ln in lines if ln["text"].strip()]
        text_bounds = (
            min((ln["bbox"][0] for ln in good_lines), default=40.0),
            max((ln["bbox"][2] for ln in good_lines), default=page.rect.width - 40.0),
        )
        columns = detect_columns(lines, page.rect, body_size)
        clusters, images = collect_graphics(page, doc, body_size)
        # 表格区域（题注在下方的簇）不参与图注分配
        tab_caps = collect_tab_captions(recs, pno)
        if tab_caps:
            table_cl = set()
            for tc in tab_caps:
                for ki, cl in enumerate(clusters):
                    gap = cl["bbox"][1] - tc["bbox"][3]
                    if -3.0 <= gap <= 40.0:
                        table_cl.add(ki)
            if table_cl:
                clusters = [cl for ki, cl in enumerate(clusters) if ki not in table_cl]
        assigned = assign_clusters(caps, clusters, table_caps=tab_caps, columns=columns)
        if debug:
            print("  p%-3d caps=%d clusters=%d -> %s" % (
                pno, len(caps), len(clusters),
                [(caps[i]["no"], len(v)) for i, v in assigned.items()]), file=sys.stderr)
        for ci, cap in enumerate(caps):
            clos = assigned.get(ci) or []
            if not clos:
                skipped.append({"page": pno, "paper_no": cap["no"], "caption": cap["caption"],
                                "reason": "没找到对应的图形区域（可能是纯文本/表格图或扫描件）"})
                continue
            # 合并簇 bbox
            ux0 = min(c["bbox"][0] for c in clos)
            uy0 = min(c["bbox"][1] for c in clos)
            ux1 = max(c["bbox"][2] for c in clos)
            uy1 = max(c["bbox"][3] for c in clos)
            union = (ux0, uy0, ux1, uy1)
            crop = expand_crop(union, cap["bbox"], columns, page.rect, text_bounds=text_bounds)
            if not crop:
                skipped.append({"page": pno, "paper_no": cap["no"], "caption": cap["caption"],
                                "reason": "图区过小（bbox 异常）"})
                continue
            # 统计图区内矢量绘图数（用于判断能否直接抽位图）
            drawings_inside = 0
            for c in clos:
                for it in c["items"]:
                    if it["kind"] == "draw":
                        b = it["bbox"]
                        if b[0] >= crop[0] - 2 and b[2] <= crop[2] + 2 and b[1] >= crop[1] - 2 and b[3] <= crop[3] + 2:
                            drawings_inside += 1
            img = None
            method = ""
            used_dpi = 0
            direct = _pick_direct_image(doc, page, [i for i in images if _bbox_contains(crop, i["bbox"])], union, drawings_inside)
            if direct is not None and direct.width >= min_width:
                img = trim_white(direct)
                method = "bitmap"
            if img is None or img.width < min_width:
                rendered, used_dpi = render_region(page, crop, dpi, min_width)
                rendered = trim_white(rendered)
                if img is None or rendered.width * rendered.height > img.width * img.height:
                    img = rendered
                    method = "render"
            if img is None or img.width < 200 or img.height < 80:
                skipped.append({"page": pno, "paper_no": cap["no"], "caption": cap["caption"],
                                "reason": "抽出图片过小（宽<200px），可能是校徽/公式截图"})
                continue
            if img.width < min_width:
                warnings.append("%s（第 %d 页）宽度只有 %dpx，低于 %dpx" % (cap["no"], pno, img.width, min_width))
            saved += 1
            fname = "fig%02d.png" % saved
            fpath = os.path.join(fig_dir, fname)
            img.save(fpath, "PNG")
            rec = {
                "file": fname,
                "page": pno,
                "caption": cap["caption"],
                "bbox": [round(v, 1) for v in crop],
                "kind": classify_figure_kind(cap["caption"]),
                "width": img.width,
                "height": img.height,
                "paper_no": cap["no"],
                "method": method,
                "dpi": used_dpi if method == "render" else 0,
            }
            records.append(rec)
            # ---- 可选：把 (a)/(b)/(c) 堆叠子图拆成单张 ----
            if split_panels:
                analysis = analyze_panels(page, crop)
                if analysis["splits"]:
                    letters = "abcdef"
                    panel_bounds = analysis["bounds"]
                    panel_labels = analysis["labels"]
                    panel_files = []
                    for pi_idx, (py0, py1) in enumerate(panel_bounds):
                        # --strip-labels：去掉子图里的 "(a) Sum-rate." 一行（PPT 图注会再写一次图号）
                        if strip_labels and pi_idx < len(panel_labels) and panel_labels[pi_idx]:
                            lby0, lby1 = panel_labels[pi_idx]
                            if lby1 - lby0 <= 26 and py0 + 30 < lby0 and lby1 < py1 - 2:
                                if (lby0 + lby1) / 2.0 > (py0 + py1) / 2.0:
                                    py1 = lby0 - 1.5
                                else:
                                    py0 = lby1 + 1.5
                        if py1 - py0 < 40:
                            continue
                        pimg, _pdpi = render_region(page, (crop[0], py0, crop[2], py1), dpi, min_width)
                        pimg = trim_white(pimg)
                        if pimg.width < 200 or pimg.height < 60:
                            continue
                        pfname = "%s%s.png" % (os.path.splitext(fname)[0], letters[pi_idx])
                        pimg.save(os.path.join(fig_dir, pfname), "PNG")
                        panel_files.append(pfname)
                        records.append(
                            {
                                "file": pfname,
                                "page": pno,
                                "caption": cap["caption"],
                                "bbox": [round(crop[0], 1), round(py0, 1), round(crop[2], 1), round(py1, 1)],
                                "kind": classify_figure_kind(cap["caption"]),
                                "width": pimg.width,
                                "height": pimg.height,
                                "paper_no": "%s(%s)" % (cap["no"], letters[pi_idx]),
                                "method": "render-split",
                                "dpi": _pdpi,
                                "panel": letters[pi_idx],
                                "parent_file": fname,
                            }
                        )
                    if panel_files:
                        rec["panels"] = panel_files
                        rec["panel_count"] = len(panel_files)
            if saved >= max_figures:
                break
        if saved >= max_figures:
            warnings.append("已达到 --max-figures=%d 上限，后续图未抽" % max_figures)
            break
    doc.close()

    result = {
        "source": {
            "path": os.path.abspath(pdf_path),
            "file": os.path.basename(pdf_path),
            "pages_total": page_count,
            "pages_scanned": pages,
            "extracted_at": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "min_width": min_width,
            "dpi": dpi,
            "split_panels": split_panels,
            "strip_labels": strip_labels,
        },
        "count": len(records),
        "figures": records,
        "skipped": skipped,
        "warnings": warnings,
    }
    json_path = os.path.join(out_dir, "figures.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=1)
        f.write("\n")
    return result, json_path


def _bbox_contains(outer, inner, tol=1.5):
    return (
        inner[0] >= outer[0] - tol
        and inner[1] >= outer[1] - tol
        and inner[2] <= outer[2] + tol
        and inner[3] <= outer[3] + tol
    )


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="从论文 PDF 抽图：figNN.png + figures.json",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  /opt/anaconda3/bin/python3 extract_figures.py --pdf paper.pdf --out work/a/figures\n"
            "  /opt/anaconda3/bin/python3 extract_figures.py --pdf paper.pdf --out work/a/figures --pages 1-8\n"
        ),
    )
    ap.add_argument("--pdf", required=True, help="论文 PDF 路径")
    ap.add_argument("--out", required=True, help="输出目录（生成 figures/ 与 figures.json）")
    ap.add_argument("--pages", default=None, help="只抽这些页，1 起，如 1-17 或 4,7,12")
    ap.add_argument("--max-figures", type=int, default=20, help="最多抽多少张（默认 20）")
    ap.add_argument("--min-width", type=int, default=600, help="每张图最小宽度 px（默认 600）")
    ap.add_argument("--dpi", type=int, default=300, help="渲染 dpi（默认 300；宽度不够会自动升）")
    ap.add_argument("--split-panels", action="store_true",
                    help="识别上下堆叠的 (a)/(b) 子图，额外输出 fig03a.png / fig03b.png 等单图")
    ap.add_argument("--strip-labels", action="store_true",
                    help="配合 --split-panels：裁掉子图里的 «(a)/(b)» 子图注（默认保留）")
    ap.add_argument("--debug", action="store_true", help="打印每页图注/图区分配过程")
    args = ap.parse_args(argv)

    if not os.path.exists(args.pdf):
        raise SystemExit("找不到输入文件：%s" % args.pdf)
    os.makedirs(args.out, exist_ok=True)
    result, json_path = extract_figures(
        args.pdf, args.out, args.pages, args.max_figures, args.min_width, args.dpi,
        args.debug, args.split_panels, args.strip_labels
    )
    print("=" * 84)
    print("extract_figures 完成：%s" % os.path.basename(args.pdf))
    print("  输出：%s" % os.path.abspath(args.out))
    full_n = sum(1 for r in result["figures"] if "panel" not in r)
    panel_n = result["count"] - full_n
    panel_note = "（含子图 %d 张）" % panel_n if panel_n else ""
    print("  抽出 %d 张图%s（figures.json: %s）" % (result["count"], panel_note, os.path.abspath(json_path)))
    no_width = 0
    for rec in result["figures"]:
        if rec["width"] < args.min_width:
            no_width += 1
        print("  %-10s p%-3d %-9s %4dx%-4d %-7s %s" % (
            rec["file"], rec["page"], rec["paper_no"], rec["width"], rec["height"],
            rec["method"], rec["caption"][:64]))
    if result["skipped"]:
        print("  跳过 %d 处：" % len(result["skipped"]))
        for sk in result["skipped"][:10]:
            print("    - p%d %s：%s" % (sk["page"], sk["paper_no"], sk["reason"]))
    if result["warnings"]:
        print("  告警：")
        for w in result["warnings"]:
            print("    - " + w)
    print("=" * 84)
    return 0


if __name__ == "__main__":
    sys.exit(main())

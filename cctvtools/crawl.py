# -*- coding: utf-8 -*-
"""
CCTV 网页成果表抓取（统一入口）

把原先三个散装脚本的流程合并成一条带闸门的流水线：

    _crawl_4126_final.py   列表式：管段状况评估表（13 行/页，按管段编号去重翻页）
    _crawl_4127.py         卡片式：管段检测与评估成果表（1 管段/页）
    _debug_flow.py         Playwright 侦察：登录 → 业务处理 → KHD → 报告列表 → 点委托

工作流阶段（每个阶段都有产物和校验闸门）：

    S1 登录      POST /loginAction（密码 base64，验证码后端不校验）
    S2 定位报告  GET /XiangMutaizhangJson?BGBH=... → report_id / parentid / fid
    S3 抓 4126   XiangMuXianShi.jsp?...&currentpage=N（GBK，21 列，13 行/页）
    S4 抓 4127   XiangMuXianShi.jsp?...&currentpage=N（GBK，卡片式，1 管段/页）
    S5 落盘      xxx-{报告编号}(网页爬取)_{YYYYMMDD_HHMM}.xlsx
    S6 交叉校验 4126 与 4127 管段集合比对（含反向管段识别）+ 清单落盘
    S7 交付      manifest.json（页数/行数/停页原因/校验结论/sha256）

用法：
    python cctv_workflow.py --bgbh 26KHD00018 all          # 全流程
    python cctv_workflow.py --bgbh 26KHD00018 locate       # 只定位（打印 rid/parentid）
    python cctv_workflow.py --bgbh 26KHD00018 crawl 4126   # 只抓一张表
    python cctv_workflow.py --bgbh 26KHD00018 verify       # 只校验现有最新两个 xlsx
    python cctv_workflow.py selftest                       # 离线自检（不联网）

依赖：openpyxl。Playwright 仅在 --browser-fallback 时使用（旧路径兜底）。
"""
from __future__ import annotations

import argparse
import base64
import datetime
import hashlib
import html
import http.cookiejar
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

VERSION = '1.0.0'


from .common import *  # noqa: F401,F403
from .common import _data_table, _cell_text

# ==================== S3 抓 4126 ====================
def parse_4126(html_text):
    """
    返回 (rows, meta)。meta 分别统计三类被跳过的 <tr>，便于区分「正常表头」与「异常」：
      header_rows      td 数 11/12 = 表头行（正常，每页都有）
      placeholder_rows 21 列但管段编号是 '/' 的空预留行（26KHD00017 表尾有 4 行）
      malformed_rows   其余 td 数异常行（页面结构漂移的信号）
    """
    body = _data_table(html_text)
    meta = {'header_rows': 0, 'placeholder_rows': 0, 'malformed_rows': 0}
    if body is None:
        return [], meta
    rows = []
    for tr in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S):
        tds = re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S)
        if len(tds) != COLS_4126:
            if len(tds) in (11, 12):
                meta['header_rows'] += 1
            else:
                meta['malformed_rows'] += 1
            continue
        cells = [_cell_text(x) for x in tds]
        if not cells or not cells[0].isdigit():
            meta['malformed_rows'] += 1
            continue
        if _seg_key(cells[1]) in ('', '/', '—', '-'):
            meta['placeholder_rows'] += 1
            continue
        rows.append(cells)
    return rows, meta


def crawl_4126(sess, report, max_pages=80, sleep=0.0, verbose=True):
    """
    列表式翻页。停页条件（按优先级）：
      1) 页面没有 data-table / 无 21 列数据行  → empty（实测越界页返回空壳）
      2) 整页管段编号都已见过                → repeat（防御服务端重复吐最后一页）
      3) 到 max_pages                        → cap
    按「规范化管段编号」去重（去空白），避免个别页重复导致整行重复。
    """
    rows, seen, page = [], set(), 0
    agg = {'header_rows': 0, 'placeholder_rows': 0, 'malformed_rows': 0}
    dup_total = 0
    stop = 'cap'
    while page < max_pages:
        text = sess.get(page_url(sess, '4126', report, page))
        page_rows, meta = parse_4126(text)
        for k in agg:
            agg[k] += meta[k]
        if not page_rows:
            stop = 'empty'
            break
        new = []
        for r in page_rows:
            k = _seg_key(r[1])
            if k in seen:
                dup_total += 1
                continue
            seen.add(k)
            new.append(r)
        if not new:
            stop = 'repeat'
            break
        rows.extend(new)
        if verbose and (page % 5 == 0 or page == 0):
            print('    [S3] 4126 第%d页 +%d 行（累计 %d）' % (page, len(new), len(rows)))
        page += 1
        if sleep:
            time.sleep(sleep)
    stats = {'pages': page, 'rows': len(rows), 'duplicate_rows': dup_total,
             'stop_reason': stop, 'unique_segments': len(seen)}
    stats.update(agg)
    return rows, stats


# ==================== S4 抓 4127 ====================
DEFECT_LABELS = ('长度(m)', '缺陷名称代码', '分值', '等级', '管道内部状况描述', '图片')


def parse_4127(html_text):
    """
    解析一页(一个管段)的卡片。

    头部字段是「标签 → 右邻值」的平铺 <td>，但**缺陷明细是纵向 4 槽位**：
        <tr> 长度(m)[rowspan=4] | 1m | 缺陷名称代码[rowspan=4] | (CJ)沉积 | 分值[rowspan=4] | 2 </tr>
        <tr> 1.2m | (BX)变形 | 0.5 </tr>
        <tr> / | / | / </tr>
        <tr> / | / | / </tr>
        <tr> 等级[rowspan=4] | 2 | 管道内部状况描述[rowspan=4] | ... | 图片[rowspan=4] | 图1 </tr>
    所以只扫「标签的右邻值」会只拿到第 1 个槽位（旧 `_crawl_4127.py` 就漏了第 2 个缺陷）。
    这里按「标签行 + 紧随的值行」成块读取，值行按列序对应标签，保留 4 个槽位的位置。
    """
    body = _data_table(html_text)
    if body is None:
        return None
    rows = [[_cell_text(td) for td in re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S)]
            for tr in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S)]
    flat = [c for r in rows for c in r]

    def val(label, default='/'):
        for i, c in enumerate(flat):
            if c == label and i + 1 < len(flat):
                return flat[i + 1] or '/'
        return default

    start = val('起始井号', '')
    if not start or start == '/':
        return None                      # 越界页：序号 0、井号空

    # ---- 缺陷块：标签行 + 其后的值行 ----
    slots = {lbl: [] for lbl in DEFECT_LABELS}
    cur = None
    for r in rows:
        labels = [c for c in r if c in DEFECT_LABELS]
        if labels:
            vals = [c for c in r if c not in DEFECT_LABELS]
            cur = labels if len(vals) == len(labels) else None
            if cur:
                for lbl, v in zip(cur, vals):
                    slots[lbl].append(v)
            continue
        if cur and len(r) == len(cur):
            for lbl, v in zip(cur, r):
                slots[lbl].append(v)
        else:
            cur = None                   # 备注/检测地点等行，结束当前块

    rec = {
        '序号': cells0(rows),
        '起始井号': start,
        '终止井号': val('终止井号'),
        '录像文件': val('录像文件'),
        '敷设年代': val('敷设年代'),
        '起点埋深': val('起点埋深'),
        '终点埋深': val('终点埋深'),
        '管段类型': val('管段类型'),
        '管段材质': val('管段材质'),
        '管段直径': val('管段直径'),
        '检测方向': val('检测方向'),
        '管段长度': val('管段长度'),
        '检测长度': val('检测长度'),
        '修复指数': val('修复指数'),
        '养护指数': val('养护指数'),
        '检测人员': val('检测人员'),
    }
    for lbl in DEFECT_LABELS:
        arr = slots[lbl]
        if lbl == '图片':
            rec[lbl] = [v for v in arr if v and v not in ('/', '--')]
        else:
            rec[lbl] = [v if v else '/' for v in arr]
    return rec


def cells0(rows):
    """卡片序号：第一行第一个单元格（越界页为 '0'）。"""
    return rows[0][0] if rows and rows[0] else ''


def crawl_4127(sess, report, max_pages=800, sleep=0.0, verbose=True):
    records, seen, page = [], set(), 0
    stop = 'cap'
    while page < max_pages:
        rec = parse_4127(sess.get(page_url(sess, '4127', report, page)))
        if rec is None:
            stop = 'empty'
            break
        key = (rec['起始井号'], rec['终止井号'])
        if key in seen:                  # 防御：服务端重复吐同一张卡
            stop = 'repeat'
            break
        seen.add(key)
        records.append(rec)
        if verbose and (page % 25 == 0 or page == 0):
            print('    [S4] 4127 第%d页 %s-%s（累计 %d）'
                  % (page, rec['起始井号'], rec['终止井号'], len(records)))
        page += 1
        if sleep:
            time.sleep(sleep)
    stats = {'pages': page, 'records': len(records), 'stop_reason': stop,
             'unique_pairs': len(seen),
             'max_defects_per_pipe': max((sum(1 for _ in r['长度(m)']) for r in records), default=0)}
    return records, stats


# ==================== S3' 抓 4122 / 4123 / 4125 ====================
def _num(v):
    """纯数字串转 int/float，其余原样返回（工程量、统计数便于在 Excel 里求和）。"""
    s = str(v).strip()
    if re.fullmatch(r'-?\d+', s):
        return int(s)
    if re.fullmatch(r'-?\d+\.\d+', s):
        return float(s)
    return v


def _num_in(text, pattern):
    """从说明文字里抠一个数字，如「其中77段管道存在缺陷」→ 77。"""
    m = re.search(pattern, str(text or ''))
    return int(m.group(1)) if m else None


# ---------- 4122 主要工程量表 ----------
def parse_4122(html_text):
    """
    单页表：说明行(1 td) + 表头(5) + 工程量行(5) + 总计行(4，序号与名称被 colspan 合并)。
    平台每页还带若干「序号有、名称空」的预留行，必须丢弃。
    """
    body = _data_table(html_text)
    if body is None:
        return '', [], None
    summary, rows, total = '', [], None
    for tr in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S):
        cells = [_cell_text(x) for x in re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S)]
        if len(cells) == 1 and cells[0]:
            if '共检测' in cells[0] or '段管道' in cells[0]:
                summary = summary or cells[0]
            continue
        if cells[:3] == ['序号', '名称', '管道长度m']:
            continue
        if cells and cells[0] == '总计':
            if len(cells) == 4:
                total = {'length': _num(cells[1]), 'detect': _num(cells[2]), 'note': cells[3]}
            elif len(cells) >= 5:
                total = {'length': _num(cells[2]), 'detect': _num(cells[3]), 'note': cells[4]}
            continue
        if len(cells) == 5 and cells[0].isdigit() and cells[1]:
            rows.append([cells[0], cells[1], _num(cells[2]), _num(cells[3]), cells[4]])
    return summary, rows, total


def crawl_4122(sess, report, max_pages=10, verbose=True):
    """
    第 0 页有数据，之后是「空壳页」：说明为空、行全空、总计为空。
    停页依据：整页既没有名称非空的行，也没有带值的总计。
    """
    summary, rows, total = '', [], None
    pages, stop = 0, 'cap'
    while pages < max_pages:
        s2, r2, t2 = parse_4122(sess.get(page_url(sess, '4122', report, pages)))
        if not r2 and not (t2 and t2.get('length') not in ('', None)):
            stop = 'empty'
            break
        summary = summary or s2
        rows.extend(r2)
        total = t2 or total
        pages += 1
    stats = {'pages': pages, 'rows': len(rows), 'stop_reason': stop,
             'has_total': bool(total), 'segments_in_note': _segments_in_note(rows)}
    if verbose:
        print('    [S3b] 4122 主要工程量表 %d 行 / %d 页 / 停页=%s'
              % (len(rows), pages, stop))
    return summary, rows, total, stats


def _segments_in_note(rows):
    """汇总备注里的「分N段」之和，用于与 4126 行数对账。"""
    total = 0
    for r in rows:
        n = _num_in(r[4] if len(r) > 4 else '', r'分\s*(\d+)\s*段')
        if n:
            total += n
    return total


# ---------- 4123 管道缺陷汇总一览表 ----------
def parse_4123(html_text):
    """说明行(1 td) + 表头(6) + 数据行(6)，每页 12 行。"""
    body = _data_table(html_text)
    if body is None:
        return '', []
    summary, rows = '', []
    for tr in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S):
        cells = [_cell_text(x) for x in re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S)]
        if len(cells) == 1 and cells[0]:
            if '共评估' in cells[0] or '段管道' in cells[0]:
                summary = summary or cells[0]
            continue
        if (len(cells) == 6 and cells[0].isdigit()
                and _seg_key(cells[1]) not in ('', '/', '管段编号')):
            rows.append(cells)
    return summary, rows


def crawl_4123(sess, report, max_pages=60, verbose=True):
    """按管段编号去重翻页；整页无数据 → empty，整页重复 → repeat。"""
    summary, rows, seen = '', [], set()
    pages, dup, stop = 0, 0, 'cap'
    while pages < max_pages:
        s2, r2 = parse_4123(sess.get(page_url(sess, '4123', report, pages)))
        if not r2:
            stop = 'empty'
            break
        new = []
        for r in r2:
            k = _seg_key(r[1])
            if k in seen:
                dup += 1
                continue
            seen.add(k)
            new.append(r)
        if not new:
            stop = 'repeat'
            break
        summary = summary or s2
        rows.extend(new)
        pages += 1
    stats = {'pages': pages, 'rows': len(rows), 'duplicate_rows': dup,
             'stop_reason': stop, 'unique_segments': len(seen)}
    if verbose:
        print('    [S3c] 4123 汇总一览表 %d 行 / %d 页 / 停页=%s'
              % (len(rows), pages, stop))
    return summary, rows, stats


# ---------- 4125 管道缺陷数量统计表 ----------
def parse_4125(html_text):
    """
    单页统计表：16 个缺陷行 + 合计行。
    行有两种格式（缺陷类别单元格按 rowspan 合并）：
        8 td: 序号 | 类别 | 缺陷名称 | 1级..4级 | 小计
        7 td: 序号 |        缺陷名称 | 1级..4级 | 小计
    合计行 7 td：'' | 合计 | 1级..4级 | 小计
    """
    body = _data_table(html_text)
    if body is None:
        return '', [], None
    summary, defects, total, cat = '', [], None, ''
    for tr in re.findall(r'<tr[^>]*>(.*?)</tr>', body, re.S):
        cells = [_cell_text(x) for x in re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S)]
        if len(cells) == 1 and cells[0]:
            if '共评估' in cells[0] or '段管道' in cells[0]:
                summary = summary or cells[0]
            continue
        if len(cells) >= 2 and cells[0] == '序号' and cells[1] == '缺陷名称':
            continue
        if len(cells) == 7 and cells[1] == '合计':
            total = {'1级': _num(cells[2]), '2级': _num(cells[3]), '3级': _num(cells[4]),
                     '4级': _num(cells[5]), '小计': _num(cells[6])}
            continue
        if not cells or not cells[0].isdigit() or cells[-1] in ('', None):
            continue          # 空壳页整行都是空值
        if len(cells) == 8:
            cat, name, nums = cells[1], cells[2], cells[3:]
        elif len(cells) == 7:
            name, nums = cells[1], cells[2:]
        else:
            continue
        if len(nums) != 5:
            continue
        defects.append({'序号': cells[0], '缺陷类别': cat, '缺陷名称': name,
                        '1级': _num(nums[0]), '2级': _num(nums[1]), '3级': _num(nums[2]),
                        '4级': _num(nums[3]), '小计': _num(nums[4])})
    return summary, defects, total


def crawl_4125(sess, report, verbose=True):
    """固定 1 页；越界页 16 行全空，按「无有效缺陷行」停页。"""
    summary, defects, total = '', [], None
    pages, stop = 0, 'cap'
    while pages < 5:
        s2, d2, t2 = parse_4125(sess.get(page_url(sess, '4125', report, pages)))
        if not d2:
            stop = 'empty'
            break
        summary = summary or s2
        defects, total = d2, t2
        pages += 1
        if len(defects) >= 16:      # 固定 16 个缺陷行，说明整表已抓完
            stop = 'full'
            break
    stats = {'pages': pages, 'rows': len(defects), 'stop_reason': stop,
             'has_total': bool(total)}
    if verbose:
        print('    [S3d] 4125 数量统计表 %d 行 / %d 页 / 停页=%s'
              % (len(defects), pages, stop))
    return summary, defects, total, stats


# ==================== S5 落盘 ====================
def _write_sheet(ws, headers, widths, rows, freeze):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(name='新宋体', size=10, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name='新宋体', size=9)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = BORDER
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze


def output_name(table, bgbh, ts):
    return '%s-%s(网页爬取)_%s.xlsx' % (TABLE_NAMES[table], bgbh, ts)


def save_4126(rows, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '管段状况评估表'
    _write_sheet(ws, HEADERS_4126, WIDTHS_4126, rows, 'B2')
    wb.save(path)
    return path


def save_4127(records, path):
    rows = []
    for rec in records:
        row = [rec.get(h if h in rec else '', '') for h in HEADERS_4127[:16]]
        for i in range(DEFECT_SLOTS):
            for lbl in ['长度(m)', '缺陷名称代码', '分值', '等级', '管道内部状况描述']:
                arr = rec.get(lbl, [])
                row.append(arr[i] if i < len(arr) else '/')
        row.append('/'.join(rec.get('图片', [])))
        rows.append(row)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '管段检测与评估成果表'
    _write_sheet(ws, HEADERS_4127, WIDTHS_4127, rows, 'C2')
    wb.save(path)
    return path


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


# ---------- 4122 / 4123 / 4125 的三段式版式：说明行 + 表头 + 表体(+总计) ----------
def _summary_block(ws, ncols, text):
    """说明行合并 A1:<ncols>1，左对齐 10pt。返回表头所在行号（1 或 3）。"""
    if not text:
        return 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=text)
    c.font = Font(name='新宋体', size=10)
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    return 3


def _table_body(ws, headers, widths, body, hr):
    """表头写在第 hr 行，表体紧随其后。返回表体下一行的行号。"""
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=hr, column=i, value=h)
        c.font = Font(name='新宋体', size=10, bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    for r, row in enumerate(body, hr + 1):
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name='新宋体', size=9)
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border = BORDER
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return hr + len(body) + 1


def _write_total(ws, row, values, merge_to=None):
    """写合计/总计行；merge_to 指定要与第 1 列合并到的列号（4122 的总计是 A:B 合并）。"""
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = Font(name='新宋体', size=9, bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = BORDER
    if merge_to:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_to)
        ws.cell(row=row, column=merge_to).border = BORDER


def save_4122(summary, rows, total, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '主要工程量表'
    hr = _summary_block(ws, len(HEADERS_4122), summary)
    r = _table_body(ws, HEADERS_4122, WIDTHS_4122, rows, hr)
    if total:
        _write_total(ws, r, ['总计', '', total.get('length'), total.get('detect'),
                             total.get('note', '')], merge_to=2)
    wb.save(path)
    return path


def save_4123(summary, rows, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '管道缺陷汇总一览表'
    hr = _summary_block(ws, len(HEADERS_4123), summary)
    _table_body(ws, HEADERS_4123, WIDTHS_4123, rows, hr)
    wb.save(path)
    return path


def save_4125(summary, defects, total, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '管道缺陷数量统计表'
    hr = _summary_block(ws, len(HEADERS_4125), summary)
    body = [[d['序号'], d['缺陷类别'], d['缺陷名称'],
             d['1级'], d['2级'], d['3级'], d['4级'], d['小计']] for d in defects]
    r = _table_body(ws, HEADERS_4125, WIDTHS_4125, body, hr)
    if total:
        _write_total(ws, r, ['', '', '合计', total.get('1级'), total.get('2级'),
                             total.get('3级'), total.get('4级'), total.get('小计')])
    wb.save(path)
    return path


# ==================== S6 交叉校验 ====================
def _seg_key(seg):
    return re.sub(r'\s+', '', str(seg or '')).strip()


def _is_real_defect(code):
    return str(code or '').strip() not in NORMAL_CODES


def _code_type(code):
    """缺陷代码 → 'S' 结构 / 'F' 功能 / '?' 未识别。"""
    m = re.match(r'\(([A-Z]+)\)', str(code or ''))
    k = m.group(1) if m else ''
    if k in STRUCT_CODES:
        return 'S'
    if k in FUNC_CODES:
        return 'F'
    return '?'


def _defect_pipes(recs_4127):
    """4127 里带真实缺陷的管段（缺陷代码不是 (ZC)正常 / -- / 空）。"""
    return [r for r in recs_4127 if any(_is_real_defect(c) for c in r.get('缺陷名称代码', []))]


def _defect_entries(recs_4127):
    return sum(1 for r in recs_4127 for c in r.get('缺陷名称代码', []) if _is_real_defect(c))


def verify(rows_4126, recs_4127, extra=None):
    """
    校验 4126 与 4127 的管段集合，并对 4122/4123/4125 做跨表对账。

    注意：两表的管段编号口径不同 —— 4126 是「W1-W2」单列，4127 是起始井号/终止井号两列。
    井号本身可能含 '-'（W66-1、W10-1-W10-2），所以**绝不能**用 split('-') 反拆 4126 的字符串；
    这里改成「用 4127 的 (起, 止) 正向/反向拼出候选串」再与 4126 比对，天然免疫多横杠。
    另有实测存在的「流向相反」管段（4126: W47-W46，4127: 起始 W46 终止 W47，
    以及 11-05 时的 W66-2/W66-3），因此先做有向比对，再对有向差集做反向比对。

    extra（可选）：{'4123': {'summary','rows'}, '4125': {'summary','defects','total'},
                   '4122': {'summary','rows','total','segments_in_note'}}
    """
    a_directed = {_seg_key(r[1]) for r in rows_4126}
    pairs = [(_seg_key(r['起始井号']), _seg_key(r['终止井号'])) for r in recs_4127]
    b_directed = {'%s-%s' % (s, e) for s, e in pairs}
    rev_of = {'%s-%s' % (e, s): (s, e) for s, e in pairs}
    pair_of = {'%s-%s' % (s, e): (s, e) for s, e in pairs}

    only_a = sorted(a_directed - b_directed)
    flipped = [seg for seg in only_a if seg in rev_of]
    unmatched_a = [seg for seg in only_a if seg not in rev_of]

    unmatched_b = []
    for seg in sorted(b_directed - a_directed):
        s, e = pair_of[seg]
        if '%s-%s' % (e, s) in a_directed:      # 已由 flipped 解释
            continue
        unmatched_b.append(seg)

    dup_a = len(rows_4126) - len(a_directed)
    dup_b = len(recs_4127) - len(b_directed)
    overflow = [r for r in recs_4127 if len(r.get('长度(m)', [])) > DEFECT_SLOTS]

    checks = [
        {'id': 'G3-4126-nonempty', 'level': 'PASS' if rows_4126 else 'FAIL',
         'msg': '4126 抓到 %d 行' % len(rows_4126)},
        {'id': 'G4-4127-nonempty', 'level': 'PASS' if recs_4127 else 'FAIL',
         'msg': '4127 抓到 %d 个管段' % len(recs_4127)},
        {'id': 'G3-no-dup-4126', 'level': 'PASS' if dup_a == 0 else 'FAIL',
         'msg': '4126 重复管段 %d 条' % dup_a},
        {'id': 'G4-no-dup-4127', 'level': 'PASS' if dup_b == 0 else 'FAIL',
         'msg': '4127 重复管段 %d 条' % dup_b},
        {'id': 'G5-count-match', 'level': 'PASS' if len(rows_4126) == len(recs_4127) else 'FAIL',
         'msg': '两表管段数 %d vs %d' % (len(rows_4126), len(recs_4127))},
        {'id': 'G5-direction-flip', 'level': 'WARN' if flipped else 'PASS',
         'msg': '流向相反的管段 %d 条%s' % (len(flipped), ('：' + '、'.join(flipped)) if flipped else '')},
        {'id': 'G5-unmatched', 'level': 'FAIL' if (unmatched_a or unmatched_b) else 'PASS',
         'msg': '仅 4126 有 %d 条%s；仅 4127 有 %d 条%s' % (
             len(unmatched_a), ('：' + '、'.join(unmatched_a[:10])) if unmatched_a else '',
             len(unmatched_b), ('：' + '、'.join(unmatched_b[:10])) if unmatched_b else '')},
        {'id': 'G6-defect-slots', 'level': 'WARN' if overflow else 'PASS',
         'msg': '缺陷超过 %d 槽位的管段 %d 条%s' % (
             DEFECT_SLOTS, len(overflow),
             ('：' + '、'.join('%s-%s' % (r['起始井号'], r['终止井号']) for r in overflow[:10]))
             if overflow else '')},
    ]

    # ---------- G7/G8/G9：与 4123 / 4125 / 4122 对账 ----------
    extra = extra or {}
    t23, t25, t22 = extra.get('4123'), extra.get('4125'), extra.get('4122')
    defect_pipes = len(_defect_pipes(recs_4127))
    entries = _defect_entries(recs_4127)

    # 4127 按缺陷代码分类（结构/功能）
    type_s = type_f = 0
    unknown = set()
    for r in recs_4127:
        for c in r.get('缺陷名称代码', []):
            if not _is_real_defect(c):
                continue
            k = _code_type(c)
            if k == 'S':
                type_s += 1
            elif k == 'F':
                type_f += 1
            else:
                unknown.add(c)
    if unknown:
        checks.append({'id': 'G8-unknown-code', 'level': 'WARN',
                       'msg': '未识别缺陷代码 %s' % '、'.join(sorted(unknown))})

    st_desc = fn_desc = None
    st25 = fn25 = None

    if t23 is not None and t23.get('rows') is not None:
        rows23 = t23['rows']
        st_desc = sum(1 for r in rows23 if str(r[4]).strip() not in ('/', '', '-', '—'))
        fn_desc = sum(1 for r in rows23 if str(r[5]).strip() not in ('/', '', '-', '—'))
        checks.append({'id': 'G7-4123-vs-4127',
                       'level': 'PASS' if len(rows23) == defect_pipes else 'FAIL',
                       'msg': '4123 缺陷段 %d vs 4127 缺陷段 %d' % (len(rows23), defect_pipes)})

        # 4123 里出现的管段必须都能在 4126 找到；只有流向相反才算可接受
        unmatched23, flipped23 = [], []
        for r in rows23:
            seg = _seg_key(r[1])
            if seg in a_directed:
                continue
            pair = rev_of.get(seg)
            if pair and '%s-%s' % (pair[0], pair[1]) in a_directed:
                flipped23.append(seg)
            else:
                unmatched23.append(seg)
        checks.append({'id': 'G7b-4123-in-4126',
                       'level': 'FAIL' if unmatched23 else 'PASS',
                       'msg': '4123 有 %d 个管段在 4126 中找不到%s' % (
                           len(unmatched23),
                           ('：' + '、'.join(unmatched23[:10])) if unmatched23 else '')})
        if flipped23:
            checks.append({'id': 'G7b-4123-flip', 'level': 'WARN',
                           'msg': '4123 与 4126 流向相反 %d 条：%s'
                                  % (len(flipped23), '、'.join(flipped23[:10]))})

        n23 = _num_in(t23.get('summary'), r'其中\s*(\d+)\s*段')
        if n23 is not None:
            checks.append({'id': 'G7-4123-summary',
                           'level': 'PASS' if n23 == len(rows23) else 'WARN',
                           'msg': '4123 说明文字称 %d 段，表内实际 %d 行' % (n23, len(rows23))})

    if t25 is not None and t25.get('total'):
        sub = t25['total'].get('小计')
        if t25.get('defects'):
            st25 = sum(d['小计'] for d in t25['defects'] if '结构' in str(d['缺陷类别']))
            fn25 = sum(d['小计'] for d in t25['defects'] if '功能' in str(d['缺陷类别']))
        if st_desc is not None:
            checks.append({'id': 'G8-4125-vs-4123',
                           'level': 'PASS' if sub == st_desc + fn_desc else 'FAIL',
                           'msg': '4125 合计 %s 处 vs 4123 缺陷描述 %d 条'
                                  % (sub, st_desc + fn_desc)})
            if st25 is not None and (st25 != st_desc or fn25 != fn_desc):
                checks.append({'id': 'G7c-4123-type', 'level': 'WARN',
                               'msg': '4123 描述分类(结构 %d/功能 %d) 与 4125(结构 %d/功能 %d) 不一致'
                                      % (st_desc, fn_desc, st25, fn25)})
        checks.append({'id': 'G8-4125-vs-4127',
                       'level': 'PASS' if sub == entries else 'WARN',
                       'msg': '4125 合计 %s 处 vs 4127 缺陷条目 %d 条' % (sub, entries)})
        if st25 is not None and (st25 != type_s or fn25 != type_f):
            checks.append({'id': 'G8c-4125-type', 'level': 'WARN',
                           'msg': '4125 分类(结构 %d/功能 %d) 与 4127 代码分类(结构 %d/功能 %d) 不一致'
                                  % (st25, fn25, type_s, type_f)})
        n25 = _num_in(t25.get('summary'), r'存在缺陷\s*(\d+)\s*处')
        if n25 is not None:
            checks.append({'id': 'G8-4125-summary',
                           'level': 'PASS' if n25 == sub else 'WARN',
                           'msg': '4125 说明文字称 %s 处，合计行 %s 处' % (n25, sub)})

    if t22 is not None and t22.get('total'):
        segs = t22.get('segments_in_note') or 0
        if segs:
            checks.append({'id': 'G9-4122-segments',
                           'level': 'PASS' if segs == len(rows_4126) else 'WARN',
                           'msg': '4122 备注「分N段」合计 %d 段 vs 4126 %d 行' % (segs, len(rows_4126))})
        length = t22['total'].get('length')
        nums = [float(str(r[4]).strip()) for r in rows_4126
                if re.fullmatch(r'[\d.]+', str(r[4]).strip())]
        if isinstance(length, (int, float)) and nums:
            s_len = sum(nums)
            diff = s_len - float(length)
            ok = abs(diff) <= max(5.0, 0.005 * s_len)   # 5m 或 0.5% 以内视为一致
            checks.append({'id': 'G9-4122-length', 'level': 'PASS' if ok else 'WARN',
                           'msg': '4122 总计长度 %s m vs 4126 合计 %.1f m（差 %+.1f m）'
                                  % (length, s_len, diff)})
        n22 = _num_in(t22.get('summary'), r'共检测\s*(\d+)\s*段')
        if n22 is not None:
            checks.append({'id': 'G9-4122-summary',
                           'level': 'PASS' if n22 == len(rows_4126) else 'WARN',
                           'msg': '4122 说明文字称检测 %d 段 vs 4126 %d 行' % (n22, len(rows_4126))})

    return {'checks': checks, 'direction_flipped': flipped,
            'only_4126': unmatched_a, 'only_4127': unmatched_b,
            'dup_4126': dup_a, 'dup_4127': dup_b,
            'defect_pipes': defect_pipes, 'defect_entries': entries,
            'hard_fail': any(c['level'] == 'FAIL' for c in checks)}


# ==================== S7 manifest ====================
def write_manifest(out_dir, report, stages, checks, outputs, params):
    os.makedirs(out_dir, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    doc = {
        'tool': 'cctv_workflow', 'version': VERSION,
        'run_at': datetime.datetime.now().isoformat(timespec='seconds'),
        'bgbh': report.get('bgbh'), 'report': report,
        'params': params, 'stages': stages, 'checks': checks,
        'outputs': [{'file': os.path.basename(p), 'bytes': os.path.getsize(p),
                     'sha256': sha256(p)} for p in outputs],
        'verdict': 'FAIL' if any(c['level'] == 'FAIL' for c in checks) else
                   ('WARN' if any(c['level'] == 'WARN' for c in checks) else 'PASS'),
    }
    path = os.path.join(out_dir, 'manifest_%s_%s.json' % (report.get('bgbh'), ts))
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    return path


# ==================== 编排 ====================
def run_all(cfg, args):
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    out_dir = args.out_dir or '.'
    os.makedirs(out_dir, exist_ok=True)
    report = resolve_report(cfg, args)
    report['bgbh'] = args.bgbh

    sess = Session(cfg, verbose=True).login()

    print('[S3] 抓取 4126 管段状况评估表 …')
    rows, st3 = crawl_4126(sess, report, max_pages=args.max_pages_4126, sleep=args.sleep)
    print('     → %d 行 / %d 页 / 停页原因=%s' % (st3['rows'], st3['pages'], st3['stop_reason']))
    print('       （表头行 %d / 每页空预留行 %d / 结构异常 %d）'
          % (st3['header_rows'], st3['placeholder_rows'], st3['malformed_rows']))

    print('[S4] 抓取 4127 管段检测与评估成果表 …')
    recs, st4 = crawl_4127(sess, report, max_pages=args.max_pages_4127, sleep=args.sleep)
    print('     → %d 个管段 / %d 页 / 停页原因=%s' % (st4['records'], st4['pages'], st4['stop_reason']))

    print('[S3c] 抓取 4123 管道缺陷汇总一览表 …')
    s23, rows23, st23 = crawl_4123(sess, report, verbose=False)
    print('     → %d 行 / %d 页 / 停页原因=%s' % (st23['rows'], st23['pages'], st23['stop_reason']))

    print('[S3d] 抓取 4125 管道缺陷数量统计表 …')
    s25, defects25, tot25, st25 = crawl_4125(sess, report, verbose=False)
    print('     → %d 行 / %d 页 / 合计=%s' % (st25['rows'], st25['pages'],
                                             (tot25 or {}).get('小计')))

    print('[S3b] 抓取 4122 主要工程量表 …')
    s22, rows22, tot22, st22 = crawl_4122(sess, report, verbose=False)
    print('     → %d 行 / %d 页 / 总计长度=%s' % (st22['rows'], st22['pages'],
                                               (tot22 or {}).get('length')))

    if not args.allow_empty:
        empty = [name for name, n in [('4126', len(rows)), ('4127', len(recs)),
                                      ('4123', len(rows23)), ('4125', len(defects25)),
                                      ('4122', len(rows22))] if not n]
        if empty:
            raise RuntimeError('%s 抓取结果为 0，拒绝落盘（加 --allow-empty 可强制）'
                               % '、'.join(empty))

    print('[S5] 落盘 …')
    paths = {
        '4126': save_4126(rows, os.path.join(out_dir, output_name('4126', args.bgbh, ts))),
        '4127': save_4127(recs, os.path.join(out_dir, output_name('4127', args.bgbh, ts))),
        '4123': save_4123(s23, rows23, os.path.join(out_dir, output_name('4123', args.bgbh, ts))),
        '4125': save_4125(s25, defects25, tot25,
                          os.path.join(out_dir, output_name('4125', args.bgbh, ts))),
        '4122': save_4122(s22, rows22, tot22,
                          os.path.join(out_dir, output_name('4122', args.bgbh, ts))),
    }
    for t in ('4126', '4127', '4123', '4125', '4122'):
        print('     %s' % os.path.basename(paths[t]))

    print('[S6] 交叉校验 …')
    v = verify(rows, recs, extra={
        '4123': {'summary': s23, 'rows': rows23},
        '4125': {'summary': s25, 'defects': defects25, 'total': tot25},
        '4122': {'summary': s22, 'rows': rows22, 'total': tot22,
                 'segments_in_note': st22['segments_in_note']},
    })
    # 把采集阶段的过滤量也纳入清单，避免「悄悄丢行」
    # 每页末尾的空预留行属于正常现象（实测每页 1 行），只记入 manifest，不告警；
    # 结构异常行（非 21 列）才是页面改版的信号，必须告警。
    if st3.get('malformed_rows'):
        v['checks'].append({'id': 'G3-malformed', 'level': 'WARN',
                            'msg': '4126 结构异常(非21列)行 %d 行' % st3['malformed_rows']})
    if st3.get('duplicate_rows'):
        v['checks'].append({'id': 'G3-page-dup', 'level': 'WARN',
                            'msg': '4126 翻页中重复行已去重 %d 行' % st3['duplicate_rows']})
    for c in v['checks']:
        print('     [%s] %s' % (c['level'], c['msg']))

    print('[S7] 写清单 …')
    m = write_manifest(out_dir, report,
                       {'4126': st3, '4127': st4, '4123': st23, '4125': st25, '4122': st22},
                       v['checks'], [paths[t] for t in ('4126', '4127', '4123', '4125', '4122')],
                       {'sleep': args.sleep, 'days': args.days})
    print('     %s' % os.path.basename(m))
    return 0 if not v['hard_fail'] else 2


def _crawl_and_save(sess, report, table, out_dir, ts, args):
    """单表抓取 + 落盘，返回 (path, stats)。0 结果且未 --allow-empty 时 path=None。"""
    path = None
    stats = {}
    if table == '4126':
        rows, stats = crawl_4126(sess, report, args.max_pages_4126, args.sleep)
        if rows or args.allow_empty:
            path = save_4126(rows, os.path.join(out_dir, output_name('4126', args.bgbh, ts)))
    elif table == '4127':
        recs, stats = crawl_4127(sess, report, args.max_pages_4127, args.sleep)
        if recs or args.allow_empty:
            path = save_4127(recs, os.path.join(out_dir, output_name('4127', args.bgbh, ts)))
    elif table == '4123':
        s23, rows, stats = crawl_4123(sess, report)
        if rows or args.allow_empty:
            path = save_4123(s23, rows, os.path.join(out_dir, output_name('4123', args.bgbh, ts)))
    elif table == '4125':
        s25, defects, total, stats = crawl_4125(sess, report)
        if defects or args.allow_empty:
            path = save_4125(s25, defects, total,
                             os.path.join(out_dir, output_name('4125', args.bgbh, ts)))
    elif table == '4122':
        s22, rows, total, stats = crawl_4122(sess, report)
        if rows or args.allow_empty:
            path = save_4122(s22, rows, total,
                             os.path.join(out_dir, output_name('4122', args.bgbh, ts)))
    return path, stats


def resolve_report(cfg, args):
    """report 定位：显式参数 > HTTP 接口 > （可选）浏览器兜底。"""
    if args.rid and args.parentid:
        report = {'bgbh': args.bgbh, 'report_id': args.rid, 'parentid': args.parentid,
                  'fid': '412', 'gcmc': '', 'igid': cfg['IGID'], 'count': None}
        print('[S2] 使用显式参数 report_id=%s parentid=%s' % (args.rid, args.parentid))
        return report
    sess = Session(cfg, verbose=True).login()
    try:
        return locate_report(sess, args.bgbh, cfg['IGID'], days=args.days)
    except RuntimeError as e:
        if not args.browser_fallback:
            raise
        print('    [S2] HTTP 定位失败，改用浏览器兜底：%s' % str(e).splitlines()[0])
        return locate_via_browser(args.bgbh, headless=not args.headless)


def main(argv=None):
    ap = argparse.ArgumentParser(description='CCTV 网页成果表抓取工作流 v%s' % VERSION)
    ap.add_argument('--bgbh', default=None,
                    help='报告编号，如 26KHD00018（优先级：--bgbh > config.py 的 BGBH > 内置 DEFAULTS）')
    ap.add_argument('--rid', help='显式 report_id（跳过定位接口）')
    ap.add_argument('--parentid', help='显式 parentid / ZhuBiaoId')
    ap.add_argument('--out-dir', default='.', help='输出目录（默认当前目录）')
    ap.add_argument('--days', type=int, default=3650, help='台账查询回溯天数（默认 3650）')
    ap.add_argument('--max-pages-4126', type=int, default=80)
    ap.add_argument('--max-pages-4127', type=int, default=800)
    ap.add_argument('--sleep', type=float, default=0.0, help='翻页间隔秒（限速用）')
    ap.add_argument('--allow-empty', action='store_true')
    ap.add_argument('--browser-fallback', action='store_true', help='HTTP 定位失败时用 Playwright')
    ap.add_argument('--headless', action='store_true', help='浏览器兜底无头模式')
    ap.add_argument('--show-config', action='store_true',
                    help='只打印实际生效的配置（含来源），不联网')
    sub = ap.add_subparsers(dest='cmd')
    sub.add_parser('all', help='定位 → 抓 4126/4127/4123/4125/4122 → 校验 → manifest')
    sub.add_parser('locate', help='只定位报告')
    p_crawl = sub.add_parser('crawl', help='只抓一张表')
    p_crawl.add_argument('table', choices=['4122', '4123', '4125', '4126', '4127'])
    sub.add_parser('verify', help='校验目录里最新的一套 xlsx（离线）')
    sub.add_parser('selftest', help='离线自检')

    args = ap.parse_args(argv)
    cfg, meta = load_config()
    explicit_bgbh = args.bgbh
    args.bgbh = args.bgbh or cfg['BGBH']
    cmd = args.cmd or 'all'

    if cmd == 'selftest':
        return selftest()

    if args.show_config:
        print_config(cfg, meta, explicit_bgbh)
        return 0

    print_config(cfg, meta, explicit_bgbh)

    if cmd == 'locate':
        sess = Session(cfg, verbose=True).login()
        info = locate_report(sess, args.bgbh, cfg['IGID'], days=args.days)
        print(json.dumps(info, ensure_ascii=False, indent=2))
        return 0

    if cmd == 'crawl':
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        report = resolve_report(cfg, args)
        report['bgbh'] = args.bgbh
        sess = Session(cfg, verbose=True).login()
        path, st = _crawl_and_save(sess, report, args.table, args.out_dir, ts, args)
        print('%s 统计: %s' % (args.table, json.dumps(st, ensure_ascii=False)))
        if path:
            print('已保存', path)
        else:
            print('无数据，未落盘（加 --allow-empty 可强制）')
        return 0 if path else 2

    if cmd == 'verify':
        return verify_latest(args.bgbh, args.out_dir)

    return run_all(cfg, args)


def _latest_file(table, bgbh, out_dir='.'):
    import glob
    fs = sorted(glob.glob(os.path.join(out_dir, '%s-%s(网页爬取)*.xlsx'
                                       % (TABLE_NAMES[table], bgbh))))
    return fs[-1] if fs else None


def _read_grid(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = [[('' if c is None else c) for c in r] for r in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()
    return rows


def _read_223_output(table, path):
    """回读 4122/4123/4125 的三段式产物（说明行 1、表头 3、表体 4 起）。"""
    grid = _read_grid(path)
    summary = str(grid[0][0]) if grid and grid[0] and grid[0][0] else ''
    body, out = grid[3:], {'summary': summary}
    if table == '4123':
        out['rows'] = [[str(x) for x in r[:6]] for r in body if r and str(r[1]).strip()]
    elif table == '4125':
        defects, total = [], None
        for r in body:
            if len(r) >= 8 and str(r[2]).strip() == '合计':
                total = {'1级': r[3], '2级': r[4], '3级': r[5], '4级': r[6], '小计': r[7]}
            elif str(r[0]).strip() and str(r[2]).strip():
                defects.append({'序号': r[0], '缺陷类别': r[1], '缺陷名称': r[2],
                                '1级': r[3], '2级': r[4], '3级': r[5],
                                '4级': r[6], '小计': r[7]})
        out.update({'defects': defects, 'total': total})
    elif table == '4122':
        rows, total = [], None
        for r in body:
            if len(r) >= 5 and str(r[0]).strip() == '总计':
                total = {'length': r[2], 'detect': r[3], 'note': r[4]}
            elif str(r[1]).strip():
                rows.append(r)
        out.update({'rows': rows, 'total': total,
                    'segments_in_note': _segments_in_note([[str(x) for x in r] for r in rows])})
    return out


def load_outputs(bgbh, out_dir='.'):
    """读取目录里最新的一套五张表，返回 (rows_4126, recs_4127, extra, files)。"""
    files = {t: _latest_file(t, bgbh, out_dir) for t in ('4126', '4127', '4123', '4125', '4122')}
    rows, recs, extra = [], [], {}
    if files['4126']:
        rows = [[str(x) for x in r] for r in _read_grid(files['4126'])[1:]
                if len(r) > 1 and str(r[1]).strip()]
    if files['4127']:
        for r in _read_grid(files['4127'])[1:]:
            if len(r) < 3 or not str(r[1]).strip():
                continue
            recs.append({'起始井号': str(r[1]).strip(), '终止井号': str(r[2]).strip(),
                         '缺陷名称代码': [str(r[17 + 5 * i]).strip() for i in range(DEFECT_SLOTS)
                                          if len(r) > 17 + 5 * i and str(r[17 + 5 * i]).strip()],
                         '长度(m)': [r[16 + 5 * i] for i in range(DEFECT_SLOTS)
                                    if len(r) > 16 + 5 * i and r[16 + 5 * i] not in (None, '/')]})
    for t in ('4123', '4125', '4122'):
        if files[t]:
            extra[t] = _read_223_output(t, files[t])
    return rows, recs, extra, files


def verify_latest(bgbh, out_dir='.'):
    """对目录里最新的一套 4126/4127/4123/4125/4122 xlsx 做离线交叉校验。"""
    rows, recs, extra, files = load_outputs(bgbh, out_dir)
    if not files['4126'] or not files['4127']:
        print('未找到 %s 的 4126/4127 输出文件' % bgbh)
        return 1
    for t in ('4126', '4127', '4123', '4125', '4122'):
        if files[t]:
            print('%s: %s' % (t, os.path.basename(files[t])))
    v = verify(rows, recs, extra)
    for c in v['checks']:
        print('[%s] %s' % (c['level'], c['msg']))
    return 0 if not v['hard_fail'] else 2


# ==================== 自检（离线） ====================
def _fake_4126(pages):
    out = []
    for pg in pages:
        trs = ''
        for i in range(13):
            n = pg * 13 + i + 1
            if n > 26:
                break
            trs += ('<tr><td>%d</td><td>W%d-W%d</td><td>300</td><td>双壁波纹管</td>'
                    '<td>3.7</td><td>顺流</td>%s</tr>'
                    % (n, n, n + 1, '<td>/</td>' * 15))
        out.append('<table class="data-table layui-table"><tr>%s</tr>%s</table>'
                   % ('<td>h</td>' * 12, trs))
    return out


def _fake_4122():
    return ('<table class="data-table">'
            '<tr><td>本次共检测2段管道,评估2段管道,详见下表:</td></tr>'
            '<tr><td>序号</td><td>名称</td><td>管道长度m</td><td>检测长度</td><td>备注</td></tr>'
            '<tr><td>1</td><td>污水管道</td><td>10.5</td><td>10.5</td><td>分2段;管径300mm</td></tr>'
            '<tr><td>2</td><td></td><td></td><td></td><td></td></tr>'
            '<tr><td>总计</td><td>10.5</td><td>10.5</td><td>共2段</td></tr>'
            '</table>')


def _fake_4123():
    return ('<table class="data-table">'
            '<tr><td>本次共评估3段管道,其中1段管道存在缺陷（结构性缺陷管道1段）,详见下表:</td></tr>'
            '<tr><td>序号</td><td>管段编号</td><td>管径(mm)</td><td>管段材质</td>'
            '<td>结构性缺陷</td><td>功能性缺陷</td></tr>'
            '<tr><td>1</td><td>W1-W2</td><td>300</td><td>双壁波纹管</td>'
            '<td>纵向0.5m处1级变形</td><td>/</td></tr>'
            '</table>')


def _fake_4125():
    return ('<table class="data-table">'
            '<tr><td>本次共评估3段管道,其中存在缺陷2处（结构性缺陷1处,功能性缺陷1处）,详见下表:</td></tr>'
            '<tr><td>序号</td><td>缺陷名称</td><td>1级</td><td>2级</td><td>3级</td>'
            '<td>4级</td><td>小计</td></tr>'
            '<tr><td>1</td><td>结构性缺陷</td><td>(BX)变形</td><td>1</td><td>0</td>'
            '<td>0</td><td>0</td><td>1</td></tr>'
            '<tr><td>2</td><td>(CJ)沉积</td><td>1</td><td>0</td><td>0</td><td>0</td><td>1</td></tr>'
            '<tr><td></td><td>合计</td><td>2</td><td>0</td><td>0</td><td>0</td><td>2</td></tr>'
            '</table>')


def _fake_4127(k, codes=None, descs=None):
    """按真实卡片的「标签行 + 3 个值行」两段式缺陷块构造一页。"""
    codes = codes or ['(ZC)正常', '/', '/', '/']
    descs = descs or ['管段暂无明显缺陷', '/', '/', '/']

    def row(cells):
        return '<tr>' + ''.join('<td>%s</td>' % c for c in cells) + '</tr>'

    head = (row([str(k), '录像文件', 'a.avi', '起始井号', 'W%d' % k,
                 '终止井号', 'W%d' % (k + 1)])
            + row(['敷设年代', '/', '起点埋深', '/', '终点埋深', '/'])
            + row(['管段类型', '污水（WS）', '管段材质', '双壁波纹管', '管段直径', '300mm'])
            + row(['检测方向', '顺流', '管段长度', '3.7m', '检测长度', '3.7m'])
            + row(['修复指数', '/', '养护指数', '/', '检测人员', '/']))
    lens = ['1m', '1.2m', '/', '/']
    scores = ['2', '0.5', '/', '/']
    block1 = row(['长度(m)', lens[0], '缺陷名称代码', codes[0], '分值', scores[0]])
    for i in range(1, 4):
        block1 += row([lens[i], codes[i], scores[i]])
    grades = ['2', '1', '/', '/']
    pics = ['图1', '', '', '']
    block2 = row(['等级', grades[0], '管道内部状况描述', descs[0], '图片', pics[0]])
    for i in range(1, 4):
        block2 += row([grades[i], descs[i], pics[i]])
    return '<table class="data-table">%s%s%s</table>' % (head, block1, block2)


def selftest():
    ok = True

    def chk(name, cond, extra=''):
        nonlocal ok
        print('  [%s] %s%s' % ('PASS' if cond else 'FAIL', name, (' — ' + extra) if extra else ''))
        ok = ok and bool(cond)

    print('== 解析器 ==')
    pages = _fake_4126([0, 1, 2])
    rows0, meta0 = parse_4126(pages[0])
    chk('4126 单页 13 行', len(rows0) == 13, 'got %d' % len(rows0))
    chk('4126 表头行计入 header_rows', meta0['header_rows'] >= 1,
        'header=%d' % meta0['header_rows'])
    empty_rows, meta_e = parse_4126('<table class="data-table"><tr><td>h</td></tr></table>')
    chk('4126 越界页返回 0 行', empty_rows == [])

    ph_html = ('<table class="data-table"><tr>' + '<td>0</td><td>/</td>' + '<td>/</td>' * 19
               + '</tr></table>')
    ph_rows, ph_meta = parse_4126(ph_html)
    chk('4126 表尾 / 占位行被丢弃', ph_rows == [] and ph_meta['placeholder_rows'] == 1,
        'rows=%d placeholder=%d' % (len(ph_rows), ph_meta['placeholder_rows']))

    rec = parse_4127(_fake_4127(1))
    chk('4127 起始井号', rec and rec['起始井号'] == 'W1', str(rec and rec['起始井号']))
    chk('4127 录像文件', rec and rec['录像文件'] == 'a.avi')
    chk('4127 无缺陷槽位以 / 占位',
        rec and rec['缺陷名称代码'] == ['(ZC)正常', '/', '/', '/'],
        str(rec and rec['缺陷名称代码']))
    chk('4127 图片只留非空', rec and rec['图片'] == ['图1'], str(rec and rec['图片']))
    rec_multi = parse_4127(_fake_4127(2, codes=['(CJ)沉积', '(BX)变形', '/', '/'],
                                      descs=['纵向1m处存在2级沉积', '纵向1.2m处存在1级变形', '/', '/']))
    chk('4127 多缺陷卡片读到第 2 槽位',
        rec_multi['缺陷名称代码'][:2] == ['(CJ)沉积', '(BX)变形']
        and rec_multi['等级'][:2] == ['2', '1']
        and rec_multi['管道内部状况描述'][1] == '纵向1.2m处存在1级变形',
        str(rec_multi['缺陷名称代码']))
    chk('4127 越界卡片返回 None',
        parse_4127('<table class="data-table"><tr><td>0</td><td>起始井号</td><td></td></tr></table>') is None)

    print('== 校验器 ==')
    rows = [[str(i + 1), 'W%d-W%d' % (i + 1, i + 2)] + [''] * 19 for i in range(3)]
    recs = [{'起始井号': 'W1', '终止井号': 'W2', '长度(m)': []},
            {'起始井号': 'W2', '终止井号': 'W3', '长度(m)': []},
            {'起始井号': 'W3', '终止井号': 'W4', '长度(m)': []}]
    v = verify(rows, recs)
    chk('完全一致 → 无 FAIL', not v['hard_fail'])
    chk('完全一致 → 无方向翻转', v['direction_flipped'] == [])

    recs_flip = [dict(r) for r in recs]
    recs_flip[1] = {'起始井号': 'W3', '终止井号': 'W2', '长度(m)': []}
    v2 = verify(rows, recs_flip)
    chk('方向相反 → 识别为 flipped', v2['direction_flipped'] == ['W2-W3'],
        str(v2['direction_flipped']))
    chk('方向相反不算 hard fail', not v2['hard_fail'])

    rows_dup = rows + [rows[0]]
    v3 = verify(rows_dup, recs)
    chk('重复行 → FAIL', v3['hard_fail'] and v3['dup_4126'] == 1)

    recs_over = [dict(r) for r in recs]
    recs_over[0] = {'起始井号': 'W1', '终止井号': 'W2', '长度(m)': ['1', '2', '3', '4', '5']}
    v4 = verify(rows, recs_over)
    chk('缺陷>4 槽位 → WARN', any(c['id'] == 'G6-defect-slots' and c['level'] == 'WARN'
                                   for c in v4['checks']))

    # 回归：井号含 '-' 时不能靠 split('-') 反拆
    rows_h = [['1', 'W66-2-W66-3'] + [''] * 19]
    recs_h = [{'起始井号': 'W66-3', '终止井号': 'W66-2', '长度(m)': []}]
    v5 = verify(rows_h, recs_h)
    chk('多横杠井号 + 反向 → 识别为 flipped', v5['direction_flipped'] == ['W66-2-W66-3'],
        str(v5['direction_flipped']))
    chk('多横杠井号 → 不误报 unmatched', not v5['only_4127'] and not v5['only_4126'],
        'only4127=%s only4126=%s' % (v5['only_4127'], v5['only_4126']))

    print('== 4122 / 4123 / 4125 解析器 ==')
    s22, r22, t22 = parse_4122(_fake_4122())
    chk('4122 只收名称非空行', len(r22) == 1, 'rows=%d' % len(r22))
    chk('4122 总计长度', (t22 or {}).get('length') == 10.5, str(t22))
    chk('4122 备注「分N段」求和', _segments_in_note(r22) == 2)

    s23, r23 = parse_4123(_fake_4123())
    chk('4123 数据行', len(r23) == 1 and r23[0][1] == 'W1-W2', str(r23))
    chk('4123 说明行', '本次共评估' in s23)

    s25, d25, t25 = parse_4125(_fake_4125())
    chk('4125 缺陷行数', len(d25) == 2, 'rows=%d' % len(d25))
    chk('4125 缺陷类别向下继承', d25[1]['缺陷类别'] == '结构性缺陷', str(d25[1]['缺陷类别']))
    chk('4125 合计', (t25 or {}).get('小计') == 2, str(t25))

    print('== 跨表对账 G7/G8/G9 ==')
    recs_def = [{'起始井号': 'W1', '终止井号': 'W2', '长度(m)': [], '缺陷名称代码': ['(BX)变形']},
                {'起始井号': 'W2', '终止井号': 'W3', '长度(m)': [], '缺陷名称代码': []},
                {'起始井号': 'W3', '终止井号': 'W4', '长度(m)': [], '缺陷名称代码': []}]
    extra_ok = {
        '4123': {'summary': '本次共评估3段管道,其中1段管道存在缺陷', 'rows': r23},
        '4125': {'summary': '本次共评估3段管道,其中存在缺陷1处',
                 'defects': [{'序号': '1', '缺陷类别': '结构性缺陷', '缺陷名称': '(BX)变形',
                              '1级': 1, '2级': 0, '3级': 0, '4级': 0, '小计': 1}],
                 'total': {'1级': 1, '2级': 0, '3级': 0, '4级': 0, '小计': 1}},
        '4122': {'summary': '本次共检测3段管道', 'rows': r22,
                 'total': {'length': 0, 'detect': 0, 'note': ''}, 'segments_in_note': 0},
    }
    v6 = verify(rows, recs_def, extra_ok)
    ids = {c['id']: c['level'] for c in v6['checks']}
    chk('G7 4123 段数与 4127 一致 → PASS', ids.get('G7-4123-vs-4127') == 'PASS', str(ids))
    chk('G7b 4123 管段都在 4126 → PASS', ids.get('G7b-4123-in-4126') == 'PASS')
    chk('G8 4125 合计与 4123 描述一致 → PASS', ids.get('G8-4125-vs-4123') == 'PASS')
    chk('G8 4125 说明与合计一致 → PASS', ids.get('G8-4125-summary') == 'PASS')
    chk('G8c 4125 分类与 4127 代码分类一致 → PASS', 'G8c-4125-type' not in ids, str(ids))
    chk('G9 4122 段数与 4126 一致 → PASS', ids.get('G9-4122-summary') == 'PASS')

    extra_bad = {
        '4123': {'summary': '本次共评估3段管道,其中2段管道存在缺陷',
                 'rows': r23 + [['2', 'W2-W3', '300', '双壁波纹管', '/', '纵向1m处1级沉积']]},
        '4125': {'summary': '本次共评估3段管道,其中存在缺陷1处', 'defects': d25,
                 'total': {'1级': 1, '2级': 0, '3级': 0, '4级': 0, '小计': 1}},
    }
    v7 = verify(rows, recs_def, extra_bad)
    chk('G7 4123 多一段 → FAIL', v7['hard_fail'])
    chk('G8 4125 合计对不上 4123 → FAIL',
        any(c['id'] == 'G8-4125-vs-4123' and c['level'] == 'FAIL' for c in v7['checks']))

    v8 = verify(rows, recs_def, {'4123': {'summary': '本次共评估3段管道,其中1段管道存在缺陷',
                                          'rows': [['1', 'X9-X10', '300', '双壁波纹管',
                                                    '纵向0.5m处1级变形', '/']]}})
    chk('G7b 4123 出现 4126 没有的管段 → FAIL',
        any(c['id'] == 'G7b-4123-in-4126' and c['level'] == 'FAIL' for c in v8['checks']))

    v9 = verify(rows, recs_def, {'4123': {'summary': '本次共评估3段管道,其中1段管道存在缺陷',
                                          'rows': [['1', 'W2-W1', '300', '双壁波纹管',
                                                    '纵向0.5m处1级变形', '/']]}})
    chk('G7b 4123 流向相反 → WARN 而非 FAIL',
        (not v9['hard_fail']
         and any(c['id'] == 'G7b-4123-flip' and c['level'] == 'WARN' for c in v9['checks'])))

    print('\n自检结果：%s' % ('全部通过' if ok else '存在失败项'))
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())

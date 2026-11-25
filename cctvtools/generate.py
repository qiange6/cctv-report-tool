# -*- coding: utf-8 -*-
"""生成：由 4126 评估表 + 4127 成果表 生成 4122 工程量表 / 4123 汇总表 / 4125 数量统计表。"""
from __future__ import annotations

import datetime
import os
import re
from collections import Counter

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from .common import EMPTY, STRUCT_CODES

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CN = {'AJ': '支管暗接', 'BX': '变形', 'CK': '错口', 'CR': '异物穿入', 'FS': '腐蚀', 'PL': '破裂',
      'QF': '起伏', 'SL': '渗漏', 'TJ': '脱节', 'TL': '接口材料脱落',
      'CJ': '沉积', 'CQ': '残墙、坝根', 'FZ': '浮渣', 'JG': '结垢', 'SG': '树根', 'ZW': '障碍物'}
STRUCT_ORDER = ['AJ', 'BX', 'CK', 'CR', 'FS', 'PL', 'QF', 'SL', 'TJ', 'TL']
FUNC_ORDER = ['CJ', 'CQ', 'FZ', 'JG', 'SG', 'ZW']


def _real(c):
    return str(c).strip() not in EMPTY and '(ZC)' not in str(c)


def _code(c):
    m = re.match(r'\(([A-Z]+)\)', str(c))
    return m.group(1) if m else ''


def _seg(v):
    """管段编号去空白（'W-15-W-21'）。"""
    return re.sub(r'\s+', '', str(v or '')).strip()


def _to_float(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or '').strip().lower().replace('m', '').replace('米', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _fmt(v):
    """数值 → 平台习惯的字符串（1608.8 / 600 / 4.4）。"""
    s = '%.1f' % float(v)
    return s[:-2] if s.endswith('.0') else s


def _pipe_name(t):
    """'污水（WS）' → '污水管道'；已带“管道”的原样保留。"""
    t = str(t or '').strip()
    m = re.match(r'^(.+?)[（(][^）)]*[）)]$', t)
    base = (m.group(1) if m else t).strip()
    if not base:
        return t
    return base if base.endswith('管道') else base + '管道'


def build_4122(rows4126, rows4127):
    """由 4126/4127 汇总 4122 主要工程量表。

    分组 = 管段类型 + 管段材质 + 管段直径（均取自 4127）；
    长度优先取 4126 的“管道长度(m)”（4127 是 '3.8m' 文本），
    按长度从大到小排序，与平台现有 4122 的排法一致。
    """
    by_seg = {}
    for r in rows4126:
        if r and len(r) > 4 and str(r[1]).strip():
            by_seg[_seg(r[1])] = r

    groups = {}
    for r in rows4127:
        if not r or len(r) < 13 or not str(r[1]).strip():
            continue
        s, e = str(r[1]).strip(), str(r[2]).strip()
        src = None
        for cand in ('%s-%s' % (s, e), '%s-%s' % (e, s)):
            src = by_seg.get(_seg(cand))
            if src is not None:
                break
        length = _to_float(src[4]) if src is not None else _to_float(r[11])
        detect = _to_float(r[12])
        name = _pipe_name(r[7])
        mat = str(r[8]).strip()
        dia = str(r[9]).strip()
        g = groups.setdefault((name, mat, dia),
                              {'n': 0, 'length': 0.0, 'detect': 0.0})
        g['n'] += 1
        g['length'] += length
        g['detect'] += detect

    n_total = len([r for r in rows4126 if r and str(r[1]).strip()])
    items = sorted(groups.items(), key=lambda kv: (-kv[1]['length'], kv[0]))
    rows22 = []
    tl = td = 0.0
    for i, ((name, mat, dia), g) in enumerate(items, 1):
        shown = '%s（%s）' % (name, mat) if mat else name
        rows22.append([str(i), shown, _fmt(g['length']), _fmt(g['detect']),
                       '分%d段(管径%s)' % (g['n'], dia)])
        tl += g['length']
        td += g['detect']

    s22 = '本次共检测%d段管道,评估%d段管道,详见下表:' % (n_total, n_total)
    total22 = {'length': _fmt(tl), 'detect': _fmt(td), 'note': '共%d段' % n_total}
    return s22, rows22, total22


def _num1(v):
    try:
        return round(float(str(v or '').strip().lower().replace('m', '')), 1)
    except ValueError:
        return None


def _row22_key(r):
    pad = (list(r) + ['', '', '', '', ''])[:5]
    return (str(pad[1]).strip(), _num1(pad[2]), _num1(pad[3]), str(pad[4]).strip())


def diff_4122(cur_rows, cur_total, cur_summary, data):
    """平台 4122 vs 生成结果 → (新增行数, 删除行数, 是否一致)。忽略行顺序。"""
    cur = Counter(_row22_key(r) for r in (cur_rows or []))
    new = Counter(_row22_key(r) for r in data['rows22'])
    add = sum((new - cur).values())
    dele = sum((cur - new).values())
    tot = cur_total or {}
    same = (not add and not dele
            and abs((_num1(tot.get('length')) or 0) - (_num1(data['total22']['length']) or 0)) <= 0.05
            and abs((_num1(tot.get('detect')) or 0) - (_num1(data['total22']['detect']) or 0)) <= 0.05
            and str(tot.get('note') or '').strip() == data['total22']['note']
            and str(cur_summary or '').strip() == data['s22'])
    return add, dele, same


def build_from_rows(rows4126, rows4127):
    """返回 dict：4122 工程量、4123 说明/行、4125 说明/行/合计，以及缺陷段明细。"""
    b = {'%s|%s' % (r[1], r[2]): r for r in rows4127 if r and r[1]}
    defects = []
    for r in rows4126:
        if not r or not r[1]:
            continue
        key = None
        for k in b:
            s, e = k.split('|', 1)
            if '%s-%s' % (s, e) == r[1] or '%s-%s' % (e, s) == r[1]:
                key = k
                break
        if not key:
            continue
        rb = b[key]
        slots = [(rb[16 + 5 * i + 1], rb[16 + 5 * i], rb[16 + 5 * i + 3])
                 for i in range(4) if _real(rb[16 + 5 * i + 1])]
        if not slots:
            continue
        st, fn = [], []
        for code, length, grade in slots:
            txt = '纵向%s处,存在%s级%s。' % (length, str(grade).strip(),
                                           CN.get(_code(code), _code(code)))
            (st if _code(code) in STRUCT_CODES else fn).append(txt)
        defects.append({'seg': r[1], 'key': key, 'dia': r[2], 'mat': r[3],
                        'struct': '；'.join(st) if st else '/',
                        'func': '；'.join(fn) if fn else '/',
                        'codes': [_code(c) for c, _, _ in slots]})

    n_total = len([r for r in rows4126 if r and r[1]])
    n_def = len(defects)
    n_s = sum(1 for d in defects if d['struct'] != '/')
    n_f = sum(1 for d in defects if d['func'] != '/')
    n_both = sum(1 for d in defects if d['struct'] != '/' and d['func'] != '/')
    s23 = ('本次共评估%d段管道,其中%d段管道存在缺陷（结构性缺陷管道%d段,功能性缺陷管道%d段,'
           '另有%d段管道同时存在结构性与功能性缺陷）,%d段管道未发现缺陷,详见下表:'
           % (n_total, n_def, n_s, n_f, n_both, n_total - n_def))
    rows23 = [[str(i + 1), d['seg'], d['dia'], d['mat'], d['struct'], d['func']]
              for i, d in enumerate(defects)]

    cnt = Counter()
    for d in defects:
        rb = b[d['key']]
        for i in range(4):
            if _real(rb[16 + 5 * i + 1]):
                cnt[(_code(rb[16 + 5 * i + 1]), str(rb[16 + 5 * i + 3]).strip())] += 1
    rows25 = []
    for code in STRUCT_ORDER + FUNC_ORDER:
        vals = [cnt.get((code, lv), 0) for lv in ('1', '2', '3', '4')]
        rows25.append([code, vals, sum(vals)])
    total25 = [sum(r[1][i] for r in rows25) for i in range(4)]
    n_st = sum(sum(r[1]) for r in rows25 if r[0] in STRUCT_CODES)
    n_fn = sum(sum(r[1]) for r in rows25 if r[0] not in STRUCT_CODES)
    s25 = ('本次共评估%d段管道,其中存在缺陷%d处（结构性缺陷%d处,功能性缺陷%d处）,详见下表:'
           % (n_total, sum(total25), n_st, n_fn))
    s22, rows22, total22 = build_4122(rows4126, rows4127)
    return {'s23': s23, 'rows23': rows23, 's25': s25, 'rows25': rows25,
            'total25': total25, 'defects': defects, 'n_total': n_total,
            's22': s22, 'rows22': rows22, 'total22': total22}


def write_4123(data, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '管道缺陷汇总一览表'
    ws.merge_cells('A1:F1')
    ws['A1'] = data['s23']
    ws['A1'].font = Font(name='新宋体', size=10)
    ws['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    head = ['序号', '管段编号', '管径(mm)', '管段材质', '结构性缺陷', '功能性缺陷']
    for i, h in enumerate(head, 1):
        c = ws.cell(3, i, h)
        c.font = Font(name='新宋体', size=10, bold=True)
        c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')
    for r, row in enumerate(data['rows23'], 4):
        for i, v in enumerate(row, 1):
            c = ws.cell(r, i, v)
            c.font = Font(name='新宋体', size=9)
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col, w in zip('ABCDEF', [6, 14, 10, 12, 42, 42]):
        ws.column_dimensions[col].width = w
    wb.save(path)
    return path


def write_4125(data, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '管道缺陷数量统计表'
    ws.merge_cells('A1:H1')
    ws['A1'] = data['s25']
    ws['A1'].font = Font(name='新宋体', size=10)
    ws['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    head = ['序号', '缺陷类别', '缺陷名称', '1级', '2级', '3级', '4级', '小计']
    for i, h in enumerate(head, 1):
        c = ws.cell(3, i, h)
        c.font = Font(name='新宋体', size=10, bold=True)
        c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')
    for k, (code, vals, sub) in enumerate(data['rows25']):
        row = ['%d' % (k + 1), '结构性缺陷' if code in STRUCT_CODES else '功能性缺陷',
               '(%s)%s' % (code, CN[code])] + list(vals) + [sub]
        for i, v in enumerate(row, 1):
            c = ws.cell(4 + k, i, v)
            c.font = Font(name='新宋体', size=9)
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center')
    tr = 4 + len(data['rows25'])
    for i, v in enumerate(['', '', '合计'] + list(data['total25']) + [sum(data['total25'])], 1):
        c = ws.cell(tr, i, v)
        c.font = Font(name='新宋体', size=9, bold=True)
        c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center')
    for i, w in enumerate([6, 12, 20, 8, 8, 8, 8, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)
    return path


def write_4122(data, path):
    """按平台抓取件的同一版式写 4122（说明行 + 表头 + 明细 + 总计）。"""
    from .crawl import save_4122
    return save_4122(data['s22'], data['rows22'], data['total22'], path)


def write_outputs(data, bgbh, outdir='.', ts=None):
    ts = ts or datetime.datetime.now().strftime('%Y%m%d_%H%M')
    p1 = os.path.join(outdir, '管道缺陷汇总一览表-%s(网页爬取)_%s.xlsx' % (bgbh, ts))
    p2 = os.path.join(outdir, '管道缺陷数量统计表-%s(网页爬取)_%s.xlsx' % (bgbh, ts))
    p3 = os.path.join(outdir, '主要工程量表-%s(生成)_%s.xlsx' % (bgbh, ts))
    write_4123(data, p1)
    write_4125(data, p2)
    write_4122(data, p3)
    return p1, p2, p3

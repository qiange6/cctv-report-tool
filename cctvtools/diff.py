# -*- coding: utf-8 -*-
"""对比：Excel（4126/4127）与平台现状逐字段比对，输出变更清单。"""
from __future__ import annotations

import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .common import EMPTY, same_value

THIN = Side(style='thin', color='FF999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill('solid', fgColor='FFD9E1F2')
OK = PatternFill('solid', fgColor='FFC6EFCE')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

LABEL_4126 = {'c_601': '管段编号', 'c_615': '管径(mm)', 'c_602': '材质', 'c_616': '管道长度(m)',
              'c_620': '水流方向', 'c_617': '管内底标高1', 'c_618': '管内底标高2',
              'c_619': '有无缺陷', 'c_603': '结构性-平均值S', 'c_604': '结构性-最大值Smax',
              'c_605': '结构性-缺陷等级', 'c_606': '结构性-缺陷密度', 'c_607': '结构性-修复指数RI',
              'c_608': '结构性-综合状况评价', 'c_609': '功能性-平均值S',
              'c_610': '功能性-最大值Smax', 'c_611': '功能性-缺陷等级',
              'c_612': '功能性-缺陷密度', 'c_613': '功能性-养护指数MI',
              'c_614': '功能性-综合状况评价'}
LABEL_4127 = {'c_701': '录像文件', 'c_702': '起始井号', 'c_703': '终止井号', 'c_704': '敷设年代',
              'c_705': '起点埋深', 'c_706': '终点埋深', 'c_707': '管段类型', 'c_708': '管段材质',
              'c_709': '管段直径', 'c_710': '检测方向', 'c_711': '管段长度', 'c_712': '检测长度',
              'c_713': '修复指数', 'c_714': '养护指数', 'c_715': '检测人员'}
_SLOT = [['c_716', 'c_717', 'c_718', 'c_719', 'c_720'],
         ['c_729', 'c_730', 'c_731', 'c_732', 'c_733'],
         ['c_735', 'c_737', 'c_739', 'c_741', 'c_743'],
         ['c_736', 'c_738', 'c_740', 'c_742', 'c_744']]
for _i, _codes in enumerate(_SLOT):
    for _j, _c in enumerate(_codes):
        LABEL_4127[_c] = '缺陷%d-%s' % (_i + 1, ['长度', '代码', '分值', '等级', '描述'][_j])


def read_workbook(path):
    """读取 Excel 中的 4126（21 列）与 4127（37 列）两张表。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows26, rows27 = None, None
    for ws in wb.worksheets:
        data = [[('' if c is None else str(c).strip()) for c in r]
                for r in ws.iter_rows(values_only=True)]
        data = [r for r in data if any(r)]
        if not data:
            continue
        head = data[0]
        if rows26 is None and len(head) >= 21 and head[1] == '管段编号':
            rows26 = [r[:21] for r in data[1:]]
        elif rows27 is None and len(head) >= 37 and head[1] == '起始井号':
            rows27 = [r[:37] for r in data[1:]]
    wb.close()
    if rows26 is None or rows27 is None:
        raise RuntimeError('工作表里没找到 4126/4127 数据（需含“管段编号”21 列与“起始井号”37 列）')
    return rows26, rows27


def compare(plat26, plat27, rows26, rows27, alias=None):
    """返回变更列表；plat26/plat27 为 {键: 行}。"""
    alias = alias or {}
    p26 = plat26
    p27 = plat27
    a = {r[1]: r for r in rows26 if r and r[1]}
    b = {'%s|%s' % (r[1], r[2]): r for r in rows27 if r and r[1]}
    changes = []
    for pseg, fseg in alias.items():
        src, cur = a.get(fseg), p26.get(pseg)
        if not src or not cur:
            continue
        for col, code in sorted(_COL26.items()):
            if not same_value(src[col - 1], cur[col - 1]):
                changes.append({'table': '4126', 'seg': pseg, 'newseg': fseg, 'code': code,
                                'field': LABEL_4126[code], 'old': cur[col - 1],
                                'new': src[col - 1]})
    for seg, src in a.items():
        if seg in alias.values():
            continue
        cur = p26.get(seg)
        if not cur:
            continue
        for col, code in sorted(_COL26.items()):
            if not same_value(src[col - 1], cur[col - 1]):
                changes.append({'table': '4126', 'seg': seg, 'newseg': seg, 'code': code,
                                'field': LABEL_4126[code], 'old': cur[col - 1],
                                'new': src[col - 1]})
    for key, src in b.items():
        cur = p27.get(key)
        if not cur:
            continue
        for col, code in sorted(_COL27.items()):
            if not same_value(src[col - 1], cur[col - 1]):
                changes.append({'table': '4127', 'seg': key.replace('|', '-'),
                                'newseg': key.replace('|', '-'), 'code': code,
                                'field': LABEL_4127[code], 'old': cur[col - 1],
                                'new': src[col - 1]})
        for i in range(4):
            for j, code in enumerate(_SLOT[i]):
                if not same_value(src[16 + 5 * i + j], cur[16 + 5 * i + j]):
                    changes.append({'table': '4127', 'seg': key.replace('|', '-'),
                                    'newseg': key.replace('|', '-'), 'code': code,
                                    'field': LABEL_4127[code], 'old': cur[16 + 5 * i + j],
                                    'new': src[16 + 5 * i + j]})
    return changes


_COL26 = {2: 'c_601', 3: 'c_615', 4: 'c_602', 5: 'c_616', 6: 'c_620', 7: 'c_617', 8: 'c_618',
          9: 'c_619', 10: 'c_603', 11: 'c_604', 12: 'c_605', 13: 'c_606', 14: 'c_607',
          15: 'c_608', 16: 'c_609', 17: 'c_610', 18: 'c_611', 19: 'c_612', 20: 'c_613',
          21: 'c_614'}
_COL27 = {2: 'c_702', 3: 'c_703', 4: 'c_701', 5: 'c_704', 6: 'c_705', 7: 'c_706', 8: 'c_707',
          9: 'c_708', 10: 'c_709', 11: 'c_710', 12: 'c_711', 13: 'c_712', 14: 'c_713',
          15: 'c_714', 16: 'c_715'}


def write_list(changes, path, bgbh=''):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '变更清单'
    head = ['序号', '表', '管段编号', '字段', '平台现值', '新值']
    ws.append(head)
    for i, c in enumerate(changes, 1):
        ws.append([i, c['table'], c['newseg'], c['field'], c['old'], c['new']])
    for c in range(1, len(head) + 1):
        cell = ws.cell(1, c)
        cell.font = Font(name='新宋体', size=10, bold=True)
        cell.fill = HEAD
        cell.alignment = CENTER
        cell.border = BORDER
    for r in range(2, len(changes) + 2):
        for c in range(1, len(head) + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name='新宋体', size=9)
            cell.alignment = CENTER
            cell.border = BORDER
    for i, w in enumerate([6, 8, 16, 22, 24, 36], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

    s = wb.create_sheet('汇总')
    from collections import Counter
    c26 = Counter(c['field'] for c in changes if c['table'] == '4126')
    c27 = Counter(c['field'] for c in changes if c['table'] == '4127')
    s.append(['报告编号', bgbh])
    s.append(['4126', len({c['seg'] for c in changes if c['table'] == '4126'}),
              sum(1 for c in changes if c['table'] == '4126')])
    s.append(['4127', len({c['seg'] for c in changes if c['table'] == '4127'}),
              sum(1 for c in changes if c['table'] == '4127')])
    s.append([])
    for k, v in c26.most_common():
        s.append(['4126', k, v])
    for k, v in c27.most_common():
        s.append(['4127', k, v])
    s.column_dimensions['A'].width = 16
    s.column_dimensions['B'].width = 30
    s.column_dimensions['C'].width = 10
    wb.save(path)
    return path

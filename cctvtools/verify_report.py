# -*- coding: utf-8 -*-
"""校验报告：对输出目录里最新一套表做交叉校验，并生成带说明的 xlsx。"""
from __future__ import annotations

import datetime
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import crawl as C

HEAD = PatternFill('solid', fgColor='FFD9E1F2')
PASSF = PatternFill('solid', fgColor='FFC6EFCE')
WARNF = PatternFill('solid', fgColor='FFFFEB9C')
FAILF = PatternFill('solid', fgColor='FFFFC7CE')
THIN = Side(style='thin', color='FF999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

TABLES = ('4126', '4127', '4123', '4125', '4122')

# 检查项说明：id → (这是什么, 不通过怎么办)
CHECK_DOC = {
    'G3-4126-nonempty': ('4126 评估表是否抓到数据', '检查报告编号/权限，重新抓取'),
    'G4-4127-nonempty': ('4127 成果表是否抓到数据', '检查报告编号/权限，重新抓取'),
    'G3-no-dup-4126': ('4126 是否有重复管段', '按管段编号去重后重抓'),
    'G4-no-dup-4127': ('4127 是否有重复管段', '检查翻页是否重复，重新抓取'),
    'G5-count-match': ('4126 与 4127 管段数是否一致', '数量不一致说明有一边缺行，先查缺哪段'),
    'G5-direction-flip': ('两表起止井号是否反向', '反向属平台数据问题，确认检测方向后统一'),
    'G5-unmatched': ('两表管段编号是否对不上', '找出不匹配段，核对井号命名'),
    'G6-defect-slots': ('单根管段缺陷是否超过 4 条', '超过 4 条会被截断，需扩槽位或另存'),
    'G7-4123-vs-4127': ('4123 汇总表段数与 4127 缺陷段数是否一致', '汇总表漏段/多段，需按 4127 重新生成'),
    'G7b-4123-in-4126': ('4123 里的管段能否在 4126 找到', '汇总表出现评估表没有的管段，核对编号'),
    'G7b-4123-flip': ('4123 与 4126 起止井号反向', '确认方向后统一'),
    'G7c-4123-type': ('4123 描述的结构/功能分类与统计表是否一致', '分类错位，按 4127 缺陷代码修正'),
    'G7-4123-summary': ('4123 说明行段数与表体行数是否一致', '说明行是平台写死的文字，与数据不符需更新'),
    'G8-4125-vs-4123': ('4125 合计与 4123 描述条数是否一致', '统计表与汇总表对不上，重新生成'),
    'G8-4125-vs-4127': ('4125 合计与 4127 缺陷条目数是否一致', '成果表槽位可能漏读，检查 4127'),
    'G8-4125-summary': ('4125 说明行处数与合计行是否一致', '说明行文字过期，需更新'),
    'G8c-4125-type': ('4125 分类与 4127 缺陷代码分类是否一致', '缺陷类型归属错，按代码修正'),
    'G8-unknown-code': ('是否存在未识别的缺陷代码', '补充代码映射表'),
    'G9-4122-segments': ('4122 备注“分N段”与 4126 行数是否一致', '工程量表与评估表脱节，核对分段'),
    'G9-4122-length': ('4122 总计长度与 4126 长度合计是否一致', '差得多说明有一边长度录错'),
    'G9-4122-summary': ('4122 说明行段数与 4126 行数是否一致', '说明行文字过期，需更新'),
}


def _level_fill(level):
    return {'PASS': PASSF, 'WARN': WARNF, 'FAIL': FAILF}.get(level, None)


def _file_info(path):
    if not path or not os.path.exists(path):
        return '', ''
    st = os.stat(path)
    return ('%.1f KB' % (st.st_size / 1024.0),
            datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M'))


def build(bgbh, out_dir='.'):
    """执行校验并生成 xlsx，返回 dict(path, verdict, checks, files)。"""
    rows, recs, extra, files = C.load_outputs(bgbh, out_dir)
    if not files.get('4126') or not files.get('4127'):
        raise RuntimeError('输出目录里没有 %s 的最新 4126/4127，请先“只抓取”或“一键全流程”' % bgbh)
    v = C.verify(rows, recs, extra)
    checks = v['checks']
    verdict = ('FAIL' if any(c['level'] == 'FAIL' for c in checks)
               else ('WARN' if any(c['level'] == 'WARN' for c in checks) else 'PASS'))

    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    path = os.path.join(out_dir, '校验报告-%s-%s.xlsx' % (bgbh, ts))

    wb = openpyxl.Workbook()
    # ---------- 1. 校验结论 ----------
    ws = wb.active
    ws.title = '校验结论'
    pairs = [
        ('报告编号', bgbh),
        ('校验时间', datetime.datetime.now().strftime('%Y-%m-%d %H:%M')),
        ('输出目录', os.path.abspath(out_dir)),
        ('总体结论', {'PASS': '通过', 'WARN': '通过（有提醒）', 'FAIL': '不通过'}[verdict]),
        ('检查项', '共 %d 项：PASS %d / WARN %d / FAIL %d' % (
            len(checks),
            sum(1 for c in checks if c['level'] == 'PASS'),
            sum(1 for c in checks if c['level'] == 'WARN'),
            sum(1 for c in checks if c['level'] == 'FAIL'))),
        ('4126 评估表', '%d 行' % len(rows)),
        ('4127 成果表', '%d 段，缺陷 %d 段 / %d 条' % (
            len(recs), v.get('defect_pipes', 0), v.get('defect_entries', 0))),
        ('4123 汇总表', '%d 行' % len((extra.get('4123') or {}).get('rows') or [])),
        ('4125 统计表', '合计 %s 处' % ((extra.get('4125') or {}).get('total') or {}).get('小计', '-')),
        ('4122 工程量表', '%d 行' % len((extra.get('4122') or {}).get('rows') or [])),
        ('结论说明', '通过=两表及派生表口径一致；有提醒=存在流向/分类等需人工确认项；'
                     '不通过=数量或内容矛盾，需修复后再上传'),
    ]
    ws.append(['项目', '内容'])
    for k, val in pairs:
        ws.append([k, val])
    for c in (1, 2):
        cell = ws.cell(1, c)
        cell.font = Font(name='新宋体', size=10, bold=True)
        cell.fill = HEAD
        cell.alignment = CENTER
        cell.border = BORDER
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).font = Font(name='新宋体', size=10, bold=True)
        ws.cell(r, 2).font = Font(name='新宋体', size=10)
        ws.cell(r, 1).border = BORDER
        ws.cell(r, 2).border = BORDER
        ws.cell(r, 2).alignment = LEFT
    ws.cell(5, 2).fill = _level_fill(verdict)
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 80

    # ---------- 2. 检查明细 ----------
    ws = wb.create_sheet('检查明细')
    ws.append(['序号', '检查项', '级别', '结果', '含义', '不通过怎么办'])
    for i, c in enumerate(checks, 1):
        doc = CHECK_DOC.get(c['id'], ('', ''))
        ws.append([i, c['id'], c['level'], c['msg'], doc[0], doc[1]])
    for c in range(1, 7):
        cell = ws.cell(1, c)
        cell.font = Font(name='新宋体', size=10, bold=True)
        cell.fill = HEAD
        cell.alignment = CENTER
        cell.border = BORDER
    for r in range(2, len(checks) + 2):
        lv = ws.cell(r, 3).value
        for c in range(1, 7):
            cell = ws.cell(r, c)
            cell.font = Font(name='新宋体', size=9)
            cell.alignment = CENTER if c != 4 else LEFT
            cell.border = BORDER
        ws.cell(r, 3).fill = _level_fill(lv) or PASSF
    for i, w in enumerate([5, 22, 8, 44, 34, 36], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

    # ---------- 3. 问题清单 ----------
    ws = wb.create_sheet('问题清单')
    ws.append(['序号', '检查项', '级别', '问题描述', '建议处理'])
    n = 0
    for c in checks:
        if c['level'] == 'PASS':
            continue
        n += 1
        doc = CHECK_DOC.get(c['id'], ('', ''))
        ws.append([n, c['id'], c['level'], c['msg'], doc[1] or '人工复核'])
    if n == 0:
        ws.append(['', '', '', '未发现问题', ''])
    for c in range(1, 6):
        cell = ws.cell(1, c)
        cell.font = Font(name='新宋体', size=10, bold=True)
        cell.fill = HEAD
        cell.alignment = CENTER
        cell.border = BORDER
    for r in range(2, ws.max_row + 1):
        lv = ws.cell(r, 3).value
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.font = Font(name='新宋体', size=9)
            cell.alignment = CENTER if c in (1, 2, 3) else LEFT
            cell.border = BORDER
        if lv:
            ws.cell(r, 3).fill = _level_fill(lv)
    for i, w in enumerate([5, 22, 8, 60, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---------- 4. 文件清单 ----------
    ws = wb.create_sheet('文件清单')
    ws.append(['表', '文件名', '大小', '修改时间'])
    for t in TABLES:
        p = files.get(t)
        kb, mt = _file_info(p)
        ws.append([t, os.path.basename(p) if p else '（未找到）', kb, mt])
    for c in range(1, 5):
        cell = ws.cell(1, c)
        cell.font = Font(name='新宋体', size=10, bold=True)
        cell.fill = HEAD
        cell.alignment = CENTER
        cell.border = BORDER
    for r in range(2, ws.max_row + 1):
        for c in range(1, 5):
            cell = ws.cell(r, c)
            cell.font = Font(name='新宋体', size=9)
            cell.alignment = CENTER if c != 2 else LEFT
            cell.border = BORDER
        if not ws.cell(r, 2).value or ws.cell(r, 2).value.startswith('（'):
            ws.cell(r, 2).fill = FAILF
    for i, w in enumerate([8, 60, 12, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)
    return {'path': path, 'verdict': verdict, 'checks': checks, 'files': files}

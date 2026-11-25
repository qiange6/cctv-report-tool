# -*- coding: utf-8 -*-
"""批量模式：逐个报告跑工作流，汇总成结果表。"""
from __future__ import annotations

import contextlib
import datetime
import io
import json
import os
import sys
import traceback

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import workflow as WF

HEAD = PatternFill('solid', fgColor='FFD9E1F2')
OKF = PatternFill('solid', fgColor='FFC6EFCE')
BADF = PatternFill('solid', fgColor='FFFFC7CE')
WARNF = PatternFill('solid', fgColor='FFFFEB9C')
THIN = Side(style='thin', color='FF999999')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


class Tee:
    """同时写屏幕与日志文件。"""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
                st.flush()
            except Exception:
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except Exception:
                pass


def read_list(path):
    """一行一个报告编号，忽略空行与 # 注释。"""
    out = []
    with open(path, encoding='utf-8-sig') as f:
        for line in f:
            s = line.strip()
            if not s or s.startswith('#'):
                continue
            # 允许“编号,工程名”这种写法
            out.append(s.split(',')[0].strip())
    return out


def _fail_warn(checks):
    f = sum(1 for c in checks if c['level'] == 'FAIL')
    w = sum(1 for c in checks if c['level'] == 'WARN')
    return f, w


def prompt_upload(bgbh):
    """逐报告确认：y=上传 / n=跳过 / a=后面全部上传 / q=退出批量。"""
    try:
        ans = input('  报告 %s：上传 4123/4125 到平台？[y=是 / n=否 / a=全部 / q=退出] '
                    % bgbh).strip().lower()
    except EOFError:
        return 'n'
    if ans in ('y', 'n', 'a', 'q'):
        return ans
    return 'n'


def run_batch(bgbh_list, out_dir='.', upload=False, dry=False, yes=False,
              aliases=None, from_xlsx=None, confirm_each=False, log=print):
    """逐个跑；单个失败不影响后续。返回结果列表。"""
    aliases = aliases or {}
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    root = os.path.join(out_dir, '批量_%s' % ts)
    os.makedirs(root, exist_ok=True)
    results = []
    log('批量开始：%d 个报告，输出目录 %s' % (len(bgbh_list), root))
    log('模式：%s%s%s' % ('演练（不写平台）' if dry else '正常',
                        '，上传' if upload else '，不上传',
                        '，逐报告确认' if (confirm_each and upload) else ''))
    upload_all = False

    for i, bgbh in enumerate(bgbh_list, 1):
        log('\n' + '#' * 60)
        log('# [%d/%d] %s' % (i, len(bgbh_list), bgbh))
        log('#' * 60)
        sub = os.path.join(root, bgbh)
        os.makedirs(sub, exist_ok=True)
        logpath = os.path.join(sub, '%s.log' % bgbh)
        row = {'bgbh': bgbh, 'status': '', 'error': '', 'gcmc': '',
               'n4126': '', 'n4127': '', 'n4123': '', 'n4125': '',
               'fail': '', 'warn': '', 'uploaded': '', 'recheck': ''}
        do_upload = upload
        if upload and confirm_each and not upload_all:
            ans = prompt_upload(bgbh)
            if ans == 'q':
                log('  用户选择退出批量，剩余报告不再执行')
                row['status'] = '未执行(退出)'
                results.append(row)
                break
            if ans == 'a':
                upload_all = True
            elif ans != 'y':
                do_upload = False
        try:
            with open(logpath, 'w', encoding='utf-8') as f:
                tee = Tee(sys.stdout, f)
                with contextlib.redirect_stdout(tee):
                    res = WF.run(bgbh, out_dir=sub, upload=do_upload, dry=dry, yes=yes,
                                 from_xlsx=from_xlsx, alias=aliases.get(bgbh, {}))
            st = res.get('stages', {})
            row.update({
                'gcmc': (res.get('report') or {}).get('gcmc', ''),
                'n4126': (st.get('4126') or {}).get('rows', ''),
                'n4127': (st.get('4127') or {}).get('records', ''),
                'n4123': (st.get('4123') or {}).get('rows', ''),
                'n4125': (res.get('totals') or {}).get('4125', (st.get('4125') or {}).get('rows', '')),
                'uploaded': ('跳过(人工)' if (upload and not do_upload)
                             else ('、'.join(res.get('uploaded') or []) or '无')),
                'recheck': ('一致' if (res.get('recheck') or {}).get('ok')
                            else ('未做' if not res.get('recheck') else '有差异')),
                'status': '成功',
            })
            f_, w_ = _fail_warn(res.get('checks') or [])
            row['fail'], row['warn'] = f_, w_
            if f_:
                row['status'] = '成功(有FAIL)'
        except Exception as e:
            row['status'] = '失败'
            row['error'] = '%s: %s' % (type(e).__name__, str(e)[:200])
            log('\n[失败] %s -> %s' % (bgbh, row['error']))
            with open(logpath, 'a', encoding='utf-8') as f:
                f.write('\n' + traceback.format_exc())
        results.append(row)

    _write_summary(results, root, ts, dry, upload, log)
    return results


def _write_summary(results, root, ts, dry, upload, log):
    xlsx = os.path.join(root, '批量结果_%s.xlsx' % ts)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '批量结果'
    head = ['序号', '报告编号', '工程名称', '4126行', '4127段', '4123行', '4125合计',
            '校验FAIL', '校验WARN', '上传', '复核', '状态', '错误信息']
    ws.append(head)
    for i, r in enumerate(results, 1):
        ws.append([i, r['bgbh'], r['gcmc'], r['n4126'], r['n4127'], r['n4123'], r['n4125'],
                   r['fail'], r['warn'], r['uploaded'], r['recheck'], r['status'], r['error']])
    for c in range(1, len(head) + 1):
        cell = ws.cell(1, c)
        cell.font = Font(name='新宋体', size=10, bold=True)
        cell.fill = HEAD
        cell.alignment = CENTER
        cell.border = BORDER
    for r in range(2, len(results) + 2):
        status = ws.cell(r, 12).value or ''
        fill = OKF if status == '成功' else (BADF if status == '失败' else WARNF)
        for c in range(1, len(head) + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name='新宋体', size=9)
            cell.alignment = CENTER
            cell.border = BORDER
        ws.cell(r, 12).fill = fill
    for i, w in enumerate([5, 14, 34, 8, 8, 8, 8, 9, 9, 12, 8, 12, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'C2'
    wb.save(xlsx)

    json_path = os.path.join(root, '批量结果_%s.json' % ts)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump({'run_at': ts, 'dry': dry, 'upload': upload, 'results': results},
                  f, ensure_ascii=False, indent=2)

    ok = sum(1 for r in results if r['status'].startswith('成功'))
    bad = len(results) - ok
    log('\n' + '=' * 60)
    log('批量完成：成功 %d / 失败 %d' % (ok, bad))
    log('结果表：%s' % xlsx)
    log('明细 JSON：%s' % json_path)
    log('每个报告的日志与产物在各子目录里。')
    return xlsx

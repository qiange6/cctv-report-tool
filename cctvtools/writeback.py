# -*- coding: utf-8 -*-
"""写回：按“修改→填值→保存”逐页提交（HTTP，无需浏览器）。"""
from __future__ import annotations

import json
import time

from .common import BASE, IGID, parse_form, put, same_value

# xlsx 列(1基) → 编辑页字段前缀
COLMAP_4126 = {2: 'c_601', 3: 'c_615', 4: 'c_602', 5: 'c_616', 6: 'c_620',
               7: 'c_617', 8: 'c_618', 9: 'c_619', 10: 'c_603', 11: 'c_604',
               12: 'c_605', 13: 'c_606', 14: 'c_607', 15: 'c_608', 16: 'c_609',
               17: 'c_610', 18: 'c_611', 19: 'c_612', 20: 'c_613', 21: 'c_614'}
COLMAP_4127 = {2: 'c_702', 3: 'c_703', 4: 'c_701', 5: 'c_704', 6: 'c_705',
               7: 'c_706', 8: 'c_707', 9: 'c_708', 10: 'c_709', 11: 'c_710',
               12: 'c_711', 13: 'c_712', 14: 'c_713', 15: 'c_714', 16: 'c_715'}
SLOT_4127 = [['c_716', 'c_717', 'c_718', 'c_719', 'c_720'],
             ['c_729', 'c_730', 'c_731', 'c_732', 'c_733'],
             ['c_735', 'c_737', 'c_739', 'c_741', 'c_743'],
             ['c_736', 'c_738', 'c_740', 'c_742', 'c_744']]


def open_edit(sess, bid, pg):
    sess.get(sess.report_url())
    sess.get(sess.view_url(bid, pg))
    sess.get(sess.anniu_url(bid))
    j = json.loads(sess.post(BASE + '/XiangMuCaoZuo',
                             {'ID': '%d$%s$%s$%s$%s$' % (pg, sess.igid, bid, sess.pid, sess.rid),
                              'addPara': 'true', 'Type': '4', 'locking': 'true'},
                             ref=sess.report_url()))
    if not j.get('success'):
        raise RuntimeError('进入编辑态失败: %s' % j)
    return sess.get(BASE + '/system/' + j['url'], encoding='gbk')


def save_page(sess, bid, pg, form):
    resp = sess.post(BASE + '/XiangMuXiuGaiBaoCun2', [tuple(x) for x in form],
                     ref=sess.edit_url(bid, pg))
    return 'XiangMuXianShi.jsp' in resp


def apply_page(sess, bid, pg, row_fields, dry=False, expect=None):
    """row_fields: {行号: {字段: 值}}；expect: 行号 → 期望管段编号（防错位）。"""
    form = parse_form(open_edit(sess, bid, pg))
    idx = {n: i for i, (n, _) in enumerate(form)}
    if expect:
        if bid == '4126':
            for row, seg in expect.items():
                nm = 'c_601_%d' % row
                if nm in idx and form[idx[nm]][1] != seg:
                    raise RuntimeError('行号错位：第%d行是 %r，期望 %r'
                                       % (row, form[idx[nm]][1], seg))
        elif 'c_702_1' in idx:
            s_, e_ = expect['seg'].split('|')
            if form[idx['c_702_1']][1] != s_ or form[idx['c_703_1']][1] != e_:
                raise RuntimeError('段错位：编辑页是 %s|%s，期望 %s'
                                   % (form[idx['c_702_1']][1], form[idx['c_703_1']][1], expect['seg']))
    changed, miss = 0, []
    for row, fields in row_fields.items():
        for code, val in fields.items():
            name = '%s_%d' % (code, row)
            if name not in idx:
                miss.append(name)
                continue
            if not same_value(form[idx[name]][1], val):
                form[idx[name]][1] = str(val)
                changed += 1
    if miss:
        raise RuntimeError('编辑页缺少字段: %s' % miss[:8])
    if dry:
        print('   第%s行 待改 %d 个字段（演练不保存）' % (list(row_fields), changed))
        return 'DRY'
    ok = save_page(sess, bid, pg, form)
    print('   保存%s（变更 %d）' % ('成功' if ok else '失败', changed))
    if not ok:
        raise RuntimeError('保存失败 %s 页%d' % (bid, pg))
    return 'OK'


def diff_rows(rows_4126, recs_4127, plat_4126, plat_4127, alias=None):
    """比较“文件”与“平台”，返回 (plan26, plan27)。

    rows_4126: 文件 21 列行；recs_4127: 文件 37 列行
    plat_4126: {管段编号: 行}；plat_4127: {(起|止): 行}
    alias: {'平台编号': '文件编号'}（处理编号纠正，如 Y16-Y17 → Y15-Y17）
    """
    alias = alias or {}
    a = {r[1]: r for r in rows_4126}
    b = {'%s|%s' % (r[1], r[2]): r for r in recs_4127}
    plan26, plan27 = {}, {}
    for pseg, fseg in alias.items():
        src = a.get(fseg)
        cur = plat_4126.get(pseg)
        if not src or not cur:
            continue
        fields = {code: src[col - 1] for col, code in COLMAP_4126.items()
                  if not same_value(src[col - 1], cur[col - 1])}
        if fields:
            plan26[pseg] = fields
    for seg, src in a.items():
        if seg in alias.values():
            continue
        cur = plat_4126.get(seg)
        if not cur:
            continue
        fields = {code: src[col - 1] for col, code in COLMAP_4126.items()
                  if not same_value(src[col - 1], cur[col - 1])}
        if fields:
            plan26[seg] = fields
    for key, src in b.items():
        cur = plat_4127.get(key)
        if not cur:
            continue
        fields = {code: src[col - 1] for col, code in COLMAP_4127.items()
                  if not same_value(src[col - 1], cur[col - 1])}
        for i in range(4):
            for j, code in enumerate(SLOT_4127[i]):
                if not same_value(src[16 + 5 * i + j], cur[16 + 5 * i + j]):
                    fields[code] = src[16 + 5 * i + j]
        if fields:
            plan27[key] = fields
    return plan26, plan27


def scan_4126_positions(sess, pages=13):
    """扫描 4126 编辑页，返回 {管段编号: (页, 行)}。"""
    pos = {}
    for pg in range(pages):
        form = dict(parse_form(open_edit(sess, '4126', pg)))
        n = len([k for k in form if k.endswith('_1') and k.startswith('c_601_')])
        rows = [r for r in range(1, 14) if 'c_601_%d' % r in form]
        for r in rows:
            seg = form.get('c_601_%d' % r)
            if seg:
                pos[seg] = (pg, r)
    return pos


def upload(sess, bid, plan, positions, page_size=None, dry=False, progress=None):
    """plan: {段: {字段:值}}；positions: {段:(页,行)} 或 None（4127 按行序）。"""
    pages = {}
    for seg, fields in plan.items():
        pg, row = positions[seg]
        pages.setdefault(pg, {}).setdefault(row, {}).update(fields)
    done = []
    for pg in sorted(pages):
        print('== 写 %s 页%d ==' % (bid, pg))
        apply_page(sess, bid, pg, pages[pg], dry=dry)
        done.append(pg)
        if progress:
            progress(bid, pg)
        if not dry:
            time.sleep(0.3)
    return done


# ==================== 派生表（4123 / 4125）写回 ====================
def upload_4123(sess, summary, rows, dry=False, page_size=12):
    """rows: [[序号,管段编号,管径,材质,结构性,功能性], ...]（自动分页、空行 Recid=0 会新增）。"""
    pages = max(1, (len(rows) + page_size - 1) // page_size)
    for pg in range(pages):
        form = parse_form(open_edit(sess, '4123', pg))
        changed = 0
        if pg == 0:
            changed += put(form, 'c_306_1', summary)
        for i in range(1, page_size + 1):
            idx = pg * page_size + i - 1
            row = rows[idx] if idx < len(rows) else ['', '', '', '', '', '']
            for code, val in (('c_301', row[1]), ('c_302', row[2]), ('c_303', row[3]),
                              ('c_304', row[4]), ('c_305', row[5])):
                changed += put(form, '%s_%d' % (code, i), val)
        if dry:
            print('   4123 页%d 待改 %d 个字段（演练）' % (pg, changed))
            continue
        ok = save_page(sess, '4123', pg, form)
        print('   4123 页%d 保存=%s（改 %d）' % (pg, ok, changed))
        if not ok:
            raise RuntimeError('4123 页%d 保存失败' % pg)
        time.sleep(0.3)


def upload_4125(sess, summary, rows16, total, dry=False):
    """rows16: [(code, [1级,2级,3级,4级], 小计), ...] 共 16 项；total: [1级,2级,3级,4级]。"""
    form = parse_form(open_edit(sess, '4125', 0))
    changed = put(form, 'c_501_1', summary)
    for r, (code, vals, sub) in enumerate(rows16, 1):
        for j, v in enumerate(vals, 1):
            changed += put(form, 'c_50%d_%d' % (1 + j, r), v)
        changed += put(form, 'c_506_%d' % r, sub)
    for j, v in enumerate(total, 1):
        changed += put(form, 'c_50%d_1' % (6 + j), v)
    changed += put(form, 'c_511_1', sum(total))
    if dry:
        print('   4125 待改 %d 个字段（演练）' % changed)
        return
    ok = save_page(sess, '4125', 0, form)
    print('   4125 保存=%s（改 %d）' % (ok, changed))
    if not ok:
        raise RuntimeError('4125 保存失败')


def upload_4122(sess, summary, rows, total, dry=False, page_size=10):
    """rows: [[序号,名称,管道长度,检测长度,备注], ...]，平台每页 10 行。

    字段：c_205_1=说明；c_201..c_204_行=名称/管道长度/检测长度/备注；
          c_206_1/c_207_1=总计长度/检测长度；c_208_1=总计备注（共N段）。
    """
    pages = max(1, (len(rows) + page_size - 1) // page_size)
    for pg in range(pages):
        form = parse_form(open_edit(sess, '4122', pg))
        have = {n for n, _ in form}
        changed = 0
        if pg == 0:
            for name in ('c_205_1', 'c_206_1', 'c_207_1', 'c_208_1'):
                if name not in have:
                    raise RuntimeError('4122 编辑页缺少字段: %s' % name)
            changed += put(form, 'c_205_1', summary)
            changed += put(form, 'c_206_1', total.get('length'))
            changed += put(form, 'c_207_1', total.get('detect'))
            changed += put(form, 'c_208_1', total.get('note'))

        miss = []
        for i in range(1, page_size + 1):
            n = pg * page_size + i
            row = list(rows[n - 1]) if n <= len(rows) else []
            row = (row + ['', '', '', '', ''])[:5]
            for code, val in (('c_201', row[1]), ('c_202', row[2]),
                              ('c_203', row[3]), ('c_204', row[4])):
                name = '%s_%d' % (code, i)
                if name not in have:
                    miss.append(name)
                else:
                    changed += put(form, name, val)
        if miss:
            raise RuntimeError('4122 编辑页缺少字段: %s' % miss[:8])

        if dry:
            print('   4122 页%d 待改 %d 个字段（演练）' % (pg, changed))
            continue
        ok = save_page(sess, '4122', pg, form)
        print('   4122 页%d 保存=%s（改 %d）' % (pg, ok, changed))
        if not ok:
            raise RuntimeError('4122 页%d 保存失败' % pg)
        time.sleep(0.3)

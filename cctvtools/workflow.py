# -*- coding: utf-8 -*-
"""一键工作流：定位 → 抓取五张表 → 校验 → 生成 4123/4125 → 对比 → 上传 → 复核。

对外只暴露一个入口：run(bgbh, ...)。命令行见 cli.py 的 workflow 子命令。
"""
from __future__ import annotations

import datetime
import json
import os
import sys

import openpyxl

from . import crawl as C
from . import diff as D
from . import generate as G
from . import writeback as WB
from .common import EMPTY, same_value

TABLES = ('4122', '4123', '4125', '4126', '4127')


def _ask(prompt):
    try:
        return input(prompt).strip().lower() in ('y', 'yes', '是')
    except EOFError:
        return False


def _same_row(a, b, n):
    return all(same_value(a[i] if i < len(a) else '', b[i] if i < len(b) else '')
               for i in range(n))


def _diff_4123(cur_rows, new_rows):
    """返回 (新增, 修改, 删除) 的简要说明。"""
    cur = {(r[1], r[4], r[5]) for r in cur_rows if r and r[1]}
    new = {(r[1], r[4], r[5]) for r in new_rows if r and r[1]}
    return len(new - cur), 0, len(cur - new)


def _diff_4125(cur_defs, cur_total, new_rows, new_total):
    cur = {d['缺陷名称']: (d['1级'], d['2级'], d['3级'], d['4级'], d['小计'])
           for d in (cur_defs or [])}
    changed = 0
    for code, vals, sub in new_rows:
        name = '(%s)%s' % (code, G.CN[code])
        cv = cur.get(name)
        if cv is None or [int(x) for x in cv] != list(vals) + [sub]:
            changed += 1
    return changed, cur_total, new_total


def run(bgbh, out_dir='.', upload=True, dry=False, yes=False, from_xlsx=None,
        alias=None, log=print):
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    os.makedirs(out_dir, exist_ok=True)
    result = {'bgbh': bgbh, 'run_at': datetime.datetime.now().isoformat(timespec='seconds'),
              'stages': {}, 'files': [], 'checks': [], 'uploaded': []}

    log('\n' + '=' * 60)
    log('CCTV 工作流  %s' % bgbh)
    log('=' * 60)

    # ---------- S1 定位 ----------
    cfg, meta = C.load_config()
    C.print_config(cfg, meta, bgbh)
    sess = C.Session(cfg, verbose=False).login()
    rep = C.locate_report(sess, bgbh, cfg['IGID'], days=3650, verbose=False)
    sess.bgbh, sess.rid, sess.pid, sess.igid = bgbh, rep['report_id'], str(rep['parentid']), cfg['IGID']
    log('[S1] report_id=%s parentid=%s 工程=%s' % (sess.rid, sess.pid, rep.get('gcmc', '')))
    result['report'] = rep

    # ---------- S2 抓取五张表 ----------
    log('\n[S2] 抓取五张表 …')
    rows26, st26 = C.crawl_4126(sess, rep, verbose=True)
    recs27, st27 = C.crawl_4127(sess, rep, verbose=True)
    s23, rows23, st23 = C.crawl_4123(sess, rep, verbose=False)
    s25, defs25, tot25, st25 = C.crawl_4125(sess, rep, verbose=False)
    s22, rows22, tot22, st22 = C.crawl_4122(sess, rep, verbose=False)
    log('   4126 %d 行 / 4127 %d 段 / 4123 %d 行 / 4125 合计 %s / 4122 %d 行'
        % (len(rows26), len(recs27), len(rows23), (tot25 or {}).get('小计'), len(rows22)))

    files = {
        '4126': C.save_4126(rows26, os.path.join(out_dir, C.output_name('4126', bgbh, ts))),
        '4127': C.save_4127(recs27, os.path.join(out_dir, C.output_name('4127', bgbh, ts))),
        '4123': C.save_4123(s23, rows23, os.path.join(out_dir, C.output_name('4123', bgbh, ts))),
        '4125': C.save_4125(s25, defs25, tot25, os.path.join(out_dir, C.output_name('4125', bgbh, ts))),
        '4122': C.save_4122(s22, rows22, tot22, os.path.join(out_dir, C.output_name('4122', bgbh, ts))),
    }
    for t in TABLES:
        log('   %s → %s' % (t, os.path.basename(files[t])))
    result['files'] = list(files.values())
    result['stages'] = {'4126': st26, '4127': st27, '4123': st23, '4125': st25, '4122': st22}

    # ---------- S3 校验 ----------
    log('\n[S3] 交叉校验 …')
    v = C.verify(rows26, recs27, extra={
        '4123': {'summary': s23, 'rows': rows23},
        '4125': {'summary': s25, 'defects': defs25, 'total': tot25},
        '4122': {'summary': s22, 'rows': rows22, 'total': tot22,
                 'segments_in_note': st22.get('segments_in_note', 0)},
    })
    for c in v['checks']:
        log('   [%s] %s' % (c['level'], c['msg']))
    result['checks'] = v['checks']
    if v['hard_fail']:
        log('   ⚠ 校验存在 FAIL，继续生成，但上传前请确认')

    # ---------- S4 生成 4122 / 4123 / 4125 ----------
    log('\n[S4] 生成 4122 工程量表 / 4123 汇总表 / 4125 统计表 …')
    if from_xlsx:
        src26, src27 = D.read_workbook(from_xlsx)
        log('   数据源：%s' % os.path.basename(from_xlsx))
    else:
        src26 = rows26
        src27 = []
        # 用刚抓的 4127 转成 37 列行
        tmp = os.path.join(out_dir, '_tmp_4127.xlsx')
        C.save_4127(recs27, tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        for r in ws.iter_rows(min_row=2, values_only=True):
            src27.append([('' if c is None else str(c).strip()) for c in r])
        wb.close()
        os.remove(tmp)
        log('   数据源：本次抓取的 4126 + 4127')
    data = G.build_from_rows(src26, src27)
    p23, p25, p22 = G.write_outputs(data, bgbh, out_dir, ts)
    log('   4122 %d 行：%s' % (len(data['rows22']), data['s22']))
    log('   4123 %d 行：%s' % (len(data['rows23']), data['s23']))
    log('   4125 合计 %d：%s' % (sum(data['total25']), data['s25']))
    result['generated'] = {'4122': p22, '4123': p23, '4125': p25,
                           'summary22': data['s22'],
                           'summary23': data['s23'], 'summary25': data['s25']}
    result['totals'] = {'4125': sum(data['total25'])}
    result['files'] += [p22, p23, p25]

    # ---------- S5 对比 + 上传 ----------
    log('\n[S5] 对比平台现状 …')
    d22 = G.diff_4122(rows22, tot22, s22, data)
    d23 = _diff_4123(rows23, data['rows23'])
    ch25 = _diff_4125(defs25, tot25, data['rows25'], data['total25'])
    need22 = not d22[2]
    need23 = d23 != (0, 0, 0)
    need25 = ch25[0] > 0
    log('   4122 新增 %d / 删除 %d / %s'
        % (d22[0], d22[1], '与平台一致' if not need22 else '有差异'))
    log('   4123 新增 %d / 修改 %d / 删除 %d' % d23)
    log('   4125 变化行 %d 行（当前合计 %s → 新合计 %d）'
        % (ch25[0], (ch25[1] or {}).get('小计'), sum(data['total25'])))
    need = need22 or need23 or need25
    result['diff'] = {'4122': [d22[0], d22[1]], '4123': d23, '4125': ch25[0]}

    if upload and need:
        go = yes or _ask('是否上传 4122/4123/4125 到平台？[y/N] ')
        if go and dry:
            if need22:
                log('   [演练] 将写回 4122 %d 行（不实际提交）' % len(data['rows22']))
            if need23:
                log('   [演练] 将写回 4123 %d 行（不实际提交）' % len(data['rows23']))
            if need25:
                log('   [演练] 将写回 4125 16 行（不实际提交）')
        elif go:
            sent = []
            if need22:
                log('   上传 4122 …')
                WB.upload_4122(sess, data['s22'], data['rows22'], data['total22'])
                sent.append('4122')
            if need23:
                log('   上传 4123 …')
                WB.upload_4123(sess, data['s23'], data['rows23'])
                sent.append('4123')
            if need25:
                log('   上传 4125 …')
                WB.upload_4125(sess, data['s25'], data['rows25'], data['total25'])
                sent.append('4125')
            result['uploaded'] += sent
    elif upload:
        log('   平台与生成结果一致，无需上传')

    # 若提供了修改后的 4126/4127，也一并写回
    if from_xlsx and upload:
        log('\n[S5b] 写回 4126/4127 …')
        plat26 = {r[1]: r for r in rows26}
        tmp = os.path.join(out_dir, '_tmp_p27.xlsx')
        C.save_4127(recs27, tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        p27 = [[('' if c is None else str(c).strip()) for c in r] for r in ws.iter_rows(values_only=True)]
        wb.close()
        os.remove(tmp)
        plat27 = {'%s|%s' % (r[1], r[2]): r for r in p27[1:] if r[1]}
        plan26, plan27 = WB.diff_rows(src26, src27, plat26, plat27, alias or {})
        log('   计划：4126 %d 段 / 4127 %d 段' % (len(plan26), len(plan27)))
        if (plan26 or plan27) and (yes or _ask('是否写回 4126/4127？[y/N] ')):
            if dry:
                log('   [演练] 将写回 4126 %d 段、4127 %d 段（不实际提交）'
                    % (len(plan26), len(plan27)))
            elif plan26:
                pages = max(1, int((st26 or {}).get('pages') or 1))
                pos = WB.scan_4126_positions(sess, pages=pages)
                WB.upload(sess, '4126', plan26, pos, dry=dry)
            if plan27:
                order = {'%s|%s' % (r[1], r[2]): i for i, r in enumerate(p27[1:]) if r[1]}
                pos = {k: (order[k], 1) for k in plan27 if k in order}
                WB.upload(sess, '4127', plan27, pos, dry=dry)
            if not dry:
                result['uploaded'] += ['4126', '4127']

    # ---------- S6 复核 ----------
    if result['uploaded'] and not dry:
        log('\n[S6] 复核（重新抓取）…')
        s2 = C.Session(cfg, verbose=False).login()
        rep2 = C.locate_report(s2, bgbh, cfg['IGID'], days=3650, verbose=False)
        s22b, r22b, t22b, _ = C.crawl_4122(s2, rep2, verbose=False)
        _, r23, _ = C.crawl_4123(s2, rep2, verbose=False)
        s25b, d25b, t25b, _ = C.crawl_4125(s2, rep2, verbose=False)
        d22 = G.diff_4122(r22b, t22b, s22b, data)
        d = _diff_4123(r23, data['rows23'])
        ch = _diff_4125(d25b, t25b, data['rows25'], data['total25'])
        ok = d22[2] and d == (0, 0, 0) and ch[0] == 0
        log('   4122 %s；4123 差异 %s；4125 变化 %d 行；结论：%s'
            % ('一致' if d22[2] else '新增%d/删除%d' % (d22[0], d22[1]),
               d, ch[0], '一致' if ok else '仍有差异'))
        result['recheck'] = {'4122': [d22[0], d22[1]], '4123': d, '4125': ch[0], 'ok': ok}
    elif dry:
        log('\n[S6] 演练模式，跳过复核')

    # ---------- 清单 ----------
    mp = os.path.join(out_dir, '工作流清单_%s_%s.json' % (bgbh, ts))
    with open(mp, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    log('\n[完成] 清单：%s' % mp)
    return result

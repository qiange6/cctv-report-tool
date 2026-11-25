# -*- coding: utf-8 -*-
"""选择性上传：按勾选的表，把输出目录里最新一套数据写回平台。"""
from __future__ import annotations

import os
import tempfile
import time

from . import crawl as C
from . import generate as G
from . import writeback as WB


def _read_rows(path, ncols):
    grid = C._read_grid(path)
    return [[str(x) for x in r[:ncols]] for r in grid[1:] if r and len(r) > 1 and str(r[1]).strip()]


def _cmp_4123(cur_rows, new_rows):
    cur = {(r[1], r[4], r[5]) for r in cur_rows if r and r[1]}
    new = {(r[1], r[4], r[5]) for r in new_rows if r and r[1]}
    return len(new - cur), len(cur - new)


def upload_selected(bgbh, out_dir, tables, dry=False, alias=None, log=print):
    """tables: {'4122','4123','4125','4126','4127'} 的子集。返回结果 dict。"""
    tables = set(tables)
    alias = alias or {}
    if not tables:
        raise RuntimeError('没有勾选任何表')

    cfg, meta = C.load_config()
    C.print_config(cfg, meta, bgbh)
    sess = C.Session(cfg, verbose=False).login()
    rep = C.locate_report(sess, bgbh, cfg['IGID'], days=3650, verbose=False)
    sess.bgbh, sess.rid, sess.pid, sess.igid = bgbh, rep['report_id'], str(rep['parentid']), cfg['IGID']
    log('[定位] %s report_id=%s parentid=%s' % (bgbh, sess.rid, sess.pid))

    rows26, recs, extra, files = C.load_outputs(bgbh, out_dir)
    miss = [t for t in ('4126', '4127') if not files.get(t)]
    if miss:
        raise RuntimeError('输出目录里没有最新的 %s，请先“只抓取”' % '、'.join(miss))
    log('[文件] 4126=%s' % os.path.basename(files['4126']))
    log('[文件] 4127=%s' % os.path.basename(files['4127']))
    rows27 = _read_rows(files['4127'], 37)

    result = {'uploaded': [], 'skipped': [], 'failed': []}

    # ---- 4126 / 4127 ----
    if tables & {'4126', '4127'}:
        plat_rows26, plat_stats = C.crawl_4126(sess, rep, verbose=False)
        plat26 = {r[1]: r for r in plat_rows26 if r and len(r) > 1 and r[1]}
        recs27, _ = C.crawl_4127(sess, rep, verbose=False)
        tmp = os.path.join(tempfile.gettempdir(), '_cctv_p27.xlsx')
        C.save_4127(recs27, tmp)
        p27 = [[('' if c is None else str(c).strip()) for c in r]
               for r in C._read_grid(tmp)]
        try:
            os.remove(tmp)
        except Exception:
            pass
        plat27 = {'%s|%s' % (r[1], r[2]): r for r in p27[1:] if r and r[1]}
        plan26, plan27 = WB.diff_rows(rows26, rows27, plat26, plat27, alias)
        log('[计划] 4126 %d 段 / 4127 %d 段' % (len(plan26), len(plan27)))

        if '4126' in tables:
            if not plan26:
                result['skipped'].append('4126')
                log('[4126] 平台已是目标值，跳过')
            elif dry:
                log('[4126] 演练：将写 %d 段（不提交）' % len(plan26))
            else:
                pages = max(1, int(plat_stats.get('pages') or 1))
                pos = WB.scan_4126_positions(sess, pages=pages)
                miss = [k for k in plan26 if k not in pos]
                if miss:
                    raise RuntimeError('4126 编辑页找不到：%s' % miss[:5])
                WB.upload(sess, '4126', plan26, pos)
                result['uploaded'].append('4126')
                log('[4126] 已上传 %d 段' % len(plan26))

        if '4127' in tables:
            if not plan27:
                result['skipped'].append('4127')
                log('[4127] 平台已是目标值，跳过')
            elif dry:
                log('[4127] 演练：将写 %d 段（不提交）' % len(plan27))
            else:
                order = {'%s|%s' % (r[1], r[2]): i for i, r in enumerate(p27[1:]) if r and r[1]}
                pos = {k: (order[k], 1) for k in plan27 if k in order}
                miss = [k for k in plan27 if k not in order]
                if miss:
                    raise RuntimeError('4127 编辑页找不到：%s' % miss[:5])
                WB.upload(sess, '4127', plan27, pos)
                result['uploaded'].append('4127')
                log('[4127] 已上传 %d 段' % len(plan27))

    # ---- 4122 / 4123 / 4125（由文件里的 4126+4127 生成后比较） ----
    if tables & {'4122', '4123', '4125'}:
        data = G.build_from_rows(rows26, rows27)
        if '4122' in tables:
            cur_s22, cur_rows22, cur_tot22, _ = C.crawl_4122(sess, rep, verbose=False)
            add, dele, same = G.diff_4122(cur_rows22, cur_tot22, cur_s22, data)
            if same:
                result['skipped'].append('4122')
                log('[4122] 平台与生成结果一致，跳过')
            elif dry:
                log('[4122] 演练：新增 %d 行 / 删除 %d 行（不提交）' % (add, dele))
            else:
                WB.upload_4122(sess, data['s22'], data['rows22'], data['total22'])
                result['uploaded'].append('4122')
                log('[4122] 已上传 %d 行（新增 %d / 删除 %d）'
                    % (len(data['rows22']), add, dele))
        if '4123' in tables:
            _, cur23, _ = C.crawl_4123(sess, rep, verbose=False)
            add, dele = _cmp_4123(cur23, data['rows23'])
            if add == 0 and dele == 0 and (extra.get('4123') or {}).get('summary') == data['s23']:
                result['skipped'].append('4123')
                log('[4123] 平台与生成结果一致，跳过')
            elif dry:
                log('[4123] 演练：新增 %d / 减少 %d（不提交）' % (add, dele))
            else:
                WB.upload_4123(sess, data['s23'], data['rows23'])
                result['uploaded'].append('4123')
                log('[4123] 已上传 %d 行（新增 %d / 减少 %d）' % (len(data['rows23']), add, dele))
        if '4125' in tables:
            cur_s, cur_defs, cur_tot, _ = C.crawl_4125(sess, rep, verbose=False)
            try:
                same_total = int((cur_tot or {}).get('小计')) == sum(data['total25'])
            except (TypeError, ValueError):
                same_total = False
            same = same_total and cur_s == data['s25']
            if same:
                result['skipped'].append('4125')
                log('[4125] 平台与生成结果一致，跳过')
            elif dry:
                log('[4125] 演练：将更新统计（不提交）')
            else:
                WB.upload_4125(sess, data['s25'], data['rows25'], data['total25'])
                result['uploaded'].append('4125')
                log('[4125] 已上传')

    log('[结果] 已上传：%s；已跳过：%s'
        % ('、'.join(result['uploaded']) or '无', '、'.join(result['skipped']) or '无'))
    return result

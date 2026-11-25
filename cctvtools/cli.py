# -*- coding: utf-8 -*-
"""统一入口：爬取 / 校验 / 生成 / 对比 / 写回。

用法：
  python cctv.py locate   --bgbh 26KHD00029
  python cctv.py crawl    --bgbh 26KHD00029 [--tables all|4126,4127,...] [--out-dir .]
  python cctv.py verify   --bgbh 26KHD00029 [--out-dir .]
  python cctv.py generate --bgbh 26KHD00029 --xlsx 对比标注.xlsx [--out-dir .]
  python cctv.py diff     --bgbh 26KHD00029 --xlsx 对比标注.xlsx [--out 变更清单.xlsx]
  python cctv.py upload   --bgbh 26KHD00029 --xlsx 对比标注.xlsx [--dry] [--only 4126]
  python cctv.py selftest
"""
from __future__ import annotations

import argparse
import datetime
import os
import sys

from . import crawl as C
from . import diff as D
from . import generate as G
from . import writeback as WB
from . import batch as BT
from . import credentials as CR
from . import uploader as UP
from . import verify_report as VR
from . import workflow as WF
from .common import parse_form, put


def make_session(bgbh):
    cfg, meta = C.load_config()
    C.print_config(cfg, meta, bgbh)
    s = C.Session(cfg, verbose=False).login()
    rep = C.locate_report(s, bgbh, cfg['IGID'], days=3650, verbose=False)
    s.bgbh, s.rid, s.pid, s.igid = bgbh, rep['report_id'], str(rep['parentid']), cfg['IGID']
    print('[定位] %s report_id=%s parentid=%s' % (bgbh, s.rid, s.pid))
    return s, rep


def cmd_workflow(args):
    WF.run(args.bgbh, out_dir=args.out_dir, upload=not args.no_upload, dry=args.dry,
           yes=args.yes, from_xlsx=args.from_xlsx,
           alias={a.split('=', 1)[0]: a.split('=', 1)[1] for a in (args.alias or [])})
    return 0


def _interactive_login():
    cfg, _ = C.load_config()
    if CR.has_login(cfg):
        return True
    print('尚未配置平台账号（配置文件：%s）' % CR.config_path())
    try:
        u = input('账号: ').strip()
        p = input('密码: ').strip()
    except EOFError:
        return False
    if not u or not p:
        print('账号或密码为空')
        return False
    new = {'BASE': cfg.get('BASE'), 'USER': u, 'PWD': p, 'IGID': cfg.get('IGID', '260')}
    path = CR.save_file(new, remember=True)
    C.OVERRIDES.update(new)
    print('已保存：%s' % path)
    return True


def interactive():
    print('=' * 52)
    print(' CCTV 报告工作流')
    print('=' * 52)
    if not _interactive_login():
        return 1
    try:
        bgbh = input('报告编号（如 26KHD00029）: ').strip()
    except EOFError:
        return 1
    if not bgbh:
        print('未输入报告编号')
        return 1
    print('''
  1) 全流程（抓取五张表 → 校验 → 生成 4123/4125 → 对比 → 上传 → 复核）
  2) 只抓取五张表
  3) 只做交叉校验（目录里最新一套）
  4) 用指定 Excel 生成 4123/4125 并上传
  5) 批量模式（读报告编号列表，逐个跑）
  0) 退出''')
    try:
        ch = input('选择 [1]: ').strip() or '1'
    except EOFError:
        return 1
    if ch == '0':
        return 0
    if ch == '2':
        return cmd_crawl(argparse.Namespace(bgbh=bgbh, table=None, tables='all',
                                            out_dir='.', max_pages_4126=80,
                                            max_pages_4127=800))
    if ch == '3':
        return cmd_verify(argparse.Namespace(bgbh=bgbh, out_dir='.'))
    if ch == '5':
        path = input('报告编号列表文件（一行一个）: ').strip().strip('"')
        up = input('要上传到平台吗？[y/N]: ').strip().lower() in ('y', 'yes', '是')
        each = False
        if up:
            each = input('逐报告确认后再上传吗？[Y/n]: ').strip().lower() not in ('n', 'no', '否')
        return cmd_batch(argparse.Namespace(bgbh=None, bgbh_list=[], file=path,
                                            out_dir='.', upload=up, dry=not up,
                                            yes=up, alias=[], confirm_each=each))
    if ch == '4':
        path = input('Excel 路径: ').strip().strip('"')
        yes = input('上传到平台？[y/N]: ').strip().lower() in ('y', 'yes', '是')
        return cmd_workflow(argparse.Namespace(bgbh=bgbh, out_dir='.', no_upload=False,
                                              dry=not yes, yes=yes, from_xlsx=path, alias=[]))
    return cmd_workflow(argparse.Namespace(bgbh=bgbh, out_dir='.', no_upload=False,
                                          dry=False, yes=False, from_xlsx=None, alias=[]))


def cmd_login(args):
    cfg, _ = C.load_config()
    if args.test:
        C.Session(cfg, verbose=True).login()
        print('登录验证通过')
    p = CR.save_file(cfg, remember=not args.no_remember)
    print('账号已保存：%s' % p)
    if args.no_remember:
        print('（未保存密码，下次启动需重新输入）')
    return 0


def cmd_verify_report(args):
    r = VR.build(args.bgbh, args.out_dir)
    print('总体结论：%s' % {'PASS': '通过', 'WARN': '通过（有提醒）', 'FAIL': '不通过'}[r['verdict']])
    print('报告文件：%s' % r['path'])
    for c in r['checks']:
        print('  [%s] %s' % (c['level'], c['msg']))
    return 0 if r['verdict'] != 'FAIL' else 2


def cmd_upload_tables(args):
    tables = [t.strip() for t in args.tables.split(',') if t.strip()]
    r = UP.upload_selected(args.bgbh, args.out_dir, tables, dry=args.dry)
    print('已上传：%s' % ('、'.join(r['uploaded']) or '无'))
    print('已跳过：%s' % ('、'.join(r['skipped']) or '无'))
    return 0


def cmd_batch(args):
    lst = list(getattr(args, 'bgbh_list', []) or [])
    if getattr(args, 'file', None):
        lst += BT.read_list(args.file)
    lst = [x for x in lst if x]
    if not lst:
        print('没有报告编号：用 `--file 列表.txt` 或直接跟编号')
        return 1
    aliases = {}
    for a in (getattr(args, 'alias', None) or []):
        if ':' in a:
            b, rest = a.split(':', 1)
            if '=' in rest:
                k, v = rest.split('=', 1)
                aliases.setdefault(b, {})[k] = v
    BT.run_batch(lst, out_dir=args.out_dir, upload=args.upload, dry=args.dry,
                 yes=args.yes, aliases=aliases, confirm_each=getattr(args, 'confirm_each', False))
    return 0


def cmd_locate(args):
    s, rep = make_session(args.bgbh)
    print(rep)
    return 0


def cmd_crawl(args):
    s, rep = make_session(args.bgbh)
    if getattr(args, 'table', None):
        tables = [args.table]
    else:
        tables = (['4122', '4123', '4125', '4126', '4127'] if args.tables == 'all'
                  else args.tables.split(','))
    out = args.out_dir
    os.makedirs(out, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    saved = []
    for t in tables:
        if t == '4126':
            rows, _ = C.crawl_4126(s, rep, max_pages=args.max_pages_4126)
            saved.append(C.save_4126(rows, os.path.join(out, C.output_name(t, args.bgbh, ts))))
        elif t == '4127':
            recs, _ = C.crawl_4127(s, rep, max_pages=args.max_pages_4127)
            saved.append(C.save_4127(recs, os.path.join(out, C.output_name(t, args.bgbh, ts))))
        elif t == '4123':
            sm, rows, _ = C.crawl_4123(s, rep)
            saved.append(C.save_4123(sm, rows, os.path.join(out, C.output_name(t, args.bgbh, ts))))
        elif t == '4125':
            sm, defs, tot, _ = C.crawl_4125(s, rep)
            saved.append(C.save_4125(sm, defs, tot, os.path.join(out, C.output_name(t, args.bgbh, ts))))
        elif t == '4122':
            sm, rows, tot, _ = C.crawl_4122(s, rep)
            saved.append(C.save_4122(sm, rows, tot, os.path.join(out, C.output_name(t, args.bgbh, ts))))
    for p in saved:
        print('已保存', p)
    return 0


def cmd_all(args):
    from types import SimpleNamespace
    cfg, meta = C.load_config()
    C.print_config(cfg, meta, args.bgbh)
    ns = SimpleNamespace(bgbh=args.bgbh, out_dir=args.out_dir, days=3650,
                         max_pages_4126=args.max_pages_4126,
                         max_pages_4127=args.max_pages_4127, sleep=0.0,
                         allow_empty=False, browser_fallback=False, headless=True,
                         rid=None, parentid=None)
    return C.run_all(cfg, ns)


def cmd_verify(args):
    return C.verify_latest(args.bgbh, args.out_dir)


def cmd_generate(args):
    rows26, rows27 = D.read_workbook(args.xlsx)
    data = G.build_from_rows(rows26, rows27)
    print('4122:', data['s22'])
    print('4123:', data['s23'])
    print('4125:', data['s25'])
    p1, p2, p3 = G.write_outputs(data, args.bgbh, args.out_dir)
    print('已生成:\n ', p1, '\n ', p2, '\n ', p3)
    return 0


def cmd_diff(args):
    rows26, rows27 = D.read_workbook(args.xlsx)
    s, rep = make_session(args.bgbh)
    p26, _ = C.crawl_4126(s, rep, verbose=False)
    p27r, _ = C.crawl_4127(s, rep, verbose=False)
    plat26 = {r[1]: r for r in p26}
    plat27 = {'%s|%s' % (r['起始井号'], r['终止井号']): None for r in p27r}
    # 平台 4127 转 37 列行
    tmp = os.path.join(args.out_dir, '_platform_4127.xlsx')
    C.save_4127(p27r, tmp)
    import openpyxl
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    p27 = [[('' if c is None else str(c).strip()) for c in r] for r in ws.iter_rows(values_only=True)]
    wb.close()
    plat27 = {'%s|%s' % (r[1], r[2]): r for r in p27[1:] if r[1]}
    alias = {}
    for a in (args.alias or []):
        k, v = a.split('=', 1)
        alias[k] = v
    changes = D.compare(plat26, plat27, rows26, rows27, alias)
    print('变更 %d 条（4126 %d / 4127 %d）' % (
        len(changes),
        sum(1 for c in changes if c['table'] == '4126'),
        sum(1 for c in changes if c['table'] == '4127')))
    out = args.out or os.path.join(args.out_dir, '变更清单-%s.xlsx' % args.bgbh)
    D.write_list(changes, out, args.bgbh)
    print('已生成', out)
    return 0


def cmd_upload(args):
    rows26, rows27 = D.read_workbook(args.xlsx)
    s, rep = make_session(args.bgbh)
    alias = {}
    for a in (args.alias or []):
        k, v = a.split('=', 1)
        alias[k] = v

    # 平台现状
    p26, _ = C.crawl_4126(s, rep, verbose=False)
    p27r, _ = C.crawl_4127(s, rep, verbose=False)
    tmp = os.path.join(args.out_dir, '_platform_4127.xlsx')
    C.save_4127(p27r, tmp)
    import openpyxl
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    p27 = [[('' if c is None else str(c).strip()) for c in r] for r in ws.iter_rows(values_only=True)]
    wb.close()
    plat26 = {r[1]: r for r in p26}
    plat27 = {'%s|%s' % (r[1], r[2]): r for r in p27[1:] if r[1]}

    plan26, plan27 = WB.diff_rows(rows26, rows27, plat26, plat27, alias)
    print('[计划] 4126 %d 段 / %d 字段；4127 %d 段 / %d 字段' % (
        len(plan26), sum(len(v) for v in plan26.values()),
        len(plan27), sum(len(v) for v in plan27.values())))

    if args.dry:
        print('[演练] 只列计划，不写回')
        for k, v in plan26.items():
            print('  4126', k, v)
        for k, v in plan27.items():
            print('  4127', k, v)
        return 0

    order27 = {'%s|%s' % (r[1], r[2]): i for i, r in enumerate(p27[1:]) if r[1]}
    if plan26:
        pos26 = WB.scan_4126_positions(s)
        missing = [k for k in plan26 if k not in pos26]
        if missing:
            raise RuntimeError('4126 编辑页找不到: %s' % missing)
        WB.upload(s, '4126', plan26, pos26)
    if plan27:
        pos27 = {k: (order27[k], 1) for k in plan27 if k in order27}
        WB.upload(s, '4127', plan27, pos27)
    print('[完成] 上传结束，请用 `cctv.py verify --bgbh %s` 复核' % args.bgbh)
    return 0


def build_parser():
    ap = argparse.ArgumentParser(prog='cctv.py', description='CCTV 报告工具（统一入口）')
    ap.add_argument('--bgbh', default=None, help='报告编号，如 26KHD00029')
    ap.add_argument('--show-config', action='store_true', help='只打印配置来源')
    ap.add_argument('--user', default=None, help='平台账号（也可用配置文件/环境变量 CCTV_USER）')
    ap.add_argument('--pwd', default=None, help='平台密码（建议用配置文件或 CCTV_PWD，命令行会留在历史里）')
    ap.add_argument('--base', default=None, help='平台地址，如 http://36.134.56.138:7500')
    sub = ap.add_subparsers(dest='cmd')

    p = sub.add_parser('locate', help='只定位报告')
    p.set_defaults(func=cmd_locate)

    p = sub.add_parser('crawl', help='抓取表格（默认全部）')
    p.add_argument('table', nargs='?', default=None, help='只抓一张，如 4126')
    p.add_argument('--tables', default='all')
    p.add_argument('--out-dir', default='.')
    p.add_argument('--max-pages-4126', type=int, default=80)
    p.add_argument('--max-pages-4127', type=int, default=800)
    p.set_defaults(func=cmd_crawl)

    p = sub.add_parser('workflow', aliases=['run'], help='一键全流程（推荐）')
    p.add_argument('--out-dir', default='.')
    p.add_argument('--from-xlsx', default=None, help='用修改后的 Excel 作为 4126/4127 数据源')
    p.add_argument('--dry', action='store_true', help='演练，不写入')
    p.add_argument('--yes', action='store_true', help='上传时不询问')
    p.add_argument('--no-upload', action='store_true')
    p.add_argument('--alias', action='append', help='编号纠正，如 Y16-Y17=Y15-Y17')
    p.set_defaults(func=cmd_workflow)

    p = sub.add_parser('login', help='保存/修改平台账号（密码 DPAPI 加密）')
    p.add_argument('--no-remember', action='store_true', help='只记账号不记密码')
    p.add_argument('--test', action='store_true', help='保存前先验证登录')
    p.set_defaults(func=cmd_login)

    p = sub.add_parser('verify-report', help='校验并生成 xlsx 报告')
    p.add_argument('--out-dir', default='.')
    p.set_defaults(func=cmd_verify_report)

    p = sub.add_parser('upload-tables', help='按表上传（选择性）')
    p.add_argument('--tables', default='4122,4126,4127,4123,4125',
                   help='逗号分隔，可选 4122,4126,4127,4123,4125')
    p.add_argument('--out-dir', default='.')
    p.add_argument('--dry', action='store_true')
    p.set_defaults(func=cmd_upload_tables)

    p = sub.add_parser('batch', help='批量模式：读报告编号列表逐个跑')
    p.add_argument('bgbh_list', nargs='*', help='一个或多个报告编号')
    p.add_argument('--file', default=None, help='报告编号列表文件（一行一个）')
    p.add_argument('--out-dir', default='.')
    p.add_argument('--upload', action='store_true', help='批量上传（需配合 --yes）')
    p.add_argument('--dry', action='store_true')
    p.add_argument('--yes', action='store_true')
    p.add_argument('--alias', action='append', help='按报告纠正编号，如 26KHD00029:Y16-Y17=Y15-Y17')
    p.add_argument('--confirm-each', action='store_true',
                   help='逐报告确认后再上传（y=是 / n=跳过 / a=全部 / q=退出）')
    p.set_defaults(func=cmd_batch)

    p = sub.add_parser('all', help='抓取五张表 + 交叉校验 + manifest（兼容旧入口）')
    p.add_argument('--out-dir', default='.')
    p.add_argument('--max-pages-4126', type=int, default=80)
    p.add_argument('--max-pages-4127', type=int, default=800)
    p.set_defaults(func=cmd_all)

    p = sub.add_parser('verify', help='离线交叉校验（目录里最新一套）')
    p.add_argument('--out-dir', default='.')
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser('generate', help='由 4126/4127 生成 4123/4125')
    p.add_argument('--xlsx', required=True, help='含 4126+4127 两个 sheet 的 Excel')
    p.add_argument('--out-dir', default='.')
    p.set_defaults(func=cmd_generate)

    p = sub.add_parser('diff', help='平台 vs Excel 生成变更清单')
    p.add_argument('--xlsx', required=True)
    p.add_argument('--out', default=None)
    p.add_argument('--out-dir', default='.')
    p.add_argument('--alias', action='append', help='编号纠正，如 Y16-Y17=Y15-Y17')
    p.set_defaults(func=cmd_diff)

    p = sub.add_parser('upload', help='按 Excel 写回平台')
    p.add_argument('--xlsx', required=True)
    p.add_argument('--out-dir', default='.')
    p.add_argument('--alias', action='append')
    p.add_argument('--dry', action='store_true')
    p.add_argument('--only', default=None, help='仅 4126 或 4127（预留）')
    p.set_defaults(func=cmd_upload)

    p = sub.add_parser('selftest', help='离线自检')
    p.set_defaults(func=lambda a: C.selftest())
    return ap


def _normalize(argv):
    """允许全局参数（--bgbh/--user/--pwd/--base）写在子命令之后。"""
    argv = list(argv)
    for flag in ('--bgbh', '--user', '--pwd', '--base'):
        while flag in argv[1:]:
            i = argv.index(flag)
            if i + 1 >= len(argv):
                break
            val = argv[i + 1]
            del argv[i:i + 2]
            argv[0:0] = [flag, val]
    return argv


def main(argv=None):
    ap = build_parser()
    argv = _normalize(sys.argv[1:] if argv is None else argv)
    args = ap.parse_args(argv)
    ov = {}
    if getattr(args, 'user', None):
        ov['USER'] = args.user
    if getattr(args, 'pwd', None):
        ov['PWD'] = args.pwd
    if getattr(args, 'base', None):
        ov['BASE'] = args.base
    if ov:
        C.OVERRIDES.update(ov)
    if getattr(args, 'show_config', False):
        cfg, meta = C.load_config()
        C.print_config(cfg, meta, args.bgbh)
        return 0
    if not args.cmd:
        if sys.stdin.isatty():
            return interactive()
        ap.print_help()
        return 1
    if args.cmd not in ('selftest', 'batch', 'login') and not getattr(args, 'bgbh', None):
        ap.error('需要 --bgbh')
    try:
        return args.func(args)
    except RuntimeError as e:
        print('\n[错误] %s' % e)
        return 2


if __name__ == '__main__':
    sys.exit(main())

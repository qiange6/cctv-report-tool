# -*- coding: utf-8 -*-
"""CCTV 报告工具 - 图形界面（Tkinter）。

双击运行：输入报告编号 → 选输出目录 → 点按钮。
"""
from __future__ import annotations

import contextlib
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk

from cctvtools import common as C
from cctvtools import credentials as CR
from cctvtools import uploader as UP
from cctvtools import verify_report as VR

# 打包成 exe 后，tkinter 需要能找到随包的 tcl/tk 数据目录
if getattr(sys, 'frozen', False):
    _base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    os.environ.setdefault('TCL_LIBRARY', os.path.join(_base, 'tcl8.6'))
    os.environ.setdefault('TK_LIBRARY', os.path.join(_base, 'tk8.6'))


def base_dir():
    """打包成 exe 后取 exe 所在目录，源码运行取当前文件目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def default_out_dir():
    d = os.path.join(base_dir(), '输出')
    os.makedirs(d, exist_ok=True)
    return d


class QueueWriter:
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(('log', s))

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CCTV 报告工具')
        self.geometry('980x640')
        self.minsize(820, 520)
        self.q = queue.Queue()
        self.running = False
        self._report_path = None
        self._build()
        self.after(120, self._poll)
        if not self._ensure_login():
            self.after(100, self.destroy)
            return
        self._update_title()

    # ---------- 界面 ----------
    def _build(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill='x')

        ttk.Label(frm, text='报告编号').grid(row=0, column=0, sticky='w')
        self.e_bgbh = ttk.Entry(frm, width=22)
        self.e_bgbh.grid(row=0, column=1, padx=6, sticky='w')
        self.e_bgbh.insert(0, '26KHD00029')
        ttk.Label(frm, text='（如 26KHD00029）', foreground='#888').grid(row=0, column=2, sticky='w')

        ttk.Label(frm, text='输出目录').grid(row=1, column=0, sticky='w', pady=(6, 0))
        self.e_out = ttk.Entry(frm, width=64)
        self.e_out.grid(row=1, column=1, columnspan=2, padx=6, sticky='w', pady=(6, 0))
        self.e_out.insert(0, default_out_dir())
        ttk.Button(frm, text='浏览…', width=8, command=self.pick_dir).grid(row=1, column=3, padx=4)

        btns = ttk.Frame(self, padding=(10, 0, 10, 8))
        btns.pack(fill='x')
        self.b_run = ttk.Button(btns, text='一键全流程', width=14, command=self.do_workflow)
        self.b_up = ttk.Button(btns, text='全流程+上传', width=14, command=self.do_workflow_upload)
        self.b_crawl = ttk.Button(btns, text='只抓取', width=10, command=self.do_crawl)
        self.b_gen = ttk.Button(btns, text='只生成 4122/4123/4125', width=21, command=self.do_generate)
        self.b_ver = ttk.Button(btns, text='校验报告', width=10, command=self.do_verify)
        self.b_upload = ttk.Button(btns, text='上传勾选', width=10, command=self.do_upload_tables)
        self.b_batch = ttk.Button(btns, text='批量…', width=8, command=self.do_batch)
        self.b_cred = ttk.Button(btns, text='账号', width=6, command=self.change_account)
        for b in (self.b_run, self.b_up, self.b_crawl, self.b_gen, self.b_ver,
                  self.b_upload, self.b_batch, self.b_cred):
            b.pack(side='left', padx=4)

        opt = ttk.LabelFrame(self, text='上传范围（勾选后点“上传勾选”）', padding=(10, 4))
        opt.pack(fill='x', padx=10, pady=(0, 6))
        self.v22 = tk.BooleanVar(value=True)
        self.v26 = tk.BooleanVar(value=True)
        self.v27 = tk.BooleanVar(value=True)
        self.v23 = tk.BooleanVar(value=True)
        self.v25 = tk.BooleanVar(value=True)
        for text, var in (('4122 工程量表', self.v22),
                          ('4126 评估表', self.v26), ('4127 成果表', self.v27),
                          ('4123 汇总表', self.v23), ('4125 统计表', self.v25)):
            ttk.Checkbutton(opt, text=text, variable=var).pack(side='left', padx=(0, 14))
        ttk.Label(opt, text='（4122 由 4126+4127 汇总生成后写回）', foreground='#888').pack(side='left')

        self.lbl = ttk.Label(self, text='就绪', padding=(10, 0))
        self.lbl.pack(fill='x')

        body = ttk.Frame(self, padding=(10, 4, 10, 10))
        body.pack(fill='both', expand=True)
        self.txt = tk.Text(body, wrap='none', font=('Consolas', 9), bg='#1e1e1e', fg='#dcdcdc')
        yscroll = ttk.Scrollbar(body, orient='vertical', command=self.txt.yview)
        xscroll = ttk.Scrollbar(body, orient='horizontal', command=self.txt.xview)
        self.txt.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.txt.grid(row=0, column=0, sticky='nsew')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll.grid(row=1, column=0, sticky='ew')
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

    # ---------- 工具 ----------
    def pick_dir(self):
        d = filedialog.askdirectory(initialdir=self.e_out.get() or base_dir())
        if d:
            self.e_out.delete(0, 'end')
            self.e_out.insert(0, d)

    def bgbh(self):
        v = self.e_bgbh.get().strip()
        if not v:
            messagebox.showwarning('提示', '请先输入报告编号')
            return None
        if not re.fullmatch(r'\d{2}KHD\d{4,6}', v, re.I):
            if not messagebox.askyesno('确认', '“%s”不像标准报告编号（如 26KHD00029），仍要继续？' % v):
                return None
        return v

    def out_dir(self):
        d = self.e_out.get().strip() or default_out_dir()
        os.makedirs(d, exist_ok=True)
        return d

    def log(self, s):
        self.txt.insert('end', s)
        self.txt.see('end')

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == 'log':
                    self.log(payload)
                elif kind == 'done':
                    self._finish(payload)
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _finish(self, ok):
        self.running = False
        for b in (self.b_run, self.b_up, self.b_crawl, self.b_gen, self.b_ver,
                  self.b_upload, self.b_batch):
            b.state(['!disabled'])
        self.lbl.config(text='完成' if ok else '失败（见上方日志）')

    def start(self, fn, title):
        if self.running:
            messagebox.showinfo('提示', '任务正在运行，请稍候')
            return
        self.running = True
        self.txt.delete('1.0', 'end')
        self.lbl.config(text='正在：%s …' % title)
        for b in (self.b_run, self.b_up, self.b_crawl, self.b_gen, self.b_ver,
                  self.b_upload, self.b_batch):
            b.state(['disabled'])
        threading.Thread(target=self._worker, args=(fn,), daemon=True).start()

    def _worker(self, fn):
        ok = True
        t0 = time.time()
        with contextlib.redirect_stdout(QueueWriter(self.q)), \
                contextlib.redirect_stderr(QueueWriter(self.q)):
            try:
                fn()
            except Exception:
                ok = False
                traceback.print_exc()
            finally:
                print('\n用时 %.1f 秒' % (time.time() - t0))
        self.q.put(('done', ok))

    # ---------- 账号 ----------
    def _update_title(self):
        cfg, _ = C.load_config()
        self.title('CCTV 报告工具 - %s' % (cfg.get('USER') or '未登录'))
        if self.lbl.cget('text') in ('就绪', ''):
            self.lbl.config(text='就绪 · 账号 %s' % (cfg.get('USER') or '未登录'))

    def _ensure_login(self):
        cfg, _ = C.load_config()
        if CR.has_login(cfg):
            return True
        return self.login_dialog()

    def login_dialog(self):
        cfg, _ = C.load_config()
        dlg = tk.Toplevel(self)
        dlg.title('平台账号')
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text='平台地址').grid(row=0, column=0, sticky='w')
        e_b = ttk.Entry(frm, width=38)
        e_b.insert(0, cfg.get('BASE', C.DEFAULTS['BASE']))
        e_b.grid(row=0, column=1, columnspan=2, padx=6, pady=3)
        ttk.Label(frm, text='账号').grid(row=1, column=0, sticky='w')
        e_u = ttk.Entry(frm, width=24)
        e_u.insert(0, cfg.get('USER', ''))
        e_u.grid(row=1, column=1, columnspan=2, padx=6, pady=3, sticky='w')
        ttk.Label(frm, text='密码').grid(row=2, column=0, sticky='w')
        e_p = ttk.Entry(frm, width=24, show='*')
        e_p.insert(0, cfg.get('PWD', ''))
        e_p.grid(row=2, column=1, columnspan=2, padx=6, pady=3, sticky='w')
        remember = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text='记住密码（Windows 加密保存）', variable=remember).grid(
            row=3, column=1, columnspan=2, sticky='w', padx=6, pady=(2, 6))
        ttk.Label(frm, text='账号保存在：%s' % CR.config_path(),
                  foreground='#888').grid(row=4, column=0, columnspan=3, sticky='w', pady=(0, 8))
        ok = {'v': False}

        def do_ok(*_):
            u, p, b = e_u.get().strip(), e_p.get(), e_b.get().strip()
            if not u or not p:
                messagebox.showwarning('提示', '账号和密码都要填', parent=dlg)
                return
            new = {'BASE': b or C.DEFAULTS['BASE'], 'USER': u, 'PWD': p,
                   'IGID': cfg.get('IGID', '260')}
            path = CR.save_file(new, remember=remember.get())
            C.OVERRIDES.update(new)
            ok['v'] = True
            print('账号已保存：%s' % path)
            dlg.destroy()

        bar = ttk.Frame(frm)
        bar.grid(row=5, column=0, columnspan=3, pady=6)
        ttk.Button(bar, text='确定', width=10, command=do_ok).pack(side='left', padx=6)
        ttk.Button(bar, text='退出', width=10, command=dlg.destroy).pack(side='left', padx=6)
        dlg.bind('<Return>', do_ok)
        dlg.wait_window()
        return ok['v']

    def change_account(self):
        if self.login_dialog():
            self._update_title()
            self.lbl.config(text='账号已更新')

    # ---------- 动作 ----------
    def _wf(self, upload=False):
        from cctvtools import workflow as WF
        b = self.bgbh()
        if not b:
            return
        d = self.out_dir()
        self.start(lambda: WF.run(b, out_dir=d, upload=upload, yes=True,
                                  log=lambda s='': print(s)), '工作流 ' + b)

    def do_workflow(self):
        self._wf(upload=False)

    def do_workflow_upload(self):
        if not self.bgbh():
            return
        if not messagebox.askyesno('确认上传', '将把 4122/4123/4125（以及需要时 4126/4127）写回平台，继续？'):
            return
        self._wf(upload=True)

    def do_crawl(self):
        from cctvtools import crawl as C
        b = self.bgbh()
        if not b:
            return
        d = self.out_dir()

        def task():
            cfg, _ = C.load_config()
            s = C.Session(cfg, verbose=False).login()
            rep = C.locate_report(s, b, cfg['IGID'], days=3650)
            import datetime
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
            rows, _ = C.crawl_4126(s, rep)
            C.save_4126(rows, os.path.join(d, C.output_name('4126', b, ts)))
            recs, _ = C.crawl_4127(s, rep)
            C.save_4127(recs, os.path.join(d, C.output_name('4127', b, ts)))
            sm, r, _ = C.crawl_4123(s, rep)
            C.save_4123(sm, r, os.path.join(d, C.output_name('4123', b, ts)))
            sm, d25, t25, _ = C.crawl_4125(s, rep)
            C.save_4125(sm, d25, t25, os.path.join(d, C.output_name('4125', b, ts)))
            sm, r, t, _ = C.crawl_4122(s, rep)
            C.save_4122(sm, r, t, os.path.join(d, C.output_name('4122', b, ts)))
            print('五张表已保存到', d)
        self.start(task, '抓取 ' + b)

    def do_generate(self):
        from cctvtools import diff as D
        from cctvtools import generate as G
        b = self.bgbh()
        if not b:
            return
        d = self.out_dir()
        files = [os.path.join(d, f) for f in os.listdir(d)]
        c26 = [f for f in files if os.path.basename(f).startswith('管段状况评估表-%s' % b)]
        c27 = [f for f in files if os.path.basename(f).startswith('管段检测与评估成果表-%s' % b)]
        if not c26 or not c27:
            messagebox.showwarning('提示', '输出目录里没有 %s 的 4126/4127 抓取文件，请先“只抓取”' % b)
            return
        f26, f27 = sorted(c26)[-1], sorted(c27)[-1]

        def task():
            print('数据源：\n  %s\n  %s' % (os.path.basename(f26), os.path.basename(f27)))
            import openpyxl
            def rows(p, n):
                wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
                ws = wb.worksheets[0]
                out = [[('' if c is None else str(c).strip()) for c in r]
                       for r in ws.iter_rows(min_row=2, values_only=True)]
                wb.close()
                return [r[:n] for r in out if r and r[1]]
            data = G.build_from_rows(rows(f26, 21), rows(f27, 37))
            p1, p2, p3 = G.write_outputs(data, b, d)
            print(data['s22'])
            print(data['s23'])
            print(data['s25'])
            print('已生成：\n ', p1, '\n ', p2, '\n ', p3)
        self.start(task, '生成 4122/4123/4125')

    def do_verify(self):
        b = self.bgbh()
        if not b:
            return
        d = self.out_dir()
        self._report_path = None

        def task():
            r = VR.build(b, d)
            self._report_path = r['path']
            print('总体结论：%s' % {'PASS': '通过', 'WARN': '通过（有提醒）',
                                  'FAIL': '不通过'}[r['verdict']])
            for c in r['checks']:
                print('  [%s] %s' % (c['level'], c['msg']))
            print('\n校验报告：%s' % r['path'])
        self.start(task, '校验并生成报告')

    def do_upload_tables(self):
        b = self.bgbh()
        if not b:
            return
        tables = [t for t, v in (('4122', self.v22), ('4126', self.v26), ('4127', self.v27),
                                 ('4123', self.v23), ('4125', self.v25)) if v.get()]
        if not tables:
            messagebox.showwarning('提示', '至少勾选一张表')
            return
        if not messagebox.askyesno('确认上传',
                                   '将上传：%s\n\n数据来源：输出目录里最新一套 xlsx\n'
                                   '平台已是目标值的表会自动跳过。继续？' % '、'.join(tables)):
            return
        d = self.out_dir()
        self.start(lambda: UP.upload_selected(b, d, tables,
                                              log=lambda s='': print(s)), '上传 ' + b)

    def do_batch(self):
        from cctvtools import batch as BT
        path = filedialog.askopenfilename(title='选择报告编号列表（一行一个）',
                                          filetypes=[('文本', '*.txt'), ('所有文件', '*.*')])
        if not path:
            return
        lst = BT.read_list(path)
        if not lst:
            messagebox.showwarning('提示', '列表为空')
            return
        up = messagebox.askyesno('批量', '共 %d 个报告。\n要上传到平台吗？\n\n选“否”= 只抓取/生成/对比' % len(lst))
        d = self.out_dir()
        self.start(lambda: BT.run_batch(lst, out_dir=d, upload=up, yes=True,
                                        log=lambda s='': print(s)),
                   '批量 %d 个报告' % len(lst))


def _write_error(text):
    try:
        if getattr(sys, 'frozen', False):
            d = os.path.dirname(sys.executable)
        else:
            d = os.path.dirname(os.path.abspath(__file__))
        p = os.path.join(d, 'error.log')
        import datetime
        with open(p, 'a', encoding='utf-8') as f:
            f.write('\n===== %s =====\n%s' % (datetime.datetime.now(), text))
        return p
    except Exception:
        return None


def main():
    if '--smoke' in sys.argv:
        C.OVERRIDES.update({'USER': 'smoke', 'PWD': 'smoke'})
        app = App()
        app.update_idletasks()
        app.update()
        app.destroy()
        print('GUI smoke OK')
        return
    try:
        app = App()
        app.report_callback_exception = lambda *a: _write_error(
            ''.join(traceback.format_exception(*a)))
        app.mainloop()
    except Exception:
        tb = traceback.format_exc()
        p = _write_error(tb)
        try:
            messagebox.showerror('启动失败', '%s\n\n日志：%s' % (tb[-800:], p))
        except Exception:
            pass


if __name__ == '__main__':
    main()

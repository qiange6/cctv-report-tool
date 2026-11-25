# -*- coding: utf-8 -*-
"""打包 exe（PyInstaller）。

用法：
    python _build/build_exe.py                  # 默认：目录版（onedir，推荐）
    python _build/build_exe.py --onefile        # 单文件版（onefile）

产物：
    dist/onedir/CCTV报告工具/CCTV报告工具.exe     图形界面（推荐）
    dist/onedir/CCTV命令行/CCTV命令行.exe         命令行版
    dist/CCTV报告工具.exe / CCTV命令行.exe        仅 --onefile 时
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VENDOR = os.path.join(BASE, '_vendor')
BUILD = os.path.join(BASE, '_build')
DIST = os.path.join(BASE, 'dist')
TMP = os.path.join(BUILD, 'tmp')
ONEFILE = '--onefile' in sys.argv
PY = sys.base_prefix  # Anaconda 根目录
BINARIES = [
    (os.path.join(PY, 'DLLs', '_tkinter.pyd'), '.'),
    (os.path.join(PY, 'DLLs', 'pyexpat.pyd'), '.'),
    (os.path.join(PY, 'DLLs', '_ctypes.pyd'), '.'),
    (os.path.join(PY, 'Library', 'bin', 'ffi-8.dll'), '.'),
    (os.path.join(PY, 'Library', 'bin', 'ffi-7.dll'), '.'),
    (os.path.join(PY, 'Library', 'bin', 'ffi.dll'), '.'),
    (os.path.join(PY, 'Library', 'bin', 'libexpat.dll'), '.'),
    (os.path.join(PY, 'Library', 'bin', 'tcl86t.dll'), '.'),
    (os.path.join(PY, 'Library', 'bin', 'tk86t.dll'), '.'),
]
DATAS = [
    (os.path.join(PY, 'Library', 'lib', 'tcl8.6'), 'tcl8.6'),
    (os.path.join(PY, 'Library', 'lib', 'tk8.6'), 'tk8.6'),
]

os.makedirs(TMP, exist_ok=True)
os.makedirs(DIST, exist_ok=True)
env = dict(os.environ)
env['PYTHONPATH'] = VENDOR + os.pathsep + BASE
env['TEMP'] = TMP
env['TMP'] = TMP


def build(name, entry, windowed):
    args = [sys.executable, '-m', 'PyInstaller',
            '--noconfirm', '--clean',
            '--onefile' if ONEFILE else '--onedir',
            '--windowed' if windowed else '--console',
            '--name', name,
            '--paths', BASE,
            '--collect-submodules', 'cctvtools',
            '--exclude-module', 'pkg_resources',
            '--exclude-module', 'setuptools',
            '--workpath', os.path.join(BUILD, 'work_' + name),
            '--specpath', BUILD]
    for src, dst in BINARIES:
        if os.path.exists(src):
            args += ['--add-binary', '%s%s%s' % (src, os.pathsep, dst)]
    for src, dst in DATAS:
        if os.path.isdir(src):
            args += ['--add-data', '%s%s%s' % (src, os.pathsep, dst)]
    if ONEFILE:
        args += ['--runtime-tmpdir', '.', '--distpath', DIST, entry]
    else:
        args += ['--distpath', os.path.join(DIST, 'onedir'), entry]
    print('\n>>>', ' '.join(args))
    r = subprocess.run(args, cwd=BASE, env=env, capture_output=True,
                       text=True, encoding='utf-8', errors='replace')
    tail = '\n'.join((r.stdout or '').splitlines()[-6:])
    print(tail)
    if r.returncode != 0:
        print((r.stderr or '')[-2000:])
        raise SystemExit('%s 打包失败（%d）' % (name, r.returncode))


if __name__ == '__main__':
    build('CCTV报告工具', 'cctv_gui.py', windowed=True)
    build('CCTV命令行', 'cctv.py', windowed=False)
    if ONEFILE:
        for f in ('CCTV报告工具.exe', 'CCTV命令行.exe'):
            p = os.path.join(DIST, f)
            print('%-12s %8.1f MB  %s' % (f, os.path.getsize(p) / 1048576, p))
    else:
        for sub in ('CCTV报告工具', 'CCTV命令行'):
            p = os.path.join(DIST, 'onedir', sub, sub + '.exe')
            print('%-12s %8.1f MB  %s' % (sub, os.path.getsize(p) / 1048576, p))

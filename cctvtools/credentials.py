# -*- coding: utf-8 -*-
"""账号配置：配置文件 + Windows DPAPI 加密保存密码。

优先级（高 → 低）：
    命令行参数 --user/--pwd/--base
    环境变量 CCTV_USER / CCTV_PWD / CCTV_BASE
    配置文件（%APPDATA%\\CCTV工具\\cctv_config.json，可用 CCTV_CONFIG 指定）
    内置 DEFAULTS（cctvtools/common.py）
"""
from __future__ import annotations

import base64
import ctypes
import ctypes.wintypes as wt
import datetime
import json
import os
import sys


# ==================== 配置文件位置 ====================
def _app_dir():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'CCTV工具')


def _candidates():
    """按优先级给出候选配置文件路径。"""
    out = []
    if os.environ.get('CCTV_CONFIG'):
        out.append(os.environ['CCTV_CONFIG'])
    if getattr(sys, 'frozen', False):
        out.append(os.path.join(_app_dir(), 'cctv_config.json'))
        out.append(os.path.join(os.path.dirname(sys.executable), 'cctv_config.json'))
    else:
        out.append(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                'cctv_config.json'))
    out.append(os.path.join(os.path.expanduser('~'), '.cctv', 'cctv_config.json'))
    return out


def _writable(path):
    try:
        d = os.path.dirname(path) or '.'
        os.makedirs(d, exist_ok=True)
        with open(path, 'a', encoding='utf-8'):
            pass
        return True
    except Exception:
        return False


def config_path(writable=False):
    """默认返回实际使用的配置文件；writable=True 时返回可写入的路径。"""
    cands = _candidates()
    if not writable:
        for p in cands:
            if os.path.exists(p):
                return p
        return next((p for p in cands if _writable(p)), cands[-1])
    for p in cands:
        if _writable(p):
            return p
    return cands[-1]


# ==================== DPAPI（Windows 用户级加密） ====================
class _BLOB(ctypes.Structure):
    _fields_ = [('cbData', wt.DWORD), ('pbData', ctypes.POINTER(ctypes.c_char))]


def _make_blob(data: bytes):
    buf = ctypes.create_string_buffer(data, len(data))
    return _BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf


def _dpapi(data: bytes, protect: bool):
    if os.name != 'nt':
        raise OSError('仅 Windows 支持 DPAPI')
    blob_in, _buf = _make_blob(data)
    blob_out = _BLOB()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    ok = fn(ctypes.byref(blob_in), None, None, None, None, 0x1, ctypes.byref(blob_out))
    if not ok:
        raise OSError('DPAPI 调用失败')
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def encrypt(text: str) -> str:
    return base64.b64encode(_dpapi(text.encode('utf-8'), True)).decode()


def decrypt(token: str) -> str:
    return _dpapi(base64.b64decode(token), False).decode('utf-8')


# ==================== 读写 ====================
def load_file():
    try:
        p = config_path()
    except Exception:
        return {}
    if not p or not os.path.exists(p):
        return {}
    try:
        with open(p, encoding='utf-8') as f:
            d = json.load(f)
    except Exception:
        return {}
    out = {}
    if d.get('base'):
        out['BASE'] = d['base']
    if d.get('user'):
        out['USER'] = d['user']
    if d.get('igid'):
        out['IGID'] = str(d['igid'])
    pwd = ''
    if d.get('pwd_enc'):
        try:
            pwd = decrypt(d['pwd_enc'])
        except Exception:
            pwd = ''
    elif d.get('pwd'):
        pwd = d['pwd']
    if pwd:
        out['PWD'] = pwd
    return out


def save_file(cfg, remember=True):
    """把账号写入配置文件；remember=True 时保存密码（DPAPI 加密）。"""
    p = config_path(writable=True)
    d = {'base': cfg.get('BASE', ''), 'user': cfg.get('USER', ''),
         'igid': str(cfg.get('IGID', '260')),
         'saved_at': datetime.datetime.now().isoformat(timespec='seconds')}
    if remember and cfg.get('PWD'):
        try:
            d['pwd_enc'] = encrypt(cfg['PWD'])
        except Exception:
            d['pwd'] = cfg['PWD']      # DPAPI 不可用时退化为明文
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    return p


def clear_password():
    p = config_path(writable=True)
    if not os.path.exists(p):
        return
    try:
        with open(p, encoding='utf-8') as f:
            d = json.load(f)
    except Exception:
        return
    d.pop('pwd_enc', None)
    d.pop('pwd', None)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


# ==================== 汇总 ====================
def resolve(cli=None):
    """返回 {'BASE','USER','PWD','IGID'}，只包含有值的项。"""
    out = {}
    try:
        f = load_file()
    except Exception:
        f = {}
    for k in ('BASE', 'USER', 'PWD', 'IGID'):
        if f.get(k):
            out[k] = f[k]
    for env, key in (('CCTV_BASE', 'BASE'), ('CCTV_USER', 'USER'),
                     ('CCTV_PWD', 'PWD'), ('CCTV_IGID', 'IGID')):
        if os.environ.get(env):
            out[key] = os.environ[env]
    for k, v in (cli or {}).items():
        if v:
            out[k] = v
    return out


def has_login(cfg):
    return bool(cfg.get('USER') and cfg.get('PWD'))

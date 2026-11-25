# -*- coding: utf-8 -*-
"""公共层：配置、登录会话、报告定位、数据页 URL、表格单元格解析。"""
import base64
import datetime
import html
import http.cookiejar
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

# ==================== 配置 ====================
# 优先级：命令行 --bgbh  >  同目录 config.py  >  这里的 DEFAULTS
#
# ⚠️ 只要 config.py 存在且定义了同名变量（本目录的 config.py 就定义了 BGBH），
#    改下面的 DEFAULTS 是**不生效**的。日常换报告请改 config.py 的 BGBH，
#    或直接 `python cctv_workflow.py --bgbh 26KHD00019 all`。
#    想确认当前到底用了哪个值：`python cctv_workflow.py --show-config`
DEFAULTS = {
    'BASE': 'http://36.134.56.138:7500',
    'USER': '',        # 启动时填写 / 配置文件 / 环境变量 CCTV_USER
    'PWD': '',         # 启动时填写 / 配置文件（DPAPI 加密） / 环境变量 CCTV_PWD
    'IGID': '260',          # 业务项：KHD-污水管道
    'BGBH': '26KHD00018',   # 仅当没有 config.py 时才作为兜底
}

# 报告台账查询的固定参数（2026-11 实测：缺 WTRQ/WTRQ2/fType/wtst 时 count=0）
PARA_NAMES = ('WTRQ,WTRQ2,iGid,GCMC,WTBH,BGBH,GCBH,YPBH,JCRQ,GGDJ,GCBW,'
              'ZZRQ,YHTJ,code,WTDW,dttype,flow,wtst,')
PCDE = '0069J'

# 4126：列表式，每页 13 行 × 21 列
ROWS_PER_PAGE_4126 = 13
COLS_4126 = 21

HEADERS_4126 = [
    '序号', '管段编号', '管径(mm)', '材质', '管道长度(m)', '水流方向',
    '管内底标高1(m)', '管内底标高2(m)', '有无缺陷',
    '结构性-平均值S', '结构性-最大值Smax', '结构性-缺陷等级', '结构性-缺陷密度',
    '结构性-修复指数RI', '结构性-综合状况评价',
    '功能性-平均值S', '功能性-最大值Smax', '功能性-缺陷等级', '功能性-缺陷密度',
    '功能性-养护指数MI', '功能性-综合状况评价',
]
WIDTHS_4126 = [6, 14, 8, 10, 10, 8, 12, 12, 8, 9, 11, 9, 9, 11, 40, 9, 11, 9, 9, 11, 40]

# 4127：卡片式，缺陷最多 4 槽位（超过部分会被丢弃，校验阶段会告警）
DEFECT_SLOTS = 4
DEFECT_FIELDS = ['长度(m)', '缺陷名称代码', '分值', '等级', '管道内部状况描述']

HEADERS_4127 = ['序号', '起始井号', '终止井号', '录像文件', '敷设年代', '起点埋深', '终点埋深',
                '管段类型', '管段材质', '管段直径', '检测方向', '管段长度', '检测长度',
                '修复指数', '养护指数', '检测人员']
for _i in range(DEFECT_SLOTS):
    HEADERS_4127 += ['缺陷%d-%s' % (_i + 1, x) for x in ['长度', '代码', '分值', '等级', '描述']]
HEADERS_4127.append('图片')
WIDTHS_4127 = [6, 10, 10, 40, 8, 8, 8, 12, 10, 8, 8, 8, 8, 8, 8, 8,
               8, 12, 8, 8, 24, 8, 12, 8, 8, 24, 8, 12, 8, 8, 24, 8, 12, 8, 8, 24, 8]

# 4123：汇总一览表，12 行/页，说明行 + 表头 + 6 列
ROWS_PER_PAGE_4123 = 12
HEADERS_4123 = ['序号', '管段编号', '管径(mm)', '管段材质', '结构性缺陷', '功能性缺陷']
WIDTHS_4123 = [6, 14, 10, 12, 40, 40]

# 4125：数量统计表，1 页 = 16 个缺陷行 + 合计
HEADERS_4125 = ['序号', '缺陷类别', '缺陷名称', '1级', '2级', '3级', '4级', '小计']
WIDTHS_4125 = [6, 12, 20, 8, 8, 8, 8, 8]

# 4122：主要工程量表，1 页 = 说明行 + 表头 + N 行 + 总计
HEADERS_4122 = ['序号', '名称', '管道长度m', '检测长度', '备注']
WIDTHS_4122 = [6, 30, 12, 12, 40]

# 五张表 → 输出文件名前缀
TABLE_NAMES = {
    '4122': '主要工程量表',
    '4123': '管道缺陷汇总一览表',
    '4125': '管道缺陷数量统计表',
    '4126': '管段状况评估表',
    '4127': '管段检测与评估成果表',
}

# 缺陷代码里代表「无缺陷」的占位值
NORMAL_CODES = ('(ZC)正常', 'ZC正常', '正常', '--', '-', '/', '')

# 缺陷代码 → 结构/功能分类（CJJ181-2012）
STRUCT_CODES = {'AJ', 'BX', 'CK', 'CR', 'FS', 'PL', 'QF', 'SL', 'TJ', 'TL'}
FUNC_CODES = {'CJ', 'CQ', 'FZ', 'JG', 'SG', 'ZW', 'ZAW', 'YWCR'}

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


OVERRIDES = {}


def load_config():
    """
    返回 (cfg, meta)。优先级（高 → 低）：
        命令行 --user/--pwd/--base（写入 OVERRIDES）
        环境变量 CCTV_USER / CCTV_PWD / CCTV_BASE
        配置文件 cctv_config.json（密码 DPAPI 加密；可用 CCTV_CONFIG 指定路径）
        同目录 config.py（旧版，保留兼容）
        内置 DEFAULTS
    """
    cfg = dict(DEFAULTS)
    meta = {'source': '内置 DEFAULTS', 'ignored': {}, 'file': None}
    try:
        import config as _cfg  # noqa: F401
    except Exception:
        _cfg = None
    if _cfg is not None:
        defined = False
        for k in cfg:
            if hasattr(_cfg, k):
                defined = True
                v = getattr(_cfg, k)
                if str(v) != str(cfg[k]):
                    meta['ignored'][k] = cfg[k]
                cfg[k] = v
        if defined:
            meta['source'] = 'config.py'
    try:
        from . import credentials as _cred
        cred = _cred.resolve(OVERRIDES)
        if cred:
            for k, v in cred.items():
                cfg[k] = v
            meta['source'] = '配置文件/环境变量/命令行'
            if os.environ.get('CCTV_CONFIG'):
                meta['file'] = os.environ['CCTV_CONFIG']
            else:
                try:
                    meta['file'] = _cred.config_path()
                except Exception:
                    pass
    except Exception as e:
        meta['cred_error'] = '%s: %s' % (type(e).__name__, e)
    return cfg, meta


def print_config(cfg, meta, bgbh=None):
    """打印本次实际生效的配置（不打印密码），并提示被 config.py 覆盖的 DEFAULTS 项。"""
    val, src = cfg['BGBH'], meta['source']
    if bgbh and bgbh != val:
        val, src = bgbh, '命令行 --bgbh'
    print('[配置] BASE=%s  IGID=%s  账号=%s  报告编号=%s（来源: %s）'
          % (cfg['BASE'], cfg['IGID'], cfg['USER'], val, src))
    if meta.get('file'):
        print('[配置] 账号配置：%s' % meta['file'])
    if meta['ignored']:
        print('[配置] 注意：cctv_workflow.py 内 DEFAULTS 的 %s 已被 config.py 覆盖，未生效'
              % '、'.join('%s=%s' % (k, v) for k, v in meta['ignored'].items()))


# ==================== 采集会话（S1） ====================
class Session:
    """登录态 + 带重试的 GET。页面为 GBK，接口为 UTF-8，由调用方指定。"""

    def __init__(self, cfg, retries=3, timeout=30, verbose=True):
        self.base = cfg['BASE'].rstrip('/')
        self.user = cfg['USER']
        self.pwd = cfg['PWD']
        self.retries = retries
        self.timeout = timeout
        self.verbose = verbose
        self.cj = http.cookiejar.CookieJar()
        self.op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cj))
        self.op.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')]
        self.logged_in = False

    # ---- S1 ----
    def login(self):
        if not self.user or not self.pwd:
            raise RuntimeError('未配置平台账号：请运行 `cctv.py login` 或启动时在弹窗里填写，'
                               '也可设置环境变量 CCTV_USER / CCTV_PWD')
        enc = lambda s: base64.b64encode(s.encode()).decode()  # noqa: E731
        data = urllib.parse.urlencode(
            {'username': enc(self.user), 'password': enc(self.pwd), 'code': 'AAAA'}).encode()
        req = urllib.request.Request(self.base + '/loginAction', data=data, method='POST')
        req.add_header('Content-Type', 'application/x-www-form-urlencoded')
        body = self._open(req).decode('utf-8', 'ignore')
        try:
            j = json.loads(body)
        except Exception:
            raise RuntimeError('登录返回非 JSON：%s' % body[:200])
        # 闸门 G1：必须 success=true，否则后面所有页面都会退化成登录页 → 空数据
        if not j.get('success'):
            raise RuntimeError('登录失败：%s' % body[:200])
        self.logged_in = True
        if self.verbose:
            print('    [S1] 登录成功（cookies: %s）'
                  % ','.join(c.name for c in self.cj))
        return self

    def _open(self, req):
        last = None
        for i in range(self.retries):
            try:
                return self.op.open(req, timeout=self.timeout).read()
            except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
                last = e
                time.sleep(1.0 + i)
        raise RuntimeError('请求失败（重试 %d 次）：%s\n%s'
                           % (self.retries, req.full_url, last))

    def get_raw(self, url, referer=None):
        req = urllib.request.Request(url)
        req.add_header('Referer', referer or (self.base + '/system/XiangMuNeiRong_FenXiang.jsp'))
        req.add_header('X-Requested-With', 'XMLHttpRequest')
        return self._open(req)

    def get(self, url, encoding='gbk', referer=None):
        return self.get_raw(url, referer).decode(encoding, 'ignore')


# ==================== S2 定位报告 ====================
def locate_report(sess, bgbh, igid, days=3650, verbose=True):
    """
    按报告编号定位 report_id / parentid / fid。

    2026-11 实测要点：/XiangMutaizhangJson 必须带
        dttype=wtdt & WTRQ/WTRQ2（委托日期区间） & fType=4 & wtst=0 & pcde
    缺任一项 → {"count":"0","data":[]}（core.find_report 当前就是缺这些，已失效）。
    返回的 data[i] 已含结构化的 report_id / parentid / fid，
    不需要再从 bgbh 的 toifrmae(...) 串里正则抠。
    """
    end = datetime.date.today()
    start = end - datetime.timedelta(days=days)
    qs = ('dttype=wtdt&WTRQ=%s&WTRQ2=%s&flow=10&BGBH=%s&WTBH=&YHTJ=&YPBH=&ZZRQ=&JCRQ='
          '&GCMC=&GCBH=&GCBW=&GGDJ=&fType=4&iGid=%s&wtst=0&pcde=%s'
          % (start.isoformat(), end.isoformat(), bgbh, igid, PCDE))
    url = '%s/XiangMutaizhangJson?ParaNames=%s&%s&page=1&limit=50' % (
        sess.base, PARA_NAMES, urllib.parse.quote(qs, safe='=&'))
    raw = sess.get_raw(url)
    # 2026-11 实测：该接口返回 GBK 字节（不是页面声明的 utf-8），
    # 用 utf-8 decode 会静默毁掉中文工程名，故先 GBK 再回落 UTF-8。
    body = None
    for enc in ('gbk', 'utf-8'):
        try:
            body = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if body is None:
        body = raw.decode('gbk', 'replace')
    try:
        j = json.loads(body)
    except Exception:
        raise RuntimeError('报告台账接口返回非 JSON：%s' % body[:200])

    rows = j.get('data') or []
    if not rows:
        raise RuntimeError(
            '未找到报告 %s（count=%s）。排查顺序：\n'
            '  1) 委托日期区间是否覆盖该报告（--days 调大，当前 %s ~ %s）\n'
            '  2) 报告是否属于 iGid=%s 业务项（KHD-污水管道）\n'
            '  3) 账号是否有该项目权限'
            % (bgbh, j.get('count'), start, end, igid))

    row = rows[0]
    rid = str(row.get('report_id') or '').strip()
    parentid = str(row.get('parentid') or '').strip()
    fid = str(row.get('fid') or '').strip()
    gcmc = str(row.get('gcmc') or '').strip()

    # 闸门 G2：RID 必须是 UUID，parentid 必须是数字，否则后面翻页 100% 拿到空壳
    if not re.fullmatch(r'[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-'
                        r'[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}', rid):
        raise RuntimeError('report_id 非法：%r' % rid)
    if not parentid.isdigit():
        raise RuntimeError('parentid 非法：%r' % parentid)
    if not fid.isdigit():
        fid = '412'

    info = {'bgbh': bgbh, 'report_id': rid, 'parentid': parentid, 'fid': fid,
            'gcmc': gcmc, 'igid': igid, 'count': j.get('count')}
    if verbose:
        print('    [S2] %s → report_id=%s parentid=%s fid=%s' % (bgbh, rid, parentid, fid))
        print('         工程：%s' % gcmc)
    return info


def locate_via_browser(bgbh, headless=True):
    """旧路径兜底：Playwright 走 GUI 找 RID（_debug_flow.py / _locate_report.py 的等价实现）。"""
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'old_scripts'))
    from _locate_report import find_report  # 延迟导入，避免无 Playwright 环境报错
    r = find_report(bgbh, headless=headless)
    if not r or not r.get('report_id'):
        raise RuntimeError('浏览器兜底定位也失败：%s' % bgbh)
    return {'bgbh': bgbh, 'report_id': r['report_id'], 'parentid': str(r['parentid']),
            'fid': '412', 'gcmc': '', 'igid': DEFAULTS['IGID'], 'count': None}


# ==================== 数据页 URL ====================
def page_url(sess, biaodan, report, page):
    return ('%s/system/XiangMuXianShi.jsp?BiaoDanId=%s&Yongtu=xianshi&id=%s'
            '&data=xianshi&Div=1&iGid=%s&ZhuBiaoId=%s&currentpage=%d'
            % (sess.base, biaodan, report['report_id'], report['igid'],
               report['parentid'], page))


def _data_table(html_text):
    m = re.search(r'<table[^>]*data-table[^>]*>(.*?)</table>', html_text, re.S)
    return m.group(1) if m else None


def _cell_text(td):
    # strip() 顺带清掉 &nbsp; 解出的 U+00A0（str.strip 认它是空白）
    return html.unescape(re.sub(r'<[^>]+>', '', td)).strip()




# ==================== 表单解析 / 写回辅助 ====================
EMPTY = {'', '/', '--', '—', '-'}
IGID = '260'


def same_value(a, b):
    """空值等价（'' / '/' / '--' 视为相同），其余按字符串比较。"""
    a = '' if a is None else str(a).strip()
    b = '' if b is None else str(b).strip()
    if a in EMPTY and b in EMPTY:
        return True
    return a == b

def parse_form(html_text):
    """解析页面 form 控件为 [[name, value], ...]（大小写不敏感，兼容 Name=/Value=）。"""
    out = []
    for m in re.finditer(r'<(input|select|textarea)\b([^>]*)>', html_text, re.I):
        attrs = m.group(2)
        nm = re.search(r'name\s*=\s*["\']?([^"\'\s>]+)', attrs, re.I)
        if not nm:
            continue
        val = (re.search(r'value\s*=\s*["\']([^"\']*)["\']', attrs, re.I)
               or re.search(r'value\s*=\s*([^\s>]+)', attrs, re.I))
        out.append([nm.group(1), val.group(1) if val else ''])
    return out


def put(form, name, val):
    """按字段名写入 form（就地修改），返回是否命中。"""
    for i, (n, _) in enumerate(form):
        if n == name:
            form[i][1] = '' if val is None else str(val)
            return True
    return False


# Session 写回相关方法（不修改原类定义，运行期挂载）
def _report_url(self):
    return self.base + '/reportShow.jsp?id=%s&BiaoDanId=412&iGid=%s' % (self.rid, self.igid)


def _view_url(self, bid, pg):
    return (self.base + '/system/XiangMuXianShi.jsp?BiaoDanId=%s&Yongtu=xianshi&id=%s'
            '&data=xianshi&Div=1&iGid=%s&ZhuBiaoId=%s&currentpage=%d'
            % (bid, self.rid, self.igid, self.pid, pg))


def _anniu_url(self, bid):
    return (self.base + '/system/XiangMuAnNiu.jsp?BiaoDanId=%s&id=%s&iGid=%s&ZhuBiaoId=%s'
            % (bid, self.rid, self.igid, self.pid))


def _edit_url(self, bid, pg):
    return (self.base + '/system/XiangMuXiuGai.jsp?BiaoDanId=%s&iGid=%s&ZhuBiaoId=%s'
            '&report_id=%s&Div=1&currentpage=%d' % (bid, self.igid, self.pid, self.rid, pg))


def _post(self, url, data, ref=None, enc='utf-8'):
    body = urllib.parse.urlencode(data, encoding='gbk', errors='replace').encode()
    req = urllib.request.Request(url, data=body, method='POST')
    req.add_header('Content-Type', 'application/x-www-form-urlencoded')
    req.add_header('Referer', ref or self.report_url())
    return self.op.open(req, timeout=120).read().decode(enc, 'ignore')


def _init_write(self, *a, **k):
    pass


for _name, _fn in (('report_url', _report_url), ('view_url', _view_url), ('anniu_url', _anniu_url),
                   ('edit_url', _edit_url), ('post', _post)):
    setattr(Session, _name, _fn)
Session.rid = None
Session.pid = None
Session.bgbh = None
Session.igid = '260'

# 模块级常量（供写回等模块直接引用）
BASE = DEFAULTS['BASE']
USER = DEFAULTS['USER']
PWD = DEFAULTS['PWD']

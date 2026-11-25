# CCTV 报告工具 原理讲解

> 适用版本：`cctvtools 1.0.0` / 图形界面 `CCTV报告工具.exe` / 命令行 `CCTV命令行.exe`  
> 代码规模：约 3700 行 Python（不含打包依赖与历史项目目录）  
> 文档日期：2026-11-25  
> 配套文件：`工具说明.md`（使用说明）、`使用说明.txt`（给最终用户）

---

## 0. 先看结论：这个程序到底在做什么

这是一个面向**城镇排水管道 CCTV 检测成果**的自动化工具。平台上有五张相互关联的报表，工具负责：

1. **登录平台**，按报告编号定位到具体工程；
2. **抓取五张表**并落盘为 Excel；
3. **交叉校验**五张表之间的段数、缺陷数、长度、说明文字是否自洽；
4. **由 4126 + 4127 重新生成** 4122 主要工程量表、4123 缺陷汇总表、4125 缺陷数量统计表；
5. **与平台现状逐字段对比**；
6. **把有差异的表写回平台**（HTTP 表单提交，不依赖浏览器）；
7. 可选：**批量**处理多个报告、生成**校验报告 xlsx**。

五张表的编号与含义：

| 表号 | 名称 | 版式 | 作用 |
|---|---|---|---|
| 4122 | 主要工程量表 | 说明 + 表头 + 明细 + 总计，10 行/页 | 按管段类型/材质/管径汇总长度 |
| 4123 | 管道缺陷汇总一览表 | 说明 + 表头 + 12 行/页 | 每段管道的结构/功能缺陷描述 |
| 4125 | 管道缺陷数量统计表 | 单页 16 个缺陷行 + 合计 | 16 类缺陷按 1–4 级计数 |
| 4126 | 管段状况评估表 | 21 列 × 13 行/页 | 每段管道的评估结论与指数 |
| 4127 | 管段检测与评估成果表 | 卡片式，1 管段/页 | 检测明细、4 个缺陷槽位、图片 |

> 4122 是“工程量”，4125 是“缺陷数量统计”。两者容易混淆，工具里是两张独立的表。

---

## 1. 总体架构

### 1.1 分层

```
┌─────────────────────────────────────────────────────────────┐
│ 界面层                                                       │
│  cctv_gui.py  (Tkinter 图形界面)                             │
│  cctv.py → cctvtools/cli.py  (argparse 命令行)               │
└───────────────────────────┬─────────────────────────────────┘
                            │ 调用
┌───────────────────────────▼─────────────────────────────────┐
│ 业务编排层                                                   │
│  workflow.py     一键全流程 S1–S6                            │
│  uploader.py     按勾选的表上传                              │
│  batch.py        批量处理多个报告                            │
│  verify_report.py 生成校验报告 xlsx                          │
│  generate.py     由 4126/4127 生成 4122/4123/4125            │
│  diff.py         平台 vs Excel 变更清单                      │
│  writeback.py    编辑页填值 → 保存（HTTP 写回）              │
└───────────────────────────┬─────────────────────────────────┘
                            │ 调用
┌───────────────────────────▼─────────────────────────────────┐
│ 基础层                                                       │
│  crawl.py        五张表的 HTML 解析、翻页、落盘、交叉校验     │
│  common.py       配置、Session、报告定位、表单解析、常量      │
│  credentials.py  配置文件查找、Windows DPAPI 加解密           │
└───────────────────────────┬─────────────────────────────────┘
                            │
              ┌─────────────┴─────────────┐
              ▼                           ▼
        平台 HTTP 接口                Excel (.xlsx)
        (GBK 页面 / UTF-8 接口)       openpyxl
```

### 1.2 模块依赖关系

- `common` 只依赖标准库和 `credentials`，是整个项目的地基。
- `crawl` 依赖 `common`，负责“平台 → 内存数据 → xlsx”。
- `generate` / `diff` / `writeback` 依赖 `common`，分别负责生成、比较、写回。
- `workflow` 把 `crawl`、`generate`、`writeback` 串起来。
- `uploader` 把 `crawl`、`generate`、`writeback` 串起来，但按“勾选的表”选择性执行。
- `verify_report` 依赖 `crawl.load_outputs` 和 `crawl.verify`，把校验结果写成带说明的 xlsx。
- `batch` 依赖 `workflow`，逐个报告跑并汇总。
- `cli` 依赖以上全部；`cctv_gui` 直接调用其中几个模块。

### 1.3 一次完整流程（S1–S6）

```
S1 登录 + 定位
      │  /loginAction → cookies
      │  /XiangMutaizhangJson → report_id / parentid
      ▼
S2 抓取五张表
      │  4126 列表翻页（13 行/页）
      │  4127 卡片翻页（1 管段/页，4 缺陷槽位）
      │  4122 / 4123 / 4125 三段式表
      ▼
S3 交叉校验
      │  管段集合、方向、重复、缺陷数、长度、说明文字
      ▼
S4 生成派生表
      │  4126 + 4127 → 4122 / 4123 / 4125
      ▼
S5 对比 + 写回
      │  平台现状 vs 生成结果
      │  只把有差异的表写回（编辑页锁定 → 填值 → 保存）
      ▼
S6 复核
      │  重新抓取 4122/4123/4125，确认与生成结果一致
      ▼
   输出目录：五张表 xlsx + 生成件 + 清单 JSON
```

---

## 2. 入口与运行模型

### 2.1 命令行入口 `cctv.py`

`cctv.py` 只有 8 行，作用是把 `cctvtools.cli.main` 暴露成脚本入口：

```python
import sys
from cctvtools.cli import main

if __name__ == '__main__':
    sys.exit(main())
```

打包成 `CCTV命令行.exe` 时，PyInstaller 的入口也是 `cctv.py`。

### 2.2 CLI 命令一览

`cli.build_parser()` 用 `argparse` 定义全局参数与子命令。全局参数可以写在子命令前后，靠 `_normalize()` 把 `--bgbh/--user/--pwd/--base` 重新挪到最前面。

| 子命令 | 作用 | 主要调用 |
|---|---|---|
| `locate` | 只定位报告 | `make_session` → `common.locate_report` |
| `crawl` | 抓一张或全部表 | `crawl.crawl_412x` + `save_412x` |
| `workflow` / `run` | 一键全流程 | `workflow.run` |
| `login` | 保存账号（DPAPI） | `credentials.save_file` |
| `verify-report` | 生成校验报告 xlsx | `verify_report.build` |
| `upload-tables` | 按表选择性上传 | `uploader.upload_selected` |
| `batch` | 批量处理 | `batch.run_batch` |
| `all` | 旧版兼容：抓五表 + 校验 + manifest | `crawl.run_all` |
| `verify` | 离线校验最新一套 | `crawl.verify_latest` |
| `generate` | 由 Excel 生成 4122/4123/4125 | `generate.build_from_rows` |
| `diff` | 平台 vs Excel 变更清单 | `diff.compare` |
| `upload` | 按 Excel 写回 4126/4127 | `writeback` |
| `selftest` | 离线自检（不联网） | `crawl.selftest` |

`cli.main()` 的流程：

1. `build_parser()` 建解析器；
2. `_normalize()` 把全局参数归位；
3. 把 `--user/--pwd/--base` 写入 `common.OVERRIDES`（最高优先级）；
4. `--show-config` 只打印配置；
5. 没有子命令且是终端 → 进入 `interactive()` 菜单；
6. 需要 `--bgbh` 的命令做参数校验；
7. 调用 `args.func(args)`。

### 2.3 图形界面入口 `cctv_gui.py`

`cctv_gui.App` 继承 `tk.Tk`，构造顺序：

1. 设置标题、尺寸 `980x640`、最小尺寸；
2. `self.q = queue.Queue()`，用于工作线程 → 主线程传日志；
3. `self._build()` 创建控件；
4. `self.after(120, self._poll)` 启动日志轮询；
5. `_ensure_login()`：没有登录信息就弹账号窗；
6. `mainloop()`。

界面分四块：

| 区域 | 内容 |
|---|---|
| 顶部表单 | 报告编号（默认 `26KHD00029`）、输出目录、浏览按钮 |
| 按钮行 | 一键全流程 / 全流程+上传 / 只抓取 / 只生成 4122/4123/4125 / 校验报告 / 上传勾选 / 批量 / 账号 |
| 上传范围 | 5 个复选框：4122、4126、4127、4123、4125 |
| 日志区 | 黑底白字 `tk.Text` + 横竖滚动条 |

按钮与动作：

| 按钮 | 方法 | 行为 |
|---|---|---|
| 一键全流程 | `do_workflow` → `_wf(False)` | `workflow.run(upload=False)` |
| 全流程+上传 | `do_workflow_upload` → `_wf(True)` | 二次确认后 `workflow.run(upload=True)` |
| 只抓取 | `do_crawl` | 登录后抓五张表并保存 |
| 只生成 4122/4123/4125 | `do_generate` | 读取目录里最新 4126/4127，生成三张表 |
| 校验报告 | `do_verify` | `verify_report.build` |
| 上传勾选 | `do_upload_tables` | 收集复选框 → `uploader.upload_selected` |
| 批量… | `do_batch` | 选列表 txt → `batch.run_batch` |
| 账号 | `change_account` | 弹出账号窗并更新标题 |

### 2.4 GUI 的线程模型

GUI 不允许在工作线程里直接改控件，所以采用“工作线程 + 队列 + 主线程轮询”：

```
用户点按钮
   │
   ▼
App.start(fn, title)
   │  running 置 True
   │  清空日志
   │  禁用按钮（包括 b_upload）
   │  threading.Thread(target=App._worker, args=(fn,), daemon=True).start()
   ▼
App._worker(fn)
   │  contextlib.redirect_stdout/stderr → QueueWriter(q)
   │  try: fn()
   │  except: traceback.print_exc()
   │  finally: 打印耗时
   │  q.put(('done', ok))
   ▼
App._poll()  （每 120ms）
   │  从队列取 ('log', s) → txt.insert
   │  取 ('done', ok) → _finish(ok)
   ▼
App._finish(ok)
   │  running 置 False
   │  恢复所有按钮（含 b_upload）
   │  标签显示“完成/失败”
```

- `QueueWriter.write()` 把标准输出的每一段文本放进队列；`flush()` 空实现。
- `_worker` 用 `contextlib.redirect_stdout/stderr` 捕获模块里的 `print`，所以业务代码不需要知道 GUI 的存在。
- `_poll` 用 `after(120, self._poll)` 形成常驻轮询。
- `_finish` 会恢复 `b_upload`，这是修复过的一个关键点：否则跑完任意任务后“上传勾选”会一直灰着。

### 2.5 打包后的路径处理

`cctv_gui.base_dir()`：

- 冻结（PyInstaller）运行时返回 `os.path.dirname(sys.executable)`（exe 所在目录）；
- 源码运行返回 `cctv_gui.py` 所在目录。

`default_out_dir()` 返回 `base_dir()/输出` 并自动创建。

冻结时设置 Tcl/Tk 数据目录：

```python
if getattr(sys, 'frozen', False):
    _base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    os.environ.setdefault('TCL_LIBRARY', os.path.join(_base, 'tcl8.6'))
    os.environ.setdefault('TK_LIBRARY', os.path.join(_base, 'tk8.6'))
```

PyInstaller 把 tcl/tk 数据放进 `_internal/tcl8.6`、`_internal/tk8.6`，这段保证 Tkinter 能找到它们。

### 2.6 打包脚本 `_build/build_exe.py`

| 参数/逻辑 | 说明 |
|---|---|
| `--onedir`（默认） | 目录版：exe + `_internal`，启动快、依赖全 |
| `--onefile` | 单文件版，运行时解压到临时目录 |
| `--windowed` | GUI 版不显示控制台 |
| `--console` | CLI 版显示控制台 |
| `--collect-submodules cctvtools` | 收集包内全部子模块 |
| `--add-binary` | 显式带 `_tkinter.pyd`、`pyexpat.pyd`、`_ctypes.pyd`、`ffi-*.dll`、`libexpat.dll`、`tcl86t.dll`、`tk86t.dll` |
| `--add-data` | 带 `tcl8.6`、`tk8.6` 数据目录 |
| `--exclude-module pkg_resources/setuptools` | 避免 `pyi_rth_pkgres` 启动报错 |
| `--paths BASE` | 保证 `cctvtools` 可导入 |
| `_vendor` + `PYTHONPATH` | PyInstaller 装在项目内，不污染全局环境 |

---

## 3. 配置、账号与登录会话

### 3.1 配置默认值 `common.DEFAULTS`

```python
DEFAULTS = {
    'BASE': 'http://36.134.56.138:7500',
    'USER': '',
    'PWD': '',
    'IGID': '260',          # 业务项：KHD-污水管道
    'BGBH': '26KHD00018',   # 仅当没有 config.py 时兜底
}
```

### 3.2 配置优先级

`common.load_config()` 的合并顺序（后者覆盖前者）：

1. `DEFAULTS`；
2. 同目录 `config.py`（旧版兼容，定义了同名变量就覆盖，并在 `meta['ignored']` 记录被覆盖项）；
3. `credentials.resolve(OVERRIDES)`：
   - 配置文件（`CCTV_CONFIG` → `%APPDATA%\CCTV工具\` → exe 同目录 → `~/.cctv/`）；
   - 环境变量 `CCTV_BASE/CCTV_USER/CCTV_PWD/CCTV_IGID`；
   - 命令行 `--base/--user/--pwd`（写入 `OVERRIDES`，最高）。

`print_config()` 只打印 `BASE/IGID/账号/报告编号/来源`，**不打印密码**。

### 3.3 配置文件查找 `credentials._candidates()`

按顺序：

| 顺序 | 路径 | 场景 |
|---|---|---|
| 1 | `%CCTV_CONFIG%` | 自定义/自动化 |
| 2 | `%APPDATA%\CCTV工具\cctv_config.json` | 冻结版正常安装 |
| 3 | exe 同目录 `cctv_config.json` | 便携使用 |
| 4 | `~/.cctv/cctv_config.json` | 兜底 |

`config_path(writable=False)` 返回第一个“存在”的候选；都不存在则返回第一个“可写”的候选。  
`config_path(writable=True)` 返回第一个可写候选，用于保存。

> 注意：`_writable()` 用 `open(path,'a')` 试探，会在磁盘上创建一个 0 字节文件。这是历史实现，打包时不要带上空配置文件，否则会挡住后面的候选。

### 3.4 Windows DPAPI 加密

密码不存明文，使用 Windows 的 `CryptProtectData / CryptUnprotectData`：

```
明文密码
  │ ctypes.windll.crypt32.CryptProtectData(..., 0x1, ...)
  ▼
DPAPI 密文（只有同一 Windows 用户能解）
  │ base64
  ▼
"pwd_enc" 字段写入 JSON
```

- `_BLOB`：`ctypes.Structure`，对应 `DATA_BLOB { DWORD cbData; char* pbData; }`。
- `_make_blob()`：把 bytes 放进 `ctypes.create_string_buffer`，返回 `_BLOB` 和缓冲区（防止被 GC）。
- `_dpapi(data, protect)`：调用 `crypt32` 的加/解密函数，`LocalFree` 释放输出内存。
- `encrypt(text)` / `decrypt(token)`：DPAPI + base64 包装。
- `load_file()`：读 JSON，兼容 `pwd_enc`（加密）与 `pwd`（明文旧格式）。
- `save_file(cfg, remember=True)`：写 `base/user/igid/saved_at`；`remember` 才写密码；DPAPI 不可用时退化为明文 `pwd`。
- `clear_password()`：删掉密码字段，保留账号。
- `resolve(cli)`：文件 → 环境变量 → 命令行，只返回有值的键。
- `has_login(cfg)`：`USER` 与 `PWD` 都非空才算已登录。

### 3.5 `common.Session`：登录态与 HTTP

`Session` 封装 `urllib` + `http.cookiejar`：

| 成员 | 作用 |
|---|---|
| `self.cj` | `CookieJar`，保存 JSESSIONID / username |
| `self.op` | `build_opener(HTTPCookieProcessor(cj))`，统一带 Cookie |
| `self.op.addheaders` | 伪装浏览器 `User-Agent` |
| `self.base` | 平台地址（去掉末尾 `/`） |
| `self.retries` / `timeout` | 重试次数（默认 3）/ 超时（默认 30s） |

方法：

- `login()`：
  - 账号/密码为空直接抛错；
  - `username/password` 各自 **base64** 后 `urlencode`，POST `/loginAction`；
  - 响应按 UTF-8 解码，必须是 JSON 且 `success=true`（闸门 G1）；
  - 成功置 `logged_in=True`。
- `_open(req)`：最多重试 `retries` 次，捕获 `URLError/HTTPError/TimeoutError/OSError`，间隔 `1.0+i` 秒；全失败抛 `RuntimeError`。
- `get_raw(url, referer)`：GET 原始 bytes，带 `Referer`（默认 `XiangMuNeiRong_FenXiang.jsp`）和 `X-Requested-With: XMLHttpRequest`。
- `get(url, encoding='gbk', referer)`：`get_raw().decode(encoding, 'ignore')`；平台页面是 GBK。

### 3.6 报告定位 `common.locate_report()`

调用 `/XiangMutaizhangJson`，关键参数：

```
ParaNames=WTRQ,WTRQ2,iGid,GCMC,WTBH,BGBH,GCBH,YPBH,JCRQ,GGDJ,GCBW,ZZRQ,YHTJ,code,WTDW,dttype,flow,wtst,
dttype=wtdt
WTRQ / WTRQ2 = 今天-N天 / 今天      # 委托日期区间
fType=4
wtst=0
pcde=0069J
page=1&limit=50
```

响应解码：先 GBK，失败再 UTF-8（实测接口返回 GBK）。  
取 `data[0]` 的 `report_id / parentid / fid / gcmc`，并做闸门 G2：

- `report_id` 必须是 UUID 形式；
- `parentid` 必须是纯数字；
- `fid` 非数字时兜底为 `'412'`。

失败时给出排查顺序（日期区间、iGid 业务项、账号权限）。

`locate_via_browser()` 是旧版兜底：把 `old_scripts` 加入 `sys.path`，延迟导入 `_locate_report.find_report`，用 Playwright 走浏览器找 RID。默认不启用，只有 `--browser-fallback` 才会走。

---

## 4. 平台协议与五张表的解析

### 4.1 页面 URL

`common.page_url(sess, biaodan, report, page)`：

```
{base}/system/XiangMuXianShi.jsp
  ?BiaoDanId={4122/4123/4125/4126/4127}
  &Yongtu=xianshi&id={report_id}&data=xianshi&Div=1
  &iGid={igid}&ZhuBiaoId={parentid}&currentpage={page}
```

`currentpage` 从 **0** 开始。平台越界页返回“空壳”（没有 data-table 或全空行），这是翻页停止的依据。

### 4.2 通用 HTML 工具

- `_data_table(html_text)`：用正则找 `<table ... data-table ...>...</table>`，返回表体 HTML；没有则 `None`。
- `_cell_text(td)`：
  - 去掉标签；
  - `html.unescape` 还原 `&nbsp;` 等实体；
  - 去掉 `\xa0`、首尾空白；
  - 空串兜底为 `'/'`。

> 平台页面大量使用 `rowspan`/`colspan`，但正则按 `<tr>`/`<td>` 逐行取，`rowspan` 的“合并单元格”不会自动重复到下一行——4127 的缺陷块正是利用这一点按“标签行 + 值行”读取。

### 4.3 4126 管段状况评估表

版式：每页 13 行 × 21 列。

`parse_4126(html_text)`：

1. 取 `data-table`；
2. 按 `<tr>` 拆行，按 `<td>` 拆列；
3. `len(tds) != 21`：
   - 11/12 列 → 表头行，`header_rows += 1`；
   - 其它 → 结构异常，`malformed_rows += 1`；
4. 第一列必须是数字（序号），否则算异常；
5. 第二列（管段编号）为空、`/`、`—`、`-` → 空预留行，`placeholder_rows += 1`；
6. 其余行保留 21 个单元格。

返回 `(rows, meta)`，`meta` 用于诊断“页面结构漂移”。

`crawl_4126()`：

```
page = 0
while page < max_pages:
    rows_page, meta = parse_4126(GET page_url(..., page))
    if not rows_page: stop='empty'; break
    按 _seg_key(管段编号) 去重
    if 整页都是重复: stop='repeat'; break
    累积 rows
    page += 1
```

停页原因三选一：`empty`（越界空页）、`repeat`（服务端重复吐最后一页）、`cap`（到最大页数）。

### 4.4 4127 管段检测与评估成果表

版式：1 管段/页的卡片；头部是“标签 → 右邻值”平铺，缺陷区是 4 个纵向槽位：

```
长度(m) | 1m   | 缺陷名称代码 | (CJ)沉积 | 分值 | 2
        | 1.2m |              | (BX)变形 |      | 0.5
        | /    |              | /        |      | /
        | /    |              | /        |      | /
等级    | 2    | 描述         | ...      | 图片 | 图1
```

`parse_4127(html_text)`：

1. 取 `data-table`，把全部 `<td>` 拉平成一维 `flat`；
2. `val(label)`：在 `flat` 里找标签，取右邻值；找不到返回 `'/'`；
3. `起始井号` 为空或 `/` → 越界页，返回 `None`；
4. 缺陷块扫描：
   - 若某行包含缺陷标签（`DEFECT_LABELS`），把该行的非标签值按顺序与标签配对，作为一块的“首行”；
   - 后续“列数等于标签数”的行，按列序追加到对应标签数组；
   - 其它行结束当前块；
5. 组装记录：
   - 16 个头部字段；
   - `长度(m)/缺陷名称代码/分值/等级/管道内部状况描述` 各 4 个槽位；
   - `图片` 只保留非空、非 `/`、非 `--` 的值。

`crawl_4127()`：

- 逐页解析；
- 用 `(起始井号, 终止井号)` 去重；
- 越界页 `parse_4127` 返回 `None` → `empty`；
- 整页重复 → `repeat`；
- 统计 `max_defects_per_pipe`（单管最多缺陷数）。

### 4.5 4122 主要工程量表

版式：说明行（1 个 td）+ 表头（5 列）+ 工程量行（5 列）+ 总计行（4 或 5 个 td，序号/名称被 colspan 合并）。

`parse_4122()`：

- 1 个 td 且含“共检测/段管道” → 说明；
- 表头 `序号/名称/管道长度m` → 跳过；
- 第一列 `总计`：
  - 4 个 td：`length=cells[1], detect=cells[2], note=cells[3]`；
  - ≥5 个 td：`length=cells[2], detect=cells[3], note=cells[4]`；
- 5 列且第一列数字、第二列非空 → 明细行，数字列用 `_num()` 转 int/float。

`crawl_4122()`：

- 第 0 页有数据，之后是空壳页；
- 停页条件：整页既没有名称非空的行，也没有带值的总计；
- 统计备注里的“分N段”之和（`_segments_in_note`），用于 G9 对账。

### 4.6 4123 管道缺陷汇总一览表

版式：说明行 + 表头（6 列）+ 数据行（6 列），12 行/页。

`parse_4123()`：

- 1 个 td 且含“共评估/段管道” → 说明；
- 6 列、第一列数字、第二列不是空/`/`/`管段编号` → 明细行。

`crawl_4123()`：

- 按管段编号去重翻页；
- 整页无数据 → `empty`；
- 整页重复 → `repeat`；
- 返回 `(summary, rows, stats)`。

### 4.7 4125 管道缺陷数量统计表

版式：单页 16 个缺陷行 + 合计行；缺陷类别单元格用 `rowspan` 合并：

- 8 个 td：`序号 | 类别 | 缺陷名称 | 1级..4级 | 小计`
- 7 个 td：`序号 | 缺陷名称 | 1级..4级 | 小计`
- 合计行 7 个 td：`'' | 合计 | 1级..4级 | 小计`

`parse_4125()`：

- 1 个 td 且含“共评估/段管道” → 说明；
- 表头 `序号/缺陷名称` → 跳过；
- `cells[1] == '合计'` → 合计；
- 第一列数字、最后一列非空 → 缺陷行；
- 8 列时第 2 列是类别，7 列时沿用上一行的类别。

`crawl_4125()`：

- 最多 5 页；
- 无有效缺陷行 → `empty`；
- 缺陷行 ≥16 → `full`（整表抓完）。

### 4.8 落盘：五张表的 Excel 版式

`crawl._write_sheet()` 是 4126/4127 的通用写法：

- 第 1 行表头，加粗、居中、自动换行、细边框；
- 第 2 行起数据，9pt；
- 设置列宽；
- 冻结窗格（4126 冻 `B2`，4127 冻 `C2`）。

4122/4123/4125 是“三段式”：

- `_summary_block()`：说明行合并 `A1:<ncols>1`，左对齐 10pt；返回表头行号（有说明为 3，无说明为 1）。
- `_table_body()`：表头写在第 `hr` 行，表体紧随其后；返回表体下一行行号。
- `_write_total()`：写合计/总计行；`merge_to` 指定与第 1 列合并到哪一列（4122 的总计是 `A:B`）。

保存函数：

| 函数 | 工作表名 | 说明 |
|---|---|---|
| `save_4126(rows, path)` | 管段状况评估表 | 21 列表头 |
| `save_4127(records, path)` | 管段检测与评估成果表 | 16 头部 + 4×5 缺陷槽 + 图片 = 37 列 |
| `save_4122(summary, rows, total, path)` | 主要工程量表 | 说明/表头/明细/总计 |
| `save_4123(summary, rows, path)` | 管道缺陷汇总一览表 | 说明/表头/明细 |
| `save_4125(summary, defects, total, path)` | 管道缺陷数量统计表 | 说明/表头/16 行/合计 |

`output_name(table, bgbh, ts)` 统一命名：

```
{TABLE_NAMES[table]}-{bgbh}(网页爬取)_{YYYYMMDD_HHMM}.xlsx
```

`sha256(path)` 分块（1MB）计算文件哈希，用于 manifest。

---

## 5. 交叉校验 G1–G9

### 5.1 为什么要校验

平台五张表是分别录入/生成的，容易出现：

- 4126 有 163 段，4127 只有 162 段；
- 同一段在 4126 是 `W47-W46`，在 4127 是 `W46 → W47`（流向相反）；
- 4123 汇总表漏了一段；
- 4125 合计与 4127 缺陷条目数对不上；
- 4122 备注“分51段”但 4126 实际 50 段；
- 说明行是平台写死的旧文字。

`crawl.verify(rows_4126, recs_4127, extra)` 做这些对账。

### 5.2 管段集合匹配算法

难点：4126 是单列 `W1-W2`，4127 是起始井号/终止井号两列；井号本身含 `-`（`W66-1`、`W10-1-W10-2`），**不能** `split('-')`。

算法：

```
a_directed = {规范化(4126.管段编号)}
pairs      = [(规范化(起), 规范化(止)) for 4127]
b_directed = {"起-止"}
rev_of     = {"止-起": (起, 止)}
pair_of    = {"起-止": (起, 止)}

only_a    = a_directed - b_directed
flipped   = [seg for seg in only_a if seg in rev_of]      # 方向相反
unmatched_a = only_a - flipped
unmatched_b = b_directed - a_directed，排除已由反向解释的
```

- 正向对不上、反向能对上 → `flipped`（WARN，不 FAIL）；
- 两边都解释不了 → `unmatched`（FAIL）；
- 重复数 = 总行数 - 去重集合大小。

### 5.3 检查项清单

| ID | 级别 | 含义 |
|---|---|---|
| G3-4126-nonempty | FAIL/PASS | 4126 是否抓到数据 |
| G4-4127-nonempty | FAIL/PASS | 4127 是否抓到数据 |
| G3-no-dup-4126 | FAIL/PASS | 4126 重复管段数 |
| G4-no-dup-4127 | FAIL/PASS | 4127 重复管段数 |
| G5-count-match | FAIL/PASS | 两表段数是否一致 |
| G5-direction-flip | WARN/PASS | 流向相反的段 |
| G5-unmatched | FAIL/PASS | 仅一边有的段 |
| G6-defect-slots | WARN/PASS | 单管缺陷超过 4 槽位（会被截断） |
| G7-4123-vs-4127 | FAIL/PASS | 4123 缺陷段数 vs 4127 缺陷段数 |
| G7b-4123-in-4126 | FAIL/PASS | 4123 的段能否在 4126 找到 |
| G7b-4123-flip | WARN | 4123 与 4126 方向相反 |
| G7c-4123-type | WARN | 4123 描述分类 vs 4125 分类 |
| G7-4123-summary | WARN/PASS | 4123 说明文字段数 vs 表体行数 |
| G8-4125-vs-4123 | FAIL/PASS | 4125 合计 vs 4123 描述条数 |
| G8-4125-vs-4127 | WARN/PASS | 4125 合计 vs 4127 缺陷条目数 |
| G8-4125-summary | WARN/PASS | 4125 说明文字处数 vs 合计行 |
| G8c-4125-type | WARN/PASS | 4125 分类 vs 4127 代码分类 |
| G8-unknown-code | WARN | 未识别缺陷代码 |
| G9-4122-segments | WARN/PASS | 4122 备注“分N段”合计 vs 4126 行数 |
| G9-4122-length | WARN/PASS | 4122 总计长度 vs 4126 长度合计（5m 或 0.5% 容差） |
| G9-4122-summary | WARN/PASS | 4122 说明段数 vs 4126 行数 |

`verify()` 返回：

```python
{
  'checks': [...],
  'direction_flipped': [...],
  'only_4126': [...], 'only_4127': [...],
  'dup_4126': n, 'dup_4127': n,
  'defect_pipes': n, 'defect_entries': n,
  'hard_fail': any(level == 'FAIL')
}
```

`run_all()` 还会把采集阶段的 `malformed_rows`、`duplicate_rows` 作为 WARN 追加进 checks。

### 5.4 校验报告 `verify_report.build()`

输入：`bgbh`、`out_dir`。  
步骤：

1. `crawl.load_outputs()` 读最新一套五张表；
2. 没有 4126/4127 → 抛错；
3. `crawl.verify()`；
4. 计算总体结论：有 FAIL → `FAIL`；否则有 WARN → `WARN`；否则 `PASS`；
5. 生成 `校验报告-<编号>-<时间>.xlsx`，四个 sheet：
   - **校验结论**：报告编号、时间、目录、总体结论、检查项统计、各表行数/段数/合计、结论说明；
   - **检查明细**：序号、检查项、级别、结果、含义、处理建议（`CHECK_DOC` 映射）；
   - **问题清单**：只列 WARN/FAIL，附处理建议；无问题写“未发现问题”；
   - **文件清单**：五张表的文件名、大小、修改时间；缺失项标红。

返回 `{'path', 'verdict', 'checks', 'files'}`。

---

## 6. 生成：由 4126 + 4127 得到 4122 / 4123 / 4125

### 6.1 数据来源

生成器的输入是两组“行”：

- `rows4126`：21 列字符串，来自 4126 xlsx 或本次抓取；
- `rows4127`：37 列字符串，来自 4127 xlsx 或本次抓取。

`generate.build_from_rows(rows4126, rows4127)` 返回一个字典：

```python
{
  's22': '本次共检测163段管道,评估163段管道,详见下表:',
  'rows22': [['1', '污水管道（双壁波纹管）', '600.1', '600.1', '分51段(管径300mm)'], ...],
  'total22': {'length': '1608.8', 'detect': '1608.8', 'note': '共163段'},

  's23': '本次共评估163段管道,其中30段管道存在缺陷（...）,详见下表:',
  'rows23': [['1', 'W-15-W-21', '300', '双壁波纹管', '纵向2.49m处,存在1级变形。', '/'], ...],

  's25': '本次共评估163段管道,其中存在缺陷30处（结构性缺陷27处,功能性缺陷3处）,详见下表:',
  'rows25': [['BX', [16, 1, 0, 0], 17], ...],   # 16 项
  'total25': [29, 1, 0, 0],

  'defects': [...],   # 缺陷段明细
  'n_total': 163,
}
```

### 6.2 4122 主要工程量表算法

`generate.build_4122(rows4126, rows4127)`：

**第一步：建立 4126 索引**

```
by_seg[去空白(4126.管段编号)] = 4126 行
```

**第二步：遍历 4127，按分组累加**

对每个 4127 记录：

1. 用 `起始井号` / `终止井号` 正反拼出 `A-B` / `B-A`，去 4126 索引里找对应行（解决流向相反）；
2. 长度：
   - 匹配到 4126 → 取 `r[4]`（管道长度(m)）；
   - 匹配不到 → 取 4127 的 `r[11]`（管段长度，形如 `3.8m`，由 `_to_float` 去掉 `m`）；
3. 检测长度取 4127 的 `r[12]`；
4. 分组键 = `(管段类型, 管段材质, 管段直径)`：
   - `管段类型` 用 `_pipe_name()` 把 `污水（WS）` 转成 `污水管道`；
   - 材质原样；
   - 直径原样（`300mm`）；
5. 每组累加段数、长度、检测长度。

**第三步：排序**

按组长度**从大到小**排序（与平台现有 4122 的排法一致），长度相同再按分组键排序。

**第四步：格式化**

- 名称：`污水管道（双壁波纹管）`；
- 长度/检测长度：`_fmt()` 保留 1 位小数，整数去掉 `.0`（`1608.8` / `600` / `4.4`）；
- 备注：`分{N}段(管径{直径})`；
- 总计：
  - `length` = 所有组长度和；
  - `detect` = 所有组检测长度和；
  - `note` = `共{N}段`；
- 说明行：`本次共检测{N}段管道,评估{N}段管道,详见下表:`。

`n_total` 取 `rows4126` 中有效行数，和 4123/4125 的口径一致。

**第五步：比较 `diff_4122()`**

把平台当前 4122 与生成结果都转成 `(名称, 长度, 检测长度, 备注)` 的多重集合：

- 忽略行顺序；
- 长度/检测长度按 0.1 容差比较；
- 再比总计长度、总计检测长度、总计备注、说明行；
- 返回 `(新增行数, 删除行数, 是否完全一致)`。

这样即使平台行序不同，只要内容一致就“跳过”，不会做无意义的全文重写。

### 6.3 4123 管道缺陷汇总一览表算法

`build_from_rows()` 内：

1. `b = {'起|止': 4127行}`；
2. 遍历 4126 的每段，用 `起-止` / `止-起` 在 `b` 里找对应 4127 卡片；
3. 读 4 个缺陷槽位：
   - 长度 `rb[16+5i]`
   - 代码 `rb[17+5i]`
   - 等级 `rb[19+5i]`
   - 只保留 `_real()` 为真的槽位（不是空/`/`、不含 `(ZC)`）；
4. 按代码分类：
   - 结构缺陷（`STRUCT_CODES`）→ 描述放进“结构性缺陷”列；
   - 功能缺陷（`FUNC_CODES`）→ 放进“功能性缺陷”列；
5. 每槽位生成一句：

```
纵向{长度}处,存在{等级}级{中文名}。
```

6. 多槽位用 `；` 连接；没有则为 `/`；
7. 说明行统计：
   - `n_total`：4126 有效段数；
   - `n_def`：有缺陷段数；
   - `n_s`：有结构缺陷段数；
   - `n_f`：有功能缺陷段数；
   - `n_both`：同时有结构+功能的段数；
   - `n_total - n_def`：无缺陷段数。

### 6.4 4125 管道缺陷数量统计表算法

1. 遍历缺陷段，对每个真实缺陷槽位，`Counter[(代码, 等级)] += 1`；
2. 按固定顺序输出 16 行：
   - 结构：`AJ, BX, CK, CR, FS, PL, QF, SL, TJ, TL`
   - 功能：`CJ, CQ, FZ, JG, SG, ZW`
3. 每行 `[code, [1级,2级,3级,4级], 小计]`；
4. `total25[j] = 各行第 j 级之和`；
5. 说明行：
   - 总处数 = `sum(total25)`；
   - 结构处数 = 结构代码行的小计和；
   - 功能处数 = 功能代码行的小计和。

缺陷代码中文名在 `generate.CN`，分类集合在 `common.STRUCT_CODES/FUNC_CODES`。

### 6.5 写生成件

`generate.write_outputs(data, bgbh, outdir, ts)`：

| 文件 | 名称 |
|---|---|
| 4123 | `管道缺陷汇总一览表-{bgbh}(网页爬取)_{ts}.xlsx` |
| 4125 | `管道缺陷数量统计表-{bgbh}(网页爬取)_{ts}.xlsx` |
| 4122 | `主要工程量表-{bgbh}(生成)_{ts}.xlsx` |

> 4122 用 `(生成)` 后缀，避免覆盖 S2 抓取下来的 `(网页爬取)` 原表；`load_outputs` 只匹配 `(网页爬取)`，所以生成件不会被误当成平台现状。

`write_4122()` 直接复用 `crawl.save_4122()` 的版式，保证生成件与抓取件列宽/合并/总计行一致。

---

## 7. 对比与写回

### 7.1 `diff.py`：平台 vs Excel 的逐字段比较

`read_workbook(path)`：

- 打开 Excel；
- 找第一张表头含 `管段编号` 且 ≥21 列的工作表 → 4126；
- 找第一张表头含 `起始井号` 且 ≥37 列的工作表 → 4127；
- 找不到抛 `RuntimeError`。

`compare(plat26, plat27, rows26, rows27, alias)`：

- `plat26 = {管段编号: 21列行}`；
- `plat27 = {'起|止': 37列行}`；
- 4126：按 `_COL26`（第 2–21 列 → `c_601`…`c_614`）逐字段比；
- 4127：前 16 列按 `_COL27` 比，后 20 列按 4 个槽位 `_SLOT` 比；
- `alias`：平台编号 → 文件编号的纠正映射（如 `Y16-Y17 → Y15-Y17`），先按 alias 比，再跳过 alias 的目标，避免重复；
- 每个差异生成一条：

```python
{'table': '4126', 'seg': ..., 'newseg': ..., 'code': 'c_602',
 'field': '材质', 'old': ..., 'new': ...}
```

`write_list(changes, path, bgbh)`：

- Sheet `变更清单`：序号、表、管段、字段、平台现值、新值；
- Sheet `汇总`：4126/4127 的段数与变更条数，按字段计数；
- 冻结首行、列宽、边框。

### 7.2 写回的核心：进入编辑态 → 填表单 → 保存

平台不是 REST API，而是老式 JSP 表单。写回分三步：

#### 7.2.1 `open_edit(sess, bid, pg)`：进入编辑态并取编辑页 HTML

```
GET reportShow.jsp
GET XiangMuXianShi.jsp?...&currentpage=pg
GET XiangMuAnNiu.jsp?...
POST XiangMuCaoZuo   （锁定/创建编辑态）
      ID = "{pg}${igid}${bid}${pid}${rid}$"
      addPara=true, Type=4, locking=true
解析 JSON：success、url
GET {base}/system/{url}   （GBK 解码）
```

- `success != true` → 抛“进入编辑态失败”；
- 编辑页的字段名形如 `c_601_1`、`c_702_1`，后缀是行号。

#### 7.2.2 `save_page(sess, bid, pg, form)`

```
POST {base}/XiangMuXiuGaiBaoCun2
     body = 整个 form 的 name/value（GBK URL 编码）
     Referer = edit_url
返回文本包含 "XiangMuXianShi.jsp" 视为保存成功
```

平台保存后通常跳回查看页，所以用这个标记判断。

#### 7.2.3 `apply_page(sess, bid, pg, row_fields, dry, expect)`

1. `parse_form(open_edit(...))` → `[[name, value], ...]`；
2. `idx = {name: 下标}`；
3. `expect` 防错位：
   - 4126：检查 `c_601_{row}` 是否等于期望管段编号；
   - 4127：检查 `c_702_1`/`c_703_1` 是否是期望的起止井号；
4. 对每个 `row_fields[row][code]`：
   - 字段名 `code_row`；
   - 用 `same_value()` 判断是否真的需要改（空值 `/`、`--`、`''` 视为等价）；
   - 不同则就地改 form 的值，`changed += 1`；
5. 缺字段 → 抛错，避免“只改了一半”；
6. `dry` → 只打印，不保存；
7. 否则 `save_page()`，失败抛错。

### 7.3 字段映射

#### 4126（21 列）

| xlsx 列 | 字段 | 含义 |
|---|---|---|
| 2 | c_601 | 管段编号 |
| 3 | c_615 | 管径 |
| 4 | c_602 | 材质 |
| 5 | c_616 | 管道长度 |
| 6 | c_620 | 水流方向 |
| 7 | c_617 | 管内底标高1 |
| 8 | c_618 | 管内底标高2 |
| 9 | c_619 | 有无缺陷 |
| 10–15 | c_603–c_608 | 结构性 S/Smax/等级/密度/RI/评价 |
| 16–21 | c_609–c_614 | 功能性 S/Smax/等级/密度/MI/评价 |

#### 4127（37 列）

- 2–16 列：`c_702, c_703, c_701, c_704 … c_715`（起止井号、录像、年代、埋深、类型、材质、直径、方向、长度、检测长度、修复/养护指数、检测人员）；
- 4 个缺陷槽位：
  - 槽 1：`c_716..c_720`
  - 槽 2：`c_729..c_733`
  - 槽 3：`c_735,c_737,c_739,c_741,c_743`
  - 槽 4：`c_736,c_738,c_740,c_742,c_744`
  - 每槽依次是 长度/代码/分值/等级/描述；
- 图片字段 `c_725` 等不在自动写回范围（工具只写文字）。

#### 4122（10 行/页）

| 字段 | 含义 |
|---|---|
| `c_205_1` | 说明行 |
| `c_201_N` | 名称 |
| `c_202_N` | 管道长度 |
| `c_203_N` | 检测长度 |
| `c_204_N` | 备注 |
| `c_206_1` | 总计长度 |
| `c_207_1` | 总计检测长度 |
| `c_208_1` | 总计备注（共N段） |

#### 4123（12 行/页）

- `c_306_1`：说明行；
- `c_301_N`：管段编号；
- `c_302_N`：管径；
- `c_303_N`：材质；
- `c_304_N`：结构性缺陷；
- `c_305_N`：功能性缺陷；
- 空行 `Recid=0`，填值保存即新增行。

#### 4125（单页 16 行 + 合计）

- `c_501_1`：说明行；
- 第 r 行：`c_502_r..c_505_r`（1–4 级），`c_506_r`（小计）；
- 合计：`c_507_1..c_510_1`（1–4 级），`c_511_1`（合计小计）。

### 7.4 `diff_rows()`：4126/4127 的写回计划

输入：

- `rows_4126` / `recs_4127`：来自文件的“目标值”；
- `plat_4126`：`{管段编号: 行}`；
- `plat_4127`：`{'起|止': 行}`；
- `alias`。

输出 `(plan26, plan27)`：

- `plan26 = {平台段号: {字段: 新值}}`；
- `plan27 = {'起|止': {字段: 新值}}`。

只收“有差异”的段；文件里有、平台没有的段不会出现在 plan 里（4126/4127 不支持自动新增行）。

### 7.5 定位与批量写回

`scan_4126_positions(sess, pages=13)`：

- 逐页 `open_edit('4126', pg)`；
- `parse_form` 后找 `c_601_1..c_601_13`；
- 建 `{管段编号: (页, 行)}`；
- `uploader` 现在按平台实际页数传入 `pages`，不再固定扫 13 页。

`upload(sess, bid, plan, positions, dry)`：

1. 把 `{段: {字段:值}}` 按 `(页, 行)` 分组；
2. 逐页 `apply_page()`；
3. 每页保存后 `sleep(0.3)`，降低平台压力。

### 7.6 派生表写回

- `upload_4122(sess, summary, rows, total, dry, page_size=10)`：
  - 每页写 10 行 `c_201..c_204`；
  - 第 0 页写说明 `c_205` 与总计 `c_206..c_208`；
  - 生成行数少于平台时，把多余行写成空串清空占位；
  - 缺字段直接抛错，避免半写。
- `upload_4123(sess, summary, rows, dry, page_size=12)`：
  - 第 0 页写说明；
  - 每页 12 行，写 `c_301..c_305`；
  - 空行填值会新增记录。
- `upload_4125(sess, summary, rows16, total, dry)`：
  - 写说明、16 行 4 个等级 + 小计、合计 4 个等级 + 合计小计。

### 7.7 `uploader.upload_selected()`：按勾选上传

入口参数：

```python
upload_selected(bgbh, out_dir, tables, dry=False, alias=None, log=print)
```

`tables` 是 `{'4122','4123','4125','4126','4127'}` 的子集。

流程：

1. 读配置、登录、定位报告；
2. `crawl.load_outputs()` 读目录里最新一套 4126/4127（以及已有的 4122/4123/4125 供对比）；
3. 4126/4127：
   - 抓平台现状；
   - `writeback.diff_rows()` 生成计划；
   - 有差异才 `scan + upload`；一致则“跳过”；
4. 4122/4123/4125：
   - `generate.build_from_rows()` 生成目标值；
   - 4122：`generate.diff_4122()` 比较；
   - 4123：按 `(管段编号, 结构性, 功能性)` 集合比较；
   - 4125：比较合计与说明；
   - 一致则跳过，有差异才写回；
5. 返回：

```python
{'uploaded': ['4122', ...], 'skipped': [...], 'failed': []}
```

`dry=True` 时只打印“将写多少行/段”，不调用保存。

---

## 8. 工作流与批量

### 8.1 `workflow.run()`：一键全流程

签名：

```python
run(bgbh, out_dir='.', upload=True, dry=False, yes=False,
    from_xlsx=None, alias=None, log=print)
```

阶段：

#### S1 定位

- `load_config()`；
- `Session.login()`；
- `locate_report()`；
- 把 `bgbh/rid/pid/igid` 挂到 `sess` 上（写回 URL 需要）。

#### S2 抓取五张表

- `crawl_4126` / `crawl_4127` / `crawl_4123` / `crawl_4125` / `crawl_4122`；
- 用同一个时间戳 `ts` 保存五张表；
- 记录各表统计到 `result['stages']`。

#### S3 校验

- `verify(rows26, recs27, extra=...)`；
- 打印每条 `[PASS/WARN/FAIL]`；
- 有 FAIL 只警告，不阻断后续生成（上传前由人确认）。

#### S4 生成 4122/4123/4125

- 数据源：
  - 默认：本次抓取的 4126 + 4127；
  - `from_xlsx`：从修改后的 Excel 读 4126/4127；
- `build_from_rows()`；
- `write_outputs()` 写三张生成件；
- 记录 `result['generated']`、`result['totals']`。

#### S5 对比 + 上传

- `diff_4122` / `_diff_4123` / `_diff_4125` 分别算差异；
- `need22/need23/need25` 独立判断；
- 只有 `upload=True` 且至少一张有差异才询问；
- `dry=True` → 只打印演练；
- 否则逐张上传**有差异**的表（4122/4123/4125 分开判断，避免“一张有差异就重写三张”）；
- `from_xlsx` 时还会走 S5b：把修改后的 4126/4127 也写回（`diff_rows` → 4126 扫描定位 → 4127 按行序）。

#### S6 复核

- 重新登录、重新定位；
- 重新抓 4122/4123/4125；
- 与生成结果比较；
- `ok = 4122一致 and 4123无差异 and 4125无变化`；
- 写回 `result['recheck']`。

最后把 `result` 写成 `工作流清单_<bgbh>_<ts>.json`。

### 8.2 `_diff_4123` / `_diff_4125`

- `_diff_4123(cur_rows, new_rows)`：
  - 用 `(管段编号, 结构性描述, 功能性描述)` 集合比较；
  - 返回 `(新增, 修改, 删除)`；修改目前计 0（集合法只能看增删）。
- `_diff_4125(cur_defs, cur_total, new_rows, new_total)`：
  - 平台 16 行缺陷名 → `(1级,2级,3级,4级,小计)`；
  - 生成结果逐行比，统计“变化行”；
  - 返回 `(变化行数, 平台合计, 新合计)`。

### 8.3 `batch.run_batch()`：批量

输入：报告编号列表（`read_list()` 支持一行一个、`#` 注释、`编号,工程名`）。

流程：

1. 创建 `批量_<ts>/` 根目录；
2. 逐个报告建子目录；
3. 每个报告：
   - 日志用 `Tee` 同时写屏幕和 `bgbh.log`；
   - `contextlib.redirect_stdout(tee)` 包住 `workflow.run`；
   - 记录 4126/4127/4123/4125 行数、上传表、复核结论、FAIL/WARN 数；
   - 单个失败不影响后续；
4. `_write_summary()`：
   - `批量结果_<ts>.xlsx`：13 列，按状态着色；
   - `批量结果_<ts>.json`：完整结果；
   - 打印成功/失败统计。

`confirm_each=True` 时，每个报告询问 `y=上传 / n=跳过 / a=后面全部上传 / q=退出`。

---

## 9. 图形界面细节

### 9.1 控件与布局

- 顶部 `ttk.Frame`：`Label + Entry + Button`，`grid` 布局；
- 按钮行 `ttk.Frame`：8 个 `ttk.Button`，`pack(side='left')`；
- 上传范围 `ttk.LabelFrame`：5 个 `ttk.Checkbutton` 绑定 5 个 `BooleanVar`；
- 日志区：`tk.Text`（黑底、`Consolas 9`）+ 两个 `ttk.Scrollbar`，`grid` 布局并设置行列权重；
- 状态标签 `self.lbl`。

### 9.2 账号窗口

`login_dialog()`：

- `tk.Toplevel`，`transient` + `grab_set`（模态）；
- 平台地址、账号、密码（`show='*'`）、记住密码复选框；
- 显示配置文件路径；
- `do_ok()`：
  - 账号/密码必填；
  - `credentials.save_file(new, remember=...)`；
  - 更新 `common.OVERRIDES`；
  - 关闭窗口；
- 回车绑定 `do_ok`；
- `wait_window()` 后返回是否成功。

`_ensure_login()`：启动时 `has_login(cfg)` 为假就弹窗；取消则 `after(100, destroy)` 退出。

### 9.3 日志与错误

- 业务代码照常 `print`，`_worker` 用 `redirect_stdout/stderr` 捕获；
- `_poll` 每 120ms 把队列内容写进 `Text`；
- 任何未捕获异常 `traceback.print_exc()`，日志区可见；
- `report_callback_exception` 把 Tk 回调异常写入 `error.log`；
- 启动失败也写 `error.log` 并弹错误框；
- `--smoke`：设置假账号，创建/更新/销毁窗口，打印 `GUI smoke OK`，用于打包后验证。

### 9.4 按钮禁用/恢复

`start()` 禁用 7 个任务按钮（含 `b_upload`），`_finish()` 恢复。  
**注意**：早期版本 `_finish` 漏了 `b_upload`，导致跑完任意任务后“上传勾选”永久变灰；现已修复。

---

## 10. 依赖库与工具清单

### 10.1 运行环境

| 项目 | 值 |
|---|---|
| 开发/打包 Python | Anaconda Python 3.13.5（Windows x64） |
| 运行方式 | 源码：Python 3.10+；发布：PyInstaller onedir，目标机**不需要装 Python** |
| 图形库 | Tkinter / Tcl-Tk 8.6（随 Anaconda 与 `_internal` 一起打包） |
| Excel 库 | openpyxl 3.1.5 |
| 平台 | Windows（DPAPI、Tk、路径均为 Windows 语义） |
| 打包工具 | PyInstaller 6.22.3（仅构建机需要） |

### 10.2 标准库

| 模块 | 用途 | 主要出现位置 |
|---|---|---|
| `argparse` | 命令行参数/子命令 | `cli.py`、`crawl.py`（旧入口） |
| `base64` | 登录账号密码编码、DPAPI 密文编码 | `common.Session.login`、`credentials` |
| `contextlib` | `redirect_stdout/stderr` 捕获日志 | `cctv_gui._worker`、`batch.run_batch` |
| `ctypes` / `ctypes.wintypes` | 调用 Windows DPAPI | `credentials._dpapi` |
| `datetime` | 时间戳、日期区间、报告时间 | 全局 |
| `glob` | 查找最新一套 xlsx | `crawl._latest_file` |
| `hashlib` | 文件 SHA256 | `crawl.sha256` |
| `html` | HTML 实体反转义 | `common._cell_text` |
| `http.cookiejar` | Cookie 容器 | `common.Session` |
| `io` | 文本流（批量 Tee） | `batch` |
| `json` | 登录/定位响应、清单文件 | `common`、`workflow`、`batch` |
| `os` / `sys` | 路径、环境、退出码 | 全局 |
| `queue` | 线程间日志队列 | `cctv_gui` |
| `re` | HTML 解析、编号规范化、正则校验 | `common`、`crawl`、`generate` |
| `threading` | GUI 工作线程 | `cctv_gui` |
| `time` | 重试间隔、翻页限速、保存间隔 | `common`、`crawl`、`writeback` |
| `traceback` | 异常堆栈写入日志 | `cctv_gui`、`batch` |
| `urllib.request` | HTTP 请求、Cookie Opener | `common`、`crawl` |
| `urllib.parse` | URL 编码、query 拼接 | `common`、`crawl` |
| `urllib.error` | 网络异常捕获 | `common.Session._open` |
| `subprocess` | 调用 PyInstaller（仅构建） | `_build/build_exe.py` |

### 10.3 第三方库

| 库 | 版本 | 用途 | 是否随 exe |
|---|---|---|---|
| `openpyxl` | 3.1.5 | 读写 .xlsx、样式、合并单元格、冻结窗格 | 是（打进 `_internal`） |
| `PyInstaller` | 6.22.3 | 把 Python 打成 exe/onedir | 否（仅构建机） |
| `Playwright` | 可选 | `locate_via_browser()` 浏览器兜底 | 否（默认不装、不启用） |
| `markdown` / `Pygments` | 文档生成用 | 把本讲解转成 HTML | 否（仅本次文档生成） |

### 10.4 关键第三方 API

| API | 作用 |
|---|---|
| `openpyxl.Workbook()` | 新建工作簿 |
| `openpyxl.load_workbook(path, read_only=True, data_only=True)` | 只读打开，取计算值 |
| `ws.iter_rows(values_only=True)` | 逐行取值 |
| `ws.cell(row, column, value)` | 写单元格 |
| `ws.merge_cells(...)` | 合并说明行/总计行 |
| `ws.freeze_panes` | 冻结表头 |
| `Font/Alignment/Border/Side/PatternFill` | 字体、对齐、边框、底色 |
| `get_column_letter(i)` | 列号 → A/B/C… |
| `urllib.request.Request` | 构造 GET/POST |
| `urllib.request.build_opener(HTTPCookieProcessor(cj))` | 带 Cookie 的 Opener |
| `urllib.parse.urlencode(..., encoding='gbk')` | 表单 GBK 编码 |
| `urllib.parse.quote(qs, safe='=&')` | 台账 query 编码 |
| `crypt32.CryptProtectData / CryptUnprotectData` | DPAPI 加解密 |
| `kernel32.LocalFree` | 释放 DPAPI 输出内存 |
| `tkinter.Tk / Toplevel / Text / BooleanVar` | 窗口、对话框、日志框、复选框变量 |
| `ttk.Frame/Label/Entry/Button/Checkbutton/LabelFrame/Scrollbar` | 控件 |
| `filedialog.askdirectory / askopenfilename` | 选目录/选文件 |
| `messagebox.showwarning/showinfo/askyesno/showerror` | 提示/确认/错误框 |

### 10.5 外部依赖

| 依赖 | 说明 |
|---|---|
| 平台 HTTP 服务 | `http://36.134.56.138:7500`，JSP + Cookie 会话 |
| 平台账号 | 需要项目权限；账号密码保存在本机配置 |
| Windows DPAPI | 同一 Windows 用户才能解密密码 |
| 文件系统 | 输出目录需可写（默认 exe 同目录 `输出`） |

---

## 11. 全部函数索引

> 下表按模块列出项目中定义的全部函数/方法（不含 `_vendor`、历史项目脚本）。  
> “关键点”一列概括实现，详细原理见前文对应章节。  
> 标准库与第三方库的调用见 10.2–10.4；本表只列项目自定义函数。

### 11.1 `cctv.py`

| 函数 | 作用 |
|---|---|
| （无自定义函数） | 仅 `from cctvtools.cli import main`，`__main__` 调 `main()` |

### 11.2 `cctv_gui.py`

| 函数/方法 | 作用 | 关键点 |
|---|---|---|
| `base_dir()` | 取 exe/脚本所在目录 | 冻结判断 `sys.frozen` |
| `default_out_dir()` | 默认输出目录 | `base_dir()/输出`，自动创建 |
| `QueueWriter.__init__(q)` | 保存队列 | — |
| `QueueWriter.write(s)` | 标准输出 → 队列 | 空串忽略 |
| `QueueWriter.flush()` | 空实现 | 兼容 print |
| `App.__init__()` | 初始化窗口 | 队列、running、登录检查、`after` 轮询 |
| `App._build()` | 创建全部控件 | 表单/按钮/复选框/日志区 |
| `App.pick_dir()` | 选择输出目录 | `filedialog.askdirectory` |
| `App.bgbh()` | 读取/校验报告编号 | 正则 `\d{2}KHD\d{4,6}`，非标准可确认 |
| `App.out_dir()` | 读取/创建输出目录 | 空则默认目录 |
| `App.log(s)` | 追加日志并滚到底 | `Text.insert/see` |
| `App._poll()` | 轮询队列 | 每 120ms，处理 log/done |
| `App._finish(ok)` | 任务结束 | 恢复按钮（含 `b_upload`）、更新标签 |
| `App.start(fn, title)` | 启动任务 | 禁用按钮、清日志、起线程 |
| `App._worker(fn)` | 工作线程 | 重定向 stdout/stderr、捕获异常、计时 |
| `App._update_title()` | 标题显示账号 | `C.load_config` |
| `App._ensure_login()` | 启动登录检查 | 未登录弹窗 |
| `App.login_dialog()` | 账号模态窗 | 地址/账号/密码/记住密码，保存 DPAPI |
| `App.change_account()` | 修改账号 | 登录窗 + 更新标题 |
| `App._wf(upload)` | 工作流包装 | `workflow.run` |
| `App.do_workflow()` | 一键全流程 | 不上传 |
| `App.do_workflow_upload()` | 全流程+上传 | 二次确认后上传 |
| `App.do_crawl()` | 只抓五张表 | 登录、抓取、保存 |
| `App.do_generate()` | 生成 4122/4123/4125 | 读最新 4126/4127，`G.write_outputs` |
| `App.do_verify()` | 校验报告 | `verify_report.build` |
| `App.do_upload_tables()` | 上传勾选 | 收集 5 个复选框 → `upload_selected` |
| `App.do_batch()` | 批量 | 选 txt → `batch.run_batch` |
| `_write_error(text)` | 写 `error.log` | 冻结时写 exe 目录 |
| `main()` | 程序入口 | `--smoke`、异常兜底、`mainloop` |

### 11.3 `cctvtools/cli.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `make_session(bgbh)` | 登录 + 定位 + 挂 rid/pid | 供 diff/upload 命令复用 |
| `cmd_workflow(args)` | `workflow` 子命令 | 组装 alias、upload/dry/yes |
| `_interactive_login()` | 交互式登录 | 无配置时 `input` 并保存 |
| `interactive()` | 无参数菜单 | 1–5/0 |
| `cmd_login(args)` | `login` 子命令 | 可 `--test` 验证、`--no-remember` |
| `cmd_verify_report(args)` | 校验报告 | `VR.build`，FAIL 返回 2 |
| `cmd_upload_tables(args)` | 选择性上传 | `UP.upload_selected` |
| `cmd_batch(args)` | 批量 | 合并列表/文件，解析 alias |
| `cmd_locate(args)` | 定位 | 打印 JSON |
| `cmd_crawl(args)` | 抓表 | 单表/多表/全部 |
| `cmd_all(args)` | 旧版全流程 | 构造 Namespace 调 `C.run_all` |
| `cmd_verify(args)` | 离线校验 | `C.verify_latest` |
| `cmd_generate(args)` | 生成三表 | `D.read_workbook` + `G.write_outputs` |
| `cmd_diff(args)` | 变更清单 | 抓平台 4126/4127 + `D.compare` |
| `cmd_upload(args)` | 按 Excel 写回 4126/4127 | `WB.diff_rows`/`scan`/`upload` |
| `build_parser()` | 构建 argparse | 全局参数 + 12 个子命令 |
| `_normalize(argv)` | 全局参数归位 | 允许写在子命令后 |
| `main(argv)` | CLI 主流程 | OVERRIDES、show-config、交互/分发 |

### 11.4 `cctvtools/common.py`

| 函数/方法 | 作用 | 关键点 |
|---|---|---|
| `load_config()` | 合并配置 | DEFAULTS → config.py → 文件/环境/命令行 |
| `print_config(cfg, meta, bgbh)` | 打印生效配置 | 不打印密码 |
| `Session.__init__(cfg, retries, timeout, verbose)` | 初始化会话 | CookieJar + Opener + UA |
| `Session.login()` | 登录 | base64 账号密码，POST `/loginAction`，G1 |
| `Session._open(req)` | 带重试请求 | 3 次，退避 1+i 秒 |
| `Session.get_raw(url, referer)` | GET bytes | Referer、XHR 头 |
| `Session.get(url, encoding='gbk', referer)` | GET 文本 | 默认 GBK |
| `locate_report(sess, bgbh, igid, days, verbose)` | 定位报告 | 台账接口、GBK/UTF-8、G2 |
| `locate_via_browser(bgbh, headless)` | 浏览器兜底 | 延迟导入 Playwright 脚本 |
| `page_url(sess, biaodan, report, page)` | 数据页 URL | `currentpage` 0 基 |
| `_data_table(html_text)` | 取 data-table | 正则 |
| `_cell_text(td)` | 取单元格文本 | 去标签、实体、nbsp |
| `same_value(a, b)` | 空值等价比较 | `''/'/'/'--'` 等价 |
| `parse_form(html_text)` | 解析表单控件 | input/select/textarea，name/value |
| `put(form, name, val)` | 按名写表单 | 返回是否命中 |
| `_report_url(self)` | 报告首页 URL | `reportShow.jsp?BiaoDanId=412` |
| `_view_url(self, bid, pg)` | 查看页 URL | `XiangMuXianShi.jsp` |
| `_anniu_url(self, bid)` | 按钮区 URL | `XiangMuAnNiu.jsp` |
| `_edit_url(self, bid, pg)` | 编辑页 URL | `XiangMuXiuGai.jsp` |
| `_post(self, url, data, ref, enc)` | POST 表单 | GBK urlencode、Referer、解码 |
| `_init_write(self)` | 空初始化 | 兼容旧接口 |

### 11.5 `cctvtools/credentials.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `_app_dir()` | `%APPDATA%\CCTV工具` | — |
| `_candidates()` | 候选配置路径 | CCTV_CONFIG → APPDATA → exe 目录 → `~/.cctv` |
| `_writable(path)` | 判断可写 | `open(path,'a')` 会建空文件 |
| `config_path(writable=False)` | 选实际/可写路径 | 存在优先/可写优先 |
| `_BLOB` | DATA_BLOB 结构 | ctypes.Structure |
| `_make_blob(data)` | 构造 DATA_BLOB | 保留 buffer 防 GC |
| `_dpapi(data, protect)` | DPAPI 加/解密 | crypt32 + LocalFree |
| `encrypt(text)` | 加密 + base64 | — |
| `decrypt(token)` | base64 + 解密 | — |
| `load_file()` | 读配置 | 兼容 pwd_enc / pwd |
| `save_file(cfg, remember=True)` | 写配置 | DPAPI 失败退化明文 |
| `clear_password()` | 清密码 | 保留账号 |
| `resolve(cli)` | 合并配置来源 | 文件 → 环境 → 命令行 |
| `has_login(cfg)` | 是否已登录 | USER 与 PWD 均非空 |

### 11.6 `cctvtools/crawl.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `parse_4126(html_text)` | 解析 4126 一页 | 21 列、13 行、过滤表头/占位/异常 |
| `crawl_4126(...)` | 翻页抓 4126 | 去重、empty/repeat/cap |
| `parse_4127(html_text)` | 解析 4127 一页 | 标签取值 + 4 缺陷槽位 |
| `cells0(rows)` | 卡片序号 | 首行首列 |
| `crawl_4127(...)` | 翻页抓 4127 | 按起止井号去重 |
| `_num(v)` | 数字串转 int/float | 其余原样 |
| `_num_in(text, pattern)` | 从文字抠数字 | 正则 |
| `parse_4122(html_text)` | 解析 4122 | 说明/表头/明细/总计 |
| `crawl_4122(...)` | 抓 4122 | 单页有效，空壳停 |
| `_segments_in_note(rows)` | 备注“分N段”求和 | 用于 G9 |
| `parse_4123(html_text)` | 解析 4123 | 6 列、12 行/页 |
| `crawl_4123(...)` | 抓 4123 | 按管段去重 |
| `parse_4125(html_text)` | 解析 4125 | 16 行 + 合计，7/8 列兼容 |
| `crawl_4125(...)` | 抓 4125 | 16 行或空停 |
| `_write_sheet(ws, headers, widths, rows, freeze)` | 通用表头/表体 | 样式、列宽、冻结 |
| `output_name(table, bgbh, ts)` | 输出文件名 | `(网页爬取)_时间` |
| `save_4126(rows, path)` | 存 4126 | 21 列 |
| `save_4127(records, path)` | 存 4127 | 16+4×5+1 列 |
| `sha256(path)` | 文件哈希 | 1MB 分块 |
| `_summary_block(ws, ncols, text)` | 写说明行 | 合并 A1:ncols1 |
| `_table_body(ws, headers, widths, body, hr)` | 写表头+表体 | 返回下一行 |
| `_write_total(ws, row, values, merge_to)` | 写合计行 | 可合并 A:B |
| `save_4122(summary, rows, total, path)` | 存 4122 | 三段式 + 总计 |
| `save_4123(summary, rows, path)` | 存 4123 | 三段式 |
| `save_4125(summary, defects, total, path)` | 存 4125 | 三段式 + 合计 |
| `_seg_key(seg)` | 编号规范化 | 去所有空白 |
| `_is_real_defect(code)` | 是否真实缺陷 | 排除 `(ZC)正常`/`--`/空 |
| `_code_type(code)` | 代码分类 | S/F/? |
| `_defect_pipes(recs_4127)` | 有缺陷的段 | 任一真实缺陷 |
| `_defect_entries(recs_4127)` | 缺陷条目数 | 所有真实槽位 |
| `verify(rows_4126, recs_4127, extra)` | 交叉校验 | G3–G9、方向翻转、重复 |
| `write_manifest(...)` | 写 manifest JSON | 版本、阶段、校验、SHA256 |
| `run_all(cfg, args)` | 旧版全流程 | 抓五表 → 校验 → manifest |
| `_crawl_and_save(...)` | 单表抓取+落盘 | 0 行默认不落盘 |
| `resolve_report(cfg, args)` | 定位报告 | 显式参数 > HTTP > 浏览器 |
| `main(argv)` | crawl 模块旧入口 | argparse + 子命令 |
| `_latest_file(table, bgbh, out_dir)` | 最新抓取件 | glob + 排序 |
| `_read_grid(path)` | 读首个工作表 | 只读、data_only |
| `_read_223_output(table, path)` | 回读三段式产物 | 说明/表头/表体/总计 |
| `load_outputs(bgbh, out_dir)` | 读最新一套五表 | 返回 rows/recs/extra/files |
| `verify_latest(bgbh, out_dir)` | 离线校验 | 0/2 退出码 |
| `_fake_4126(pages)` | 自检假 4126 | HTML 构造 |
| `_fake_4122()` | 自检假 4122 | — |
| `_fake_4123()` | 自检假 4123 | — |
| `_fake_4125()` | 自检假 4125 | — |
| `_fake_4127(k, codes, descs)` | 自检假 4127 | 标签行 + 3 值行 |
| `selftest()` | 离线自检 | 解析器/校验器全跑 |

### 11.7 `cctvtools/generate.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `_real(c)` | 是否真实缺陷文本 | 非空、非 `/`、不含 `(ZC)` |
| `_code(c)` | 从 `(BX)变形` 取 `BX` | 正则 |
| `_seg(v)` | 编号去空白 | `W-15-W-21` |
| `_to_float(v)` | 文本转数字 | 去 `m/米`，失败 0 |
| `_fmt(v)` | 数字格式化 | 1 位小数、去 `.0` |
| `_pipe_name(t)` | 管段类型 → 名称 | `污水（WS）` → `污水管道` |
| `build_4122(rows4126, rows4127)` | 生成 4122 | 分组/长度/排序/格式化 |
| `_num1(v)` | 比较用数值 | 1 位小数 |
| `_row22_key(r)` | 4122 行键 | `(名称, 长度, 检测, 备注)` |
| `diff_4122(...)` | 比较平台 vs 生成 | 多重集合、忽略顺序 |
| `build_from_rows(...)` | 生成三表总入口 | 4122/4123/4125 + defects |
| `write_4123(data, path)` | 写 4123 xlsx | 样式 |
| `write_4125(data, path)` | 写 4125 xlsx | 样式 |
| `write_4122(data, path)` | 写 4122 xlsx | 复用 `crawl.save_4122` |
| `write_outputs(data, bgbh, outdir, ts)` | 写三张生成件 | 4122 带 `(生成)` 后缀 |

### 11.8 `cctvtools/diff.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `read_workbook(path)` | 读 4126/4127 | 按表头识别 |
| `compare(plat26, plat27, rows26, rows27, alias)` | 逐字段比较 | `_COL26/_COL27/_SLOT` |
| `write_list(changes, path, bgbh)` | 写变更清单 | 明细 + 汇总两 sheet |

### 11.9 `cctvtools/writeback.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `open_edit(sess, bid, pg)` | 进入编辑态取表单页 | 锁定 + JSON url + GBK |
| `save_page(sess, bid, pg, form)` | 保存整页 | POST `XiangMuXiuGaiBaoCun2` |
| `apply_page(sess, bid, pg, row_fields, dry, expect)` | 填一页并保存 | 防错位、same_value、缺字段报错 |
| `diff_rows(...)` | 生成 4126/4127 计划 | 字段映射、alias |
| `scan_4126_positions(sess, pages)` | 段 → (页,行) | 逐页解析 `c_601_N` |
| `upload(sess, bid, plan, positions, dry, progress)` | 写 4126/4127 | 按页分组 |
| `upload_4123(...)` | 写 4123 | 12 行/页 |
| `upload_4125(...)` | 写 4125 | 16 行 + 合计 |
| `upload_4122(...)` | 写 4122 | 10 行/页 + 说明/总计 |

### 11.10 `cctvtools/workflow.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `_ask(prompt)` | 命令行询问 y/N | EOF 视为否 |
| `_same_row(a, b, n)` | 整行比较 | `same_value` |
| `_diff_4123(cur_rows, new_rows)` | 4123 差异 | 集合增删 |
| `_diff_4125(...)` | 4125 差异 | 逐缺陷行比较 |
| `run(...)` | 一键全流程 | S1–S6 + manifest |

### 11.11 `cctvtools/uploader.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `_read_rows(path, ncols)` | 读 xlsx 为字符串行 | 跳过表头/空行 |
| `_cmp_4123(cur_rows, new_rows)` | 4123 集合比较 | 返回增/删 |
| `upload_selected(...)` | 按勾选上传 | 五表独立跳过/写回 |

### 11.12 `cctvtools/verify_report.py`

| 函数 | 作用 | 关键点 |
|---|---|---|
| `_level_fill(level)` | 级别 → 填充色 | PASS/WARN/FAIL |
| `_file_info(path)` | 文件大小/时间 | KB、`YYYY-MM-DD HH:MM` |
| `build(bgbh, out_dir)` | 生成校验报告 | 4 个 sheet |

### 11.13 `cctvtools/batch.py`

| 函数/方法 | 作用 | 关键点 |
|---|---|---|
| `Tee.__init__(*streams)` | 多流输出 | 屏幕 + 日志文件 |
| `Tee.write(s)` | 写所有流 | 单流异常不影响其它 |
| `Tee.flush()` | 刷新所有流 | — |
| `read_list(path)` | 读报告编号列表 | 忽略空行/`#`，支持 `编号,工程名` |
| `_fail_warn(checks)` | 统计 FAIL/WARN | — |
| `prompt_upload(bgbh)` | 逐报告确认 | y/n/a/q |
| `run_batch(...)` | 批量执行 | 单报告失败不中断 |
| `_write_summary(...)` | 批量结果 xlsx/json | 状态着色、统计 |

---

### 11.14 函数签名速查（按源码 AST 自动生成）

> 本表覆盖 `cctv.py`、`cctv_gui.py`、`cctvtools/*.py` 中定义的全部函数与方法，便于按签名检索；实现细节见上文各模块章节。

| 模块 | 函数 / 方法签名 | 源码行 |
|---|---|---|
| `cctv_gui.py` | `base_dir()` | 31 |
| `cctv_gui.py` | `default_out_dir()` | 38 |
| `cctv_gui.py` | `QueueWriter.__init__(self, q)` | 45 |
| `cctv_gui.py` | `QueueWriter.write(self, s)` | 48 |
| `cctv_gui.py` | `QueueWriter.flush(self)` | 52 |
| `cctv_gui.py` | `App.__init__(self)` | 57 |
| `cctv_gui.py` | `App._build(self)` | 73 |
| `cctv_gui.py` | `App.pick_dir(self)` | 132 |
| `cctv_gui.py` | `App.bgbh(self)` | 138 |
| `cctv_gui.py` | `App.out_dir(self)` | 148 |
| `cctv_gui.py` | `App.log(self, s)` | 153 |
| `cctv_gui.py` | `App._poll(self)` | 157 |
| `cctv_gui.py` | `App._finish(self, ok)` | 169 |
| `cctv_gui.py` | `App.start(self, fn, title)` | 176 |
| `cctv_gui.py` | `App._worker(self, fn)` | 188 |
| `cctv_gui.py` | `App._update_title(self)` | 203 |
| `cctv_gui.py` | `App._ensure_login(self)` | 209 |
| `cctv_gui.py` | `App.login_dialog(self)` | 215 |
| `cctv_gui.py` | `App.change_account(self)` | 264 |
| `cctv_gui.py` | `App._wf(self, upload=False)` | 270 |
| `cctv_gui.py` | `App.do_workflow(self)` | 279 |
| `cctv_gui.py` | `App.do_workflow_upload(self)` | 282 |
| `cctv_gui.py` | `App.do_crawl(self)` | 289 |
| `cctv_gui.py` | `App.do_generate(self)` | 315 |
| `cctv_gui.py` | `App.do_verify(self)` | 348 |
| `cctv_gui.py` | `App.do_upload_tables(self)` | 365 |
| `cctv_gui.py` | `App.do_batch(self)` | 382 |
| `cctv_gui.py` | `_write_error(text)` | 399 |
| `cctv_gui.py` | `main()` | 414 |
| `cctvtools.batch` | `Tee.__init__(self, *streams)` | 31 |
| `cctvtools.batch` | `Tee.write(self, s)` | 34 |
| `cctvtools.batch` | `Tee.flush(self)` | 42 |
| `cctvtools.batch` | `read_list(path)` | 50 |
| `cctvtools.batch` | `_fail_warn(checks)` | 63 |
| `cctvtools.batch` | `prompt_upload(bgbh)` | 69 |
| `cctvtools.batch` | `run_batch(bgbh_list, out_dir='.', upload=False, dry=False, yes=False, aliases=None, from_xlsx=None, confirm_each=False, log=print)` | 81 |
| `cctvtools.batch` | `_write_summary(results, root, ts, dry, upload, log)` | 152 |
| `cctvtools.cli` | `make_session(bgbh)` | 32 |
| `cctvtools.cli` | `cmd_workflow(args)` | 42 |
| `cctvtools.cli` | `_interactive_login()` | 49 |
| `cctvtools.cli` | `interactive()` | 69 |
| `cctvtools.cli` | `cmd_login(args)` | 119 |
| `cctvtools.cli` | `cmd_verify_report(args)` | 131 |
| `cctvtools.cli` | `cmd_upload_tables(args)` | 140 |
| `cctvtools.cli` | `cmd_batch(args)` | 148 |
| `cctvtools.cli` | `cmd_locate(args)` | 168 |
| `cctvtools.cli` | `cmd_crawl(args)` | 174 |
| `cctvtools.cli` | `cmd_all(args)` | 206 |
| `cctvtools.cli` | `cmd_verify(args)` | 218 |
| `cctvtools.cli` | `cmd_generate(args)` | 222 |
| `cctvtools.cli` | `cmd_diff(args)` | 233 |
| `cctvtools.cli` | `cmd_upload(args)` | 264 |
| `cctvtools.cli` | `build_parser()` | 312 |
| `cctvtools.cli` | `_normalize(argv)` | 404 |
| `cctvtools.cli` | `main(argv=None)` | 418 |
| `cctvtools.common` | `load_config()` | 103 |
| `cctvtools.common` | `print_config(cfg, meta, bgbh=None)` | 148 |
| `cctvtools.common` | `Session.__init__(self, cfg, retries=3, timeout=30, verbose=True)` | 166 |
| `cctvtools.common` | `Session.login(self)` | 179 |
| `cctvtools.common` | `Session._open(self, req)` | 202 |
| `cctvtools.common` | `Session.get_raw(self, url, referer=None)` | 213 |
| `cctvtools.common` | `Session.get(self, url, encoding='gbk', referer=None)` | 219 |
| `cctvtools.common` | `locate_report(sess, bgbh, igid, days=3650, verbose=True)` | 224 |
| `cctvtools.common` | `locate_via_browser(bgbh, headless=True)` | 290 |
| `cctvtools.common` | `page_url(sess, biaodan, report, page)` | 302 |
| `cctvtools.common` | `_data_table(html_text)` | 309 |
| `cctvtools.common` | `_cell_text(td)` | 314 |
| `cctvtools.common` | `same_value(a, b)` | 326 |
| `cctvtools.common` | `parse_form(html_text)` | 334 |
| `cctvtools.common` | `put(form, name, val)` | 348 |
| `cctvtools.common` | `_report_url(self)` | 358 |
| `cctvtools.common` | `_view_url(self, bid, pg)` | 362 |
| `cctvtools.common` | `_anniu_url(self, bid)` | 368 |
| `cctvtools.common` | `_edit_url(self, bid, pg)` | 373 |
| `cctvtools.common` | `_post(self, url, data, ref=None, enc='utf-8')` | 378 |
| `cctvtools.common` | `_init_write(self, *a, **k)` | 386 |
| `cctvtools.crawl` | `parse_4126(html_text)` | 58 |
| `cctvtools.crawl` | `crawl_4126(sess, report, max_pages=80, sleep=0.0, verbose=True)` | 89 |
| `cctvtools.crawl` | `parse_4127(html_text)` | 136 |
| `cctvtools.crawl` | `cells0(rows)` | 211 |
| `cctvtools.crawl` | `crawl_4127(sess, report, max_pages=800, sleep=0.0, verbose=True)` | 216 |
| `cctvtools.crawl` | `_num(v)` | 243 |
| `cctvtools.crawl` | `_num_in(text, pattern)` | 253 |
| `cctvtools.crawl` | `parse_4122(html_text)` | 260 |
| `cctvtools.crawl` | `crawl_4122(sess, report, max_pages=10, verbose=True)` | 288 |
| `cctvtools.crawl` | `_segments_in_note(rows)` | 312 |
| `cctvtools.crawl` | `parse_4123(html_text)` | 323 |
| `cctvtools.crawl` | `crawl_4123(sess, report, max_pages=60, verbose=True)` | 341 |
| `cctvtools.crawl` | `parse_4125(html_text)` | 373 |
| `cctvtools.crawl` | `crawl_4125(sess, report, verbose=True)` | 413 |
| `cctvtools.crawl` | `_write_sheet(ws, headers, widths, rows, freeze)` | 437 |
| `cctvtools.crawl` | `output_name(table, bgbh, ts)` | 454 |
| `cctvtools.crawl` | `save_4126(rows, path)` | 458 |
| `cctvtools.crawl` | `save_4127(records, path)` | 467 |
| `cctvtools.crawl` | `sha256(path)` | 486 |
| `cctvtools.crawl` | `_summary_block(ws, ncols, text)` | 495 |
| `cctvtools.crawl` | `_table_body(ws, headers, widths, body, hr)` | 506 |
| `cctvtools.crawl` | `_write_total(ws, row, values, merge_to=None)` | 524 |
| `cctvtools.crawl` | `save_4122(summary, rows, total, path)` | 536 |
| `cctvtools.crawl` | `save_4123(summary, rows, path)` | 549 |
| `cctvtools.crawl` | `save_4125(summary, defects, total, path)` | 559 |
| `cctvtools.crawl` | `_seg_key(seg)` | 575 |
| `cctvtools.crawl` | `_is_real_defect(code)` | 579 |
| `cctvtools.crawl` | `_code_type(code)` | 583 |
| `cctvtools.crawl` | `_defect_pipes(recs_4127)` | 594 |
| `cctvtools.crawl` | `_defect_entries(recs_4127)` | 599 |
| `cctvtools.crawl` | `verify(rows_4126, recs_4127, extra=None)` | 603 |
| `cctvtools.crawl` | `write_manifest(out_dir, report, stages, checks, outputs, params)` | 780 |
| `cctvtools.crawl` | `run_all(cfg, args)` | 800 |
| `cctvtools.crawl` | `_crawl_and_save(sess, report, table, out_dir, ts, args)` | 882 |
| `cctvtools.crawl` | `resolve_report(cfg, args)` | 911 |
| `cctvtools.crawl` | `main(argv=None)` | 928 |
| `cctvtools.crawl` | `_latest_file(table, bgbh, out_dir='.')` | 992 |
| `cctvtools.crawl` | `_read_grid(path)` | 999 |
| `cctvtools.crawl` | `_read_223_output(table, path)` | 1006 |
| `cctvtools.crawl` | `load_outputs(bgbh, out_dir='.')` | 1035 |
| `cctvtools.crawl` | `verify_latest(bgbh, out_dir='.')` | 1057 |
| `cctvtools.crawl` | `_fake_4126(pages)` | 1073 |
| `cctvtools.crawl` | `_fake_4122()` | 1089 |
| `cctvtools.crawl` | `_fake_4123()` | 1099 |
| `cctvtools.crawl` | `_fake_4125()` | 1109 |
| `cctvtools.crawl` | `_fake_4127(k, codes=None, descs=None)` | 1121 |
| `cctvtools.crawl` | `selftest()` | 1148 |
| `cctvtools.credentials` | `_app_dir()` | 22 |
| `cctvtools.credentials` | `_candidates()` | 27 |
| `cctvtools.credentials` | `_writable(path)` | 42 |
| `cctvtools.credentials` | `config_path(writable=False)` | 53 |
| `cctvtools.credentials` | `_make_blob(data: bytes)` | 72 |
| `cctvtools.credentials` | `_dpapi(data: bytes, protect: bool)` | 77 |
| `cctvtools.credentials` | `encrypt(text: str)` | 92 |
| `cctvtools.credentials` | `decrypt(token: str)` | 96 |
| `cctvtools.credentials` | `load_file()` | 101 |
| `cctvtools.credentials` | `save_file(cfg, remember=True)` | 133 |
| `cctvtools.credentials` | `clear_password()` | 149 |
| `cctvtools.credentials` | `resolve(cli=None)` | 165 |
| `cctvtools.credentials` | `has_login(cfg)` | 185 |
| `cctvtools.diff` | `read_workbook(path)` | 40 |
| `cctvtools.diff` | `compare(plat26, plat27, rows26, rows27, alias=None)` | 61 |
| `cctvtools.diff` | `write_list(changes, path, bgbh='')` | 118 |
| `cctvtools.generate` | `_real(c)` | 26 |
| `cctvtools.generate` | `_code(c)` | 30 |
| `cctvtools.generate` | `_seg(v)` | 35 |
| `cctvtools.generate` | `_to_float(v)` | 40 |
| `cctvtools.generate` | `_fmt(v)` | 50 |
| `cctvtools.generate` | `_pipe_name(t)` | 56 |
| `cctvtools.generate` | `build_4122(rows4126, rows4127)` | 66 |
| `cctvtools.generate` | `_num1(v)` | 115 |
| `cctvtools.generate` | `_row22_key(r)` | 122 |
| `cctvtools.generate` | `diff_4122(cur_rows, cur_total, cur_summary, data)` | 127 |
| `cctvtools.generate` | `build_from_rows(rows4126, rows4127)` | 142 |
| `cctvtools.generate` | `write_4123(data, path)` | 204 |
| `cctvtools.generate` | `write_4125(data, path)` | 230 |
| `cctvtools.generate` | `write_4122(data, path)` | 264 |
| `cctvtools.generate` | `write_outputs(data, bgbh, outdir='.', ts=None)` | 270 |
| `cctvtools.uploader` | `_read_rows(path, ncols)` | 14 |
| `cctvtools.uploader` | `_cmp_4123(cur_rows, new_rows)` | 19 |
| `cctvtools.uploader` | `upload_selected(bgbh, out_dir, tables, dry=False, alias=None, log=print)` | 25 |
| `cctvtools.verify_report` | `_level_fill(level)` | 51 |
| `cctvtools.verify_report` | `_file_info(path)` | 55 |
| `cctvtools.verify_report` | `build(bgbh, out_dir='.')` | 63 |
| `cctvtools.workflow` | `_ask(prompt)` | 24 |
| `cctvtools.workflow` | `_same_row(a, b, n)` | 31 |
| `cctvtools.workflow` | `_diff_4123(cur_rows, new_rows)` | 36 |
| `cctvtools.workflow` | `_diff_4125(cur_defs, cur_total, new_rows, new_total)` | 43 |
| `cctvtools.workflow` | `run(bgbh, out_dir='.', upload=True, dry=False, yes=False, from_xlsx=None, alias=None, log=print)` | 55 |
| `cctvtools.writeback` | `open_edit(sess, bid, pg)` | 24 |
| `cctvtools.writeback` | `save_page(sess, bid, pg, form)` | 37 |
| `cctvtools.writeback` | `apply_page(sess, bid, pg, row_fields, dry=False, expect=None)` | 43 |
| `cctvtools.writeback` | `diff_rows(rows_4126, recs_4127, plat_4126, plat_4127, alias=None)` | 81 |
| `cctvtools.writeback` | `scan_4126_positions(sess, pages=13)` | 126 |
| `cctvtools.writeback` | `upload(sess, bid, plan, positions, page_size=None, dry=False, progress=None)` | 140 |
| `cctvtools.writeback` | `upload_4123(sess, summary, rows, dry=False, page_size=12)` | 159 |
| `cctvtools.writeback` | `upload_4125(sess, summary, rows16, total, dry=False)` | 183 |
| `cctvtools.writeback` | `upload_4122(sess, summary, rows, total, dry=False, page_size=10)` | 203 |

## 12. 关键设计取舍与边界

### 12.1 为什么用正则解析 HTML，而不是 BeautifulSoup/lxml

- 平台是固定模板的 JSP，`data-table` 标记稳定；
- 只依赖标准库 + openpyxl，打包体积小；
- 代价：页面改版会影响解析。因此每个 parser 都带 `meta`/停页原因，`run_all` 会把 `malformed_rows` 作为 WARN 写进清单，便于发现结构漂移。

### 12.2 为什么 4127 不能用“标签右邻值”简单扫

4127 的缺陷块是纵向 4 槽位，`rowspan` 导致第 2–4 个槽位的标签不在同一行。只扫标签右邻值会漏掉第 2 条及以后的缺陷。`parse_4127` 按“标签行 + 后续等宽值行”成块读取，保留 4 个位置。

### 12.3 为什么管段匹配不能 `split('-')`

井号本身含 `-`（`W66-1`、`W10-1-W10-2`），拆开后无法还原。工具改为用 4127 的 `(起, 止)` 正向/反向拼候选串，再与 4126 的单列编号比对，天然免疫多横杠。

### 12.4 为什么写回要“锁定/进入编辑态”

平台是老的 JSP 表单：必须先 `XiangMuCaoZuo` 进入编辑态，编辑页才可提交；直接 POST 保存接口可能被服务端拒绝。`open_edit()` 封装了这个流程，并在每页保存后 sleep 0.3s。

### 12.5 为什么要有 G1/G2 闸门

- G1：登录必须 `success=true`，否则后续所有页面都会退化成登录页，抓到的“空数据”会被误当成真实结果；
- G2：`report_id` 必须是 UUID、`parentid` 必须是数字，否则翻页 100% 拿到空壳。

这两个闸门是“快速失败”，避免把错误数据写进 Excel/平台。

### 12.6 已知边界

| 边界 | 说明 |
|---|---|
| 4127 缺陷最多 4 条 | 超过 4 条会被截断，校验 G6 给 WARN；需要扩槽位或另存 |
| 4126 每页 13 行 | 平台固定；若改版需调整 `ROWS_PER_PAGE_4126` |
| 4123 每页 12 行 | `upload_4123` 按 12 行分页 |
| 4122 每页 10 行 | `upload_4122` 按 10 行分页；超过 10 行需要平台存在第 2 页 |
| 4126/4127 不能自动新增段 | `diff_rows` 只处理平台已有段；文件里新增的段不会写回 |
| 4122 生成件不覆盖抓取件 | 生成件用 `(生成)` 后缀 |
| 保存成功判断 | 以响应包含 `XiangMuXianShi.jsp` 为准；平台若改返回格式需调整 |
| 配置文件 0 字节 | `_writable()` 探测会建空文件；分发时不要带空配置 |
| DPAPI | 密码绑定 Windows 用户，换机/换用户需重填 |
| 平台账号权限 | 定位/抓取/写回都受项目权限限制 |

### 12.7 4122 与 4125 的区别（容易混）

- **4122 主要工程量表**：按“管段类型 + 材质 + 管径”汇总**长度/检测长度**，行数等于分组数；
- **4125 管道缺陷数量统计表**：按 16 类缺陷统计**1–4 级数量**，固定 16 行 + 合计；
- 4122 的“工程量”是管道长度，不是缺陷数量。

---

## 13. 排错与运维

### 13.1 自检

```bash
# 源码
python cctv.py selftest
# 打包版
CCTV命令行.exe selftest
```

自检不联网，覆盖：

- 4126 解析（13 行、表头、占位行、越界页）；
- 4127 解析（起始井号、图片、多缺陷第 2 槽位、越界页）；
- 4122/4123/4125 解析；
- 校验器（完全一致、方向翻转、缺段、重复等）。

### 13.2 只读验证

```bash
# 五张表 dry-run（不写平台）
CCTV命令行.exe upload-tables --bgbh 26KHD00029 --out-dir 输出 --tables 4122,4126,4127,4123,4125 --dry
```

日志会显示每张表“已上传/跳过/演练”，不会提交保存。

### 13.3 常见问题

| 现象 | 原因 | 处理 |
|---|---|---|
| 双击没反应/被拦截 | 360/Defender 拦截未签名 exe | 允许运行/加信任 |
| 未配置平台账号 | 配置文件缺失/密码解密失败 | 启动弹窗填写，或 `login` 命令 |
| 定位失败 count=0 | 日期区间不够/iGid 不对/无权限 | 调大 `--days`、核对账号权限 |
| 抓取 0 行 | 登录失效/页面改版/越界 | 先 `selftest`，再 `--show-config` |
| 进入编辑态失败 | 锁定失败/并发编辑/权限 | 关闭网页端编辑，重试 |
| 保存失败 | 响应不含跳转标记/字段缺失 | 看日志中“缺少字段”，核对平台页面 |
| 上传按钮变灰 | 旧版 `_finish` 漏恢复 | 已修复；重启程序也可恢复 |
| 中文乱码（CLI） | 控制台编码 | `chcp 65001` |
| Excel 打不开 | 文件被占用 | 关闭 Excel 再跑 |
| 4127 缺陷少了 | 单管超过 4 条 | 看 G6 WARN，扩槽位/另存 |

### 13.4 日志文件

- GUI 未捕获异常 → exe 同目录 `error.log`；
- 批量 → 每个报告子目录 `<bgbh>.log`；
- 工作流 → `工作流清单_<bgbh>_<ts>.json`；
- 旧版 `all` → `manifest_<bgbh>_<ts>.json`（含文件 SHA256）。

---

## 14. 附录

### 14.1 URL 清单

| 用途 | 路径 |
|---|---|
| 登录 | `POST /loginAction` |
| 台账定位 | `GET /XiangMutaizhangJson?ParaNames=...` |
| 报告首页 | `GET /reportShow.jsp?id={rid}&BiaoDanId=412&iGid={igid}` |
| 数据查看 | `GET /system/XiangMuXianShi.jsp?BiaoDanId=...&currentpage=N` |
| 按钮区 | `GET /system/XiangMuAnNiu.jsp?BiaoDanId=...` |
| 进入编辑 | `POST /XiangMuCaoZuo` |
| 编辑页 | `GET /system/XiangMuXiuGai.jsp?...&currentpage=N` |
| 保存 | `POST /XiangMuXiuGaiBaoCun2` |

### 14.2 文件名规则

| 类型 | 规则 |
|---|---|
| 抓取件 | `{表名}-{bgbh}(网页爬取)_{YYYYMMDD_HHMM}.xlsx` |
| 生成件 4123/4125 | `{表名}-{bgbh}(网页爬取)_{ts}.xlsx` |
| 生成件 4122 | `主要工程量表-{bgbh}(生成)_{ts}.xlsx` |
| 校验报告 | `校验报告-{bgbh}-{YYYYMMDD_HHMM}.xlsx` |
| 变更清单 | `变更清单-{bgbh}.xlsx` |
| 工作流清单 | `工作流清单_{bgbh}_{ts}.json` |
| 批量结果 | `批量_{ts}/批量结果_{ts}.xlsx` / `.json` |

### 14.3 缺陷代码中文名

| 代码 | 名称 | 类别 |
|---|---|---|
| AJ | 支管暗接 | 结构 |
| BX | 变形 | 结构 |
| CK | 错口 | 结构 |
| CR | 异物穿入 | 结构 |
| FS | 腐蚀 | 结构 |
| PL | 破裂 | 结构 |
| QF | 起伏 | 结构 |
| SL | 渗漏 | 结构 |
| TJ | 脱节 | 结构 |
| TL | 接口材料脱落 | 结构 |
| CJ | 沉积 | 功能 |
| CQ | 残墙、坝根 | 功能 |
| FZ | 浮渣 | 功能 |
| JG | 结垢 | 功能 |
| SG | 树根 | 功能 |
| ZW | 障碍物 | 功能 |

### 14.4 4127 缺陷槽位字段

| 槽位 | 长度 | 代码 | 分值 | 等级 | 描述 |
|---|---|---|---|---|---|
| 1 | c_716 | c_717 | c_718 | c_719 | c_720 |
| 2 | c_729 | c_730 | c_731 | c_732 | c_733 |
| 3 | c_735 | c_737 | c_739 | c_741 | c_743 |
| 4 | c_736 | c_738 | c_740 | c_742 | c_744 |

### 14.5 版本与文件

| 项 | 值 |
|---|---|
| `crawl.VERSION` | `1.0.0` |
| GUI/CLI 发布形式 | PyInstaller onedir（exe + `_internal`） |
| 本文档 | `docs/CCTV报告工具-原理讲解.md` / `.html` |

### 14.6 主要调用链

```
一键全流程
  cctv_gui.App.do_workflow
    → workflow.run
        → common.load_config / Session.login / locate_report
        → crawl.crawl_4126 / crawl_4127 / crawl_4123 / crawl_4125 / crawl_4122
        → crawl.save_412x
        → crawl.verify
        → generate.build_from_rows → generate.write_outputs
        → generate.diff_4122 / workflow._diff_4123 / _diff_4125
        → writeback.upload_4122 / upload_4123 / upload_4125
        → crawl.crawl_412x（复核）

上传勾选
  cctv_gui.App.do_upload_tables
    → uploader.upload_selected
        → common.load_config / Session.login / locate_report
        → crawl.load_outputs
        → crawl.crawl_4126 / 4127（仅勾选 4126/4127 时）
        → writeback.diff_rows → scan_4126_positions → upload
        → generate.build_from_rows
        → generate.diff_4122 / uploader._cmp_4123 / 4125 比较
        → writeback.upload_4122 / upload_4123 / upload_4125

校验报告
  cctv_gui.App.do_verify
    → verify_report.build
        → crawl.load_outputs
        → crawl.verify
        → openpyxl 写 4 个 sheet

批量
  cctv_gui.App.do_batch
    → batch.run_batch
        → batch.read_list
        → workflow.run（每个报告）
        → batch._write_summary

命令行
  cctv.py → cli.main
    → cli.build_parser / _normalize
    → cmd_* → 上述业务模块
```

### 14.7 源码阅读顺序建议

1. `cctvtools/common.py`：先搞清配置、Session、URL、表单；
2. `cctvtools/credentials.py`：账号与 DPAPI；
3. `cctvtools/crawl.py`：五张表的解析、翻页、落盘、校验；
4. `cctvtools/generate.py`：派生表算法；
5. `cctvtools/writeback.py`：编辑态、填值、保存；
6. `cctvtools/diff.py`：逐字段比较；
7. `cctvtools/workflow.py` / `uploader.py` / `batch.py`：编排；
8. `cctv_gui.py` / `cctvtools/cli.py`：界面与命令行。

---

> 本文档依据 2026-11-25 的源码与打包产物整理。若平台页面结构、字段名或保存接口发生变化，应同步修改 `cctvtools/crawl.py`、`writeback.py` 与本文档对应章节。

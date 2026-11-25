# CCTV 管道检测报告工具 v1.0.0

首个整理版：把五张平台报表的抓取、校验、生成、对比、写回收敛到一套 GUI/CLI 工具。

## 功能

- 抓取 4122/4123/4125/4126/4127，自动翻页、去重、过滤空壳页
- G1–G9 交叉校验：段数、重复、流向、缺陷数、长度、说明文字
- 由 4126+4127 生成 4122 主要工程量表、4123 缺陷汇总表、4125 缺陷数量统计表
- 平台现状与生成结果逐字段对比，一致自动跳过
- 按勾选写回 4122/4123/4125/4126/4127，支持 `--dry` 演练
- Tkinter 图形界面 + 命令行 + 批量模式 + 校验报告 xlsx
- Windows DPAPI 加密保存密码；配置文件不进仓库
- 离线自检 `selftest`，不需要联网

## 安装

```bash
git clone https://github.com/qiange6/cctv-report-tool.git
cd cctv-report-tool
python -m pip install -r requirements.txt
python cctv_gui.py
```

Windows 源码运行要求 Python 3.10+；exe 版见 Releases。

## 资产

- `CCTV报告工具_便携版_20261125.zip`：GUI + 命令行 + 各自 `_internal` 依赖 + 使用说明，解压即用。

## 文档

- [README](README.md)
- [原理讲解（Markdown）](docs/CCTV报告工具-原理讲解.md)
- [原理讲解（HTML）](docs/CCTV报告工具-原理讲解.html)
- [使用说明](工具说明.md)

## 注意

- 平台地址、页面结构、账号和报告数据均属业务敏感信息，请仅在授权范围内使用。
- 仓库不包含 `cctv_config.json`、`输出/`、`exe/`、`_internal/` 与任何业务数据。

# huachuang
coding in huachuang
# 华创固收研究自动化工作区

本仓库只管理可维护、可复用的源代码和说明文件。业务数据、账号信息、
报告成品、运行缓存和日期快照保留在本机，不上传 Git。

## 目录

- `src/daily`：日报、底稿更新、发行日历等日频流程。
- `src/weekly`：华创、人保、PA、PH 等周报流程。
- `src/clauses`：强赎、下修及条款公告流程。
- `src/valuation`：拟合溢价率和估值计算。
- `src/backtest`：策略回测。
- `src/common`：Parquet 等公共读写模块。
- `scripts`：研究分析和维护脚本；为保持历史脚本兼容，暂不继续拆层。
- `codex-skills`：独立的 Codex skills 源文件。
- `config`：不含真实账号的配置示例。
- `tests`：验证脚本和测试说明。
- `assets`：字体、图片和模板等本地资产，默认不上传。
- `data`：Excel、Parquet 等业务数据，禁止上传。
- `private`：账号、邮箱和授权信息，禁止上传。
- `runs`：按日期生成的工作目录，禁止上传。
- `outputs`、`tmp`、`archive`：成品、临时文件和备份，禁止上传。

## 运行约定

建议从项目根目录运行脚本，例如：

```powershell
py "src\daily\【日报】转债日报.py"
```

安装通用 Python 依赖：

```powershell
py -m pip install -r requirements.txt
```

`iFinDPy`、Wind、Microsoft Excel 等本机组件需要按各自客户端方式安装，
不随 Git 仓库分发。

## 安全约定

- 不使用 `git add .` 盲目提交整个工作区。
- 提交前先运行 `git status` 和 `git diff --cached --stat`。
- 真实账号、密码、授权码和业务数据不得进入 Git。
- 工作区迁移记录见 `docs/MIGRATION_20260824.md`。

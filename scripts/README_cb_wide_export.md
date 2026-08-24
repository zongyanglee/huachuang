# Parquet 分片合并为“宽表”并导入 PostgreSQL（DBeaver）

本仓库里 `D:\JupyterFiles\huachuang\转债个券历史序列\{year}\{yyyymm}.parquet` 的结构是：

- 行：`__sheet_name`（指标名） + `__row_id`（转债代码）
- 列：当月的交易日（列名形如 `2015-01-05 00:00:00`）

你要的“宽表”是：每个指标一张表，**每行=日期**，**每列=转债**。

## 1) 生成宽表文件 + 建表 DDL

在 PowerShell 里运行：

```powershell
py -3.12 D:\JupyterFiles\huachuang\scripts\cb_parquet_wide_export.py
```

输出目录（默认）：

- `D:\JupyterFiles\huachuang\cb_wide_export\data\`：每个指标一个 `csv/parquet`
- `D:\JupyterFiles\huachuang\cb_wide_export\ddl\`：每个指标一个建表 SQL
- `D:\JupyterFiles\huachuang\cb_wide_export\meta\bond_column_mapping.json`：转债代码 -> 列名映射（列名会把 `110009.SH` 变成 `b110009_SH`，方便做 PostgreSQL 列名）
- `D:\JupyterFiles\huachuang\cb_wide_export\meta\metric_registry.json`：指标 -> 文件名/表名映射

只导出指定指标（例：收盘价、转股溢价率）：

```powershell
py -3.12 D:\JupyterFiles\huachuang\scripts\cb_parquet_wide_export.py --metrics 收盘价,转股溢价率
```

## 2) 在 PostgreSQL（云端）建表

在 DBeaver 连上你的云端 PostgreSQL 后：

- 打开 SQL Editor
- 把 `D:\JupyterFiles\huachuang\cb_wide_export\ddl\收盘价.sql`（或任意指标的 SQL）内容粘贴执行

这样会创建 schema：`cb_wide`，以及表：`cb_wide."收盘价"`（表名用双引号，支持中文）

## 3) 用 DBeaver 导入 CSV

- 右键目标表（例如 `cb_wide."收盘价"`）→ `Import Data`
- Data source 选择 `CSV`
- 选择 `D:\JupyterFiles\huachuang\cb_wide_export\data\收盘价.csv`
- 确认第一列是 `date`，其余列名匹配（例如 `b110009_SH`）
- Start

提示：
- CSV 用 `utf-8-sig` 写出，DBeaver 一般能自动识别中文。
- 若某些指标本质是文本（如评级/交易状态），建表 DDL 会自动推断为 `text`；其余多数为 `double precision`。

## 4) 不用 DBeaver：用 Python 直接建表 + 导入（推荐自动化）

脚本：`D:\JupyterFiles\huachuang\scripts\cb_wide_to_postgres.py`

先安装依赖（只需一次）：

```powershell
py -3.12 -m pip install psycopg[binary]
```

建议用环境变量保存 DSN（避免把密码写进命令历史）：

```powershell
$env:PG_DSN = "postgresql://hcgszgn:<你的密码>@pgm-2zecgh9960s16bvsuo.pg.rds.aliyuncs.com:5432/zgn_db"
```

把导出的宽表（CSV）直接 upsert 到云端（默认 upsert）：

```powershell
py -3.12 D:\JupyterFiles\huachuang\scripts\cb_wide_to_postgres.py --schema cb_wide --metrics 收盘价
```

模式说明：
- `--mode replace`：先 `TRUNCATE` 再全量导入
- `--mode append`：只插入新日期（已有日期跳过）
- `--mode upsert`：按 `date` 主键插入/更新（推荐）

# Huachuang workspace migration manifest

Date: 2026-08-24

This manifest records the top-level reorganization. No listed item is deleted.
To roll back an individual move, move the destination back to the source path.

## Source code

- `【更新】日报数据更新.py` -> `src/daily/【更新】日报数据更新.py`
- `【日报】转债日报.py` -> `src/daily/【日报】转债日报.py`
- `底稿更新.py` -> `src/daily/底稿更新.py`
- `日报文本.py` -> `src/daily/日报文本.py`
- `发行日历.py` -> `src/daily/发行日历.py`
- `转债每日数量余额汇总.py` -> `src/daily/转债每日数量余额汇总.py`
- `华创转债周报自动更新.py` -> `src/weekly/华创转债周报自动更新.py`
- `人保周报自动更新.py` -> `src/weekly/人保周报自动更新.py`
- `PA周报.py` -> `src/weekly/PA周报.py`
- `PH周报.py` -> `src/weekly/PH周报.py`
- `【条款】条款公告更新.py` -> `src/clauses/【条款】条款公告更新.py`
- `【条款】P强赎进度跟踪.py` -> `src/clauses/【条款】P强赎进度跟踪.py`
- `【条款】P下修进度跟踪.py` -> `src/clauses/【条款】P下修进度跟踪.py`
- `【高频】百元拟合溢价率.py` -> `src/valuation/【高频】百元拟合溢价率.py`
- `【计算】单一因子分组幂衰减拟合溢价率.py` -> `src/valuation/【计算】单一因子分组幂衰减拟合溢价率.py`
- `【计算】分组多因子修正拟合溢价率.py` -> `src/valuation/【计算】分组多因子修正拟合溢价率.py`
- `【回测】回测合集.py` -> `src/backtest/【回测】回测合集.py`
- `转债Parquet标准读写模块.py` -> `src/common/转债Parquet标准读写模块.py`

The existing `scripts/` directory remains at its original depth because many
scripts calculate the project root using `Path(__file__).parents[1]`.

## Private files

- `ifind账号.txt` -> `private/ifind账号.txt`
- `条款跟踪邮箱列表.xlsx` -> `private/条款跟踪邮箱列表.xlsx`
- `设置126邮箱授权码.py` -> `private/设置126邮箱授权码.py`

## Data and assets

- `转债个券历史序列` -> `data/转债个券历史序列`
- `【华创固收】赎回和不赎回公告统计.xlsx` -> `data/clauses/【华创固收】赎回和不赎回公告统计.xlsx`
- `【华创固收】下修和不下修公告统计.xlsx` -> `data/clauses/【华创固收】下修和不下修公告统计.xlsx`
- `KaiTi_GB2312.ttf` -> `assets/fonts/KaiTi_GB2312.ttf`
- `条款表头.png` -> `assets/images/条款表头.png`
- `PPT模版.pptx` -> `assets/templates/PPT模版.pptx`

## Dated runs and generated workspaces

- `0821数据更新` -> `runs/daily/0821数据更新`
- `0822数据更新` -> `runs/daily/0822数据更新`
- `0823数据更新` -> `runs/daily/0823数据更新`
- `20260821_转债日报` -> `runs/daily/20260821_转债日报`
- `【华创】转债周报20260821` -> `runs/weekly/【华创】转债周报20260821`
- `人保周报20260824` -> `runs/weekly/人保周报20260824`
- `PA周报20260821` -> `runs/weekly/PA周报20260821`
- `策略回测20260820` -> `runs/research/策略回测20260820`
- `策略回测20260821` -> `runs/research/策略回测20260821`
- `多因子修正拟合溢价率_20260821` -> `runs/research/多因子修正拟合溢价率_20260821`
- `PPT策略净值图_20260819` -> `runs/presentations/PPT策略净值图_20260819`

## Codex skills and archives

- `bond-report-frontmatter` -> `codex-skills/bond-report-frontmatter`
- `docx-to-pptx` -> `codex-skills/docx-to-pptx`
- `report-proofreader` -> `codex-skills/report-proofreader`
- `report-proofreader.skill` -> `codex-skills/packages/report-proofreader.skill`
- `backups` -> `archive/backups`

## Temporary directories

- `__pycache__` -> `tmp/cache/root_pycache`
- `.ipynb_checkpoints` -> `tmp/cache/ipynb_checkpoints`
- `.cache_multifactor_fit` -> `tmp/cache/multifactor_fit`
- `.cache_parquet` -> `tmp/cache/parquet`
- `.cache_power_decay_fit` -> `tmp/cache/power_decay_fit`
- `_tmp_email_split` -> `tmp/legacy/email_split`
- `.codex_tmp` -> `tmp/legacy/codex_tmp`
- `.codex_tmp_display_rebuild` -> `tmp/legacy/codex_tmp_display_rebuild`
- `.codex_tmp_skill_expand` -> `tmp/legacy/codex_tmp_skill_expand`
- `.codex_doc_review` -> `tmp/review/codex_doc_review`

## Intentionally retained at repository root

- `.git`, `.gitignore`, `.agents`, `.vscode`
- `README.md`, `docs`, `config`, `tests`
- `src`, `scripts`, `data`, `assets`, `private`, `runs`, `outputs`, `tmp`,
  `archive`, `codex-skills`
- `node_modules` junction, temporarily retained so existing Node scripts continue
  resolving installed packages. It remains excluded from Git.

param(
    [switch]$Undo
)

$ErrorActionPreference = "Stop"
$workspaceRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..")).Path

$moves = @(
    @{ Source = "【更新】日报数据更新.py"; Destination = "src\daily\【更新】日报数据更新.py" },
    @{ Source = "【日报】转债日报.py"; Destination = "src\daily\【日报】转债日报.py" },
    @{ Source = "底稿更新.py"; Destination = "src\daily\底稿更新.py" },
    @{ Source = "日报文本.py"; Destination = "src\daily\日报文本.py" },
    @{ Source = "发行日历.py"; Destination = "src\daily\发行日历.py" },
    @{ Source = "转债每日数量余额汇总.py"; Destination = "src\daily\转债每日数量余额汇总.py" },
    @{ Source = "华创转债周报自动更新.py"; Destination = "src\weekly\华创转债周报自动更新.py" },
    @{ Source = "人保周报自动更新.py"; Destination = "src\weekly\人保周报自动更新.py" },
    @{ Source = "PA周报.py"; Destination = "src\weekly\PA周报.py" },
    @{ Source = "PH周报.py"; Destination = "src\weekly\PH周报.py" },
    @{ Source = "【条款】条款公告更新.py"; Destination = "src\clauses\【条款】条款公告更新.py" },
    @{ Source = "【条款】P强赎进度跟踪.py"; Destination = "src\clauses\【条款】P强赎进度跟踪.py" },
    @{ Source = "【条款】P下修进度跟踪.py"; Destination = "src\clauses\【条款】P下修进度跟踪.py" },
    @{ Source = "【高频】百元拟合溢价率.py"; Destination = "src\valuation\【高频】百元拟合溢价率.py" },
    @{ Source = "【计算】单一因子分组幂衰减拟合溢价率.py"; Destination = "src\valuation\【计算】单一因子分组幂衰减拟合溢价率.py" },
    @{ Source = "【计算】分组多因子修正拟合溢价率.py"; Destination = "src\valuation\【计算】分组多因子修正拟合溢价率.py" },
    @{ Source = "【回测】回测合集.py"; Destination = "src\backtest\【回测】回测合集.py" },
    @{ Source = "转债Parquet标准读写模块.py"; Destination = "src\common\转债Parquet标准读写模块.py" },
    @{ Source = "ifind账号.txt"; Destination = "private\ifind账号.txt" },
    @{ Source = "条款跟踪邮箱列表.xlsx"; Destination = "private\条款跟踪邮箱列表.xlsx" },
    @{ Source = "设置126邮箱授权码.py"; Destination = "private\设置126邮箱授权码.py" },
    @{ Source = "转债个券历史序列"; Destination = "data\转债个券历史序列" },
    @{ Source = "【华创固收】赎回和不赎回公告统计.xlsx"; Destination = "data\clauses\【华创固收】赎回和不赎回公告统计.xlsx" },
    @{ Source = "【华创固收】下修和不下修公告统计.xlsx"; Destination = "data\clauses\【华创固收】下修和不下修公告统计.xlsx" },
    @{ Source = "KaiTi_GB2312.ttf"; Destination = "assets\fonts\KaiTi_GB2312.ttf" },
    @{ Source = "条款表头.png"; Destination = "assets\images\条款表头.png" },
    @{ Source = "PPT模版.pptx"; Destination = "assets\templates\PPT模版.pptx" },
    @{ Source = "0821数据更新"; Destination = "runs\daily\0821数据更新" },
    @{ Source = "0822数据更新"; Destination = "runs\daily\0822数据更新" },
    @{ Source = "0823数据更新"; Destination = "runs\daily\0823数据更新" },
    @{ Source = "20260821_转债日报"; Destination = "runs\daily\20260821_转债日报" },
    @{ Source = "【华创】转债周报20260821"; Destination = "runs\weekly\【华创】转债周报20260821" },
    @{ Source = "人保周报20260824"; Destination = "runs\weekly\人保周报20260824" },
    @{ Source = "PA周报20260821"; Destination = "runs\weekly\PA周报20260821" },
    @{ Source = "策略回测20260820"; Destination = "runs\research\策略回测20260820" },
    @{ Source = "策略回测20260821"; Destination = "runs\research\策略回测20260821" },
    @{ Source = "多因子修正拟合溢价率_20260821"; Destination = "runs\research\多因子修正拟合溢价率_20260821" },
    @{ Source = "PPT策略净值图_20260819"; Destination = "runs\presentations\PPT策略净值图_20260819" },
    @{ Source = "bond-report-frontmatter"; Destination = "codex-skills\bond-report-frontmatter" },
    @{ Source = "docx-to-pptx"; Destination = "codex-skills\docx-to-pptx" },
    @{ Source = "report-proofreader"; Destination = "codex-skills\report-proofreader" },
    @{ Source = "report-proofreader.skill"; Destination = "codex-skills\packages\report-proofreader.skill" },
    @{ Source = "backups"; Destination = "archive\backups" },
    @{ Source = "__pycache__"; Destination = "tmp\cache\root_pycache" },
    @{ Source = ".ipynb_checkpoints"; Destination = "tmp\cache\ipynb_checkpoints" },
    @{ Source = ".cache_multifactor_fit"; Destination = "tmp\cache\multifactor_fit" },
    @{ Source = ".cache_parquet"; Destination = "tmp\cache\parquet" },
    @{ Source = ".cache_power_decay_fit"; Destination = "tmp\cache\power_decay_fit" },
    @{ Source = "_tmp_email_split"; Destination = "tmp\legacy\email_split" },
    @{ Source = ".codex_tmp"; Destination = "tmp\legacy\codex_tmp" },
    @{ Source = ".codex_tmp_display_rebuild"; Destination = "tmp\legacy\codex_tmp_display_rebuild" },
    @{ Source = ".codex_tmp_skill_expand"; Destination = "tmp\legacy\codex_tmp_skill_expand" },
    @{ Source = ".codex_doc_review"; Destination = "tmp\review\codex_doc_review" }
)

function Resolve-WorkspacePath([string]$relativePath) {
    $resolved = [IO.Path]::GetFullPath((Join-Path $workspaceRoot $relativePath))
    $prefix = $workspaceRoot + [IO.Path]::DirectorySeparatorChar
    if (-not $resolved.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path escaped workspace: $resolved"
    }
    return $resolved
}

$results = foreach ($move in $moves) {
    $sourceRelative = if ($Undo) { $move.Destination } else { $move.Source }
    $destinationRelative = if ($Undo) { $move.Source } else { $move.Destination }
    $source = Resolve-WorkspacePath $sourceRelative
    $destination = Resolve-WorkspacePath $destinationRelative

    $sourceExists = Test-Path -LiteralPath $source
    $destinationExists = Test-Path -LiteralPath $destination
    if (-not $sourceExists -and $destinationExists) {
        [PSCustomObject]@{ Status = "AlreadyMoved"; Source = $sourceRelative; Destination = $destinationRelative }
        continue
    }
    if (-not $sourceExists) {
        throw "Source does not exist: $source"
    }
    if ($destinationExists) {
        throw "Destination already exists: $destination"
    }

    $parent = Split-Path -Parent $destination
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    Move-Item -LiteralPath $source -Destination $destination
    [PSCustomObject]@{ Status = "Moved"; Source = $sourceRelative; Destination = $destinationRelative }
}

$results

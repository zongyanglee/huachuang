$ErrorActionPreference = "Stop"
$workspaceRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..")).Path

$replacements = [ordered]@{
    'Path(r"D:\JupyterFiles\huachuang")' = 'Path(__file__).resolve().parents[2]'
    '"转债个券历史序列"' = '"data/转债个券历史序列"'
    "'转债个券历史序列'" = "'data/转债个券历史序列'"
    '"ifind账号.txt"' = '"private/ifind账号.txt"'
    "'ifind账号.txt'" = "'private/ifind账号.txt'"
    '"条款表头.png"' = '"assets/images/条款表头.png"'
    "'条款表头.png'" = "'assets/images/条款表头.png'"
    '"KaiTi_GB2312.ttf"' = '"assets/fonts/KaiTi_GB2312.ttf"'
    "'KaiTi_GB2312.ttf'" = "'assets/fonts/KaiTi_GB2312.ttf'"
    '"PPT模版.pptx"' = '"assets/templates/PPT模版.pptx"'
    "'PPT模版.pptx'" = "'assets/templates/PPT模版.pptx'"
    '"【华创固收】赎回和不赎回公告统计.xlsx"' = '"data/clauses/【华创固收】赎回和不赎回公告统计.xlsx"'
    "'【华创固收】赎回和不赎回公告统计.xlsx'" = "'data/clauses/【华创固收】赎回和不赎回公告统计.xlsx'"
    '"【华创固收】下修和不下修公告统计.xlsx"' = '"data/clauses/【华创固收】下修和不下修公告统计.xlsx"'
    "'【华创固收】下修和不下修公告统计.xlsx'" = "'data/clauses/【华创固收】下修和不下修公告统计.xlsx'"
    '"backups"' = '"archive/backups"'
    "'backups'" = "'archive/backups'"
    '".cache_parquet"' = '"tmp/cache/parquet"'
    "'.cache_parquet'" = "'tmp/cache/parquet'"
    '".cache_power_decay_fit"' = '"tmp/cache/power_decay_fit"'
    "'.cache_power_decay_fit'" = "'tmp/cache/power_decay_fit'"
    '".cache_multifactor_fit"' = '"tmp/cache/multifactor_fit"'
    "'.cache_multifactor_fit'" = "'tmp/cache/multifactor_fit'"
    'WORKSPACE / f"{run_date:%Y%m%d}_转债日报"' = 'WORKSPACE / "runs" / "daily" / f"{run_date:%Y%m%d}_转债日报"'
    'SCRIPT_DIR / f"{payload_date.strftime(''%m%d'')}数据更新"' = 'SCRIPT_DIR / "runs" / "daily" / f"{payload_date.strftime(''%m%d'')}数据更新"'
    'ROOT / f"{OUTPUT_MMDD}数据更新"' = 'ROOT / "runs" / "daily" / f"{OUTPUT_MMDD}数据更新"'
    'ROOT / f"【华创】转债周报{week_end:%Y%m%d}"' = 'ROOT / "runs" / "weekly" / f"【华创】转债周报{week_end:%Y%m%d}"'
    'ROOT / f"人保周报{report_date:%Y%m%d}"' = 'ROOT / "runs" / "weekly" / f"人保周报{report_date:%Y%m%d}"'
    'ROOT / f"PA周报{LATEST_TRADE_DATE:%Y%m%d}"' = 'ROOT / "runs" / "weekly" / f"PA周报{LATEST_TRADE_DATE:%Y%m%d}"'
    'ROOT / f"鹏华周报{date_tag}"' = 'ROOT / "runs" / "weekly" / f"鹏华周报{date_tag}"'
    'WORKSPACE / f"幂衰减拟合溢价率_{RUN_DATE}"' = 'WORKSPACE / "runs" / "research" / f"幂衰减拟合溢价率_{RUN_DATE}"'
    'WORKSPACE / f"多因子修正拟合溢价率_{RUN_DATE}"' = 'WORKSPACE / "runs" / "research" / f"多因子修正拟合溢价率_{RUN_DATE}"'
}

function Read-TextFile([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encoding = [Text.UTF8Encoding]::new($true, $true)
        return @{ Text = $encoding.GetString($bytes, 3, $bytes.Length - 3); Encoding = $encoding; Preamble = $encoding.GetPreamble() }
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encoding = [Text.UnicodeEncoding]::new($false, $true, $true)
        return @{ Text = $encoding.GetString($bytes, 2, $bytes.Length - 2); Encoding = $encoding; Preamble = $encoding.GetPreamble() }
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encoding = [Text.UnicodeEncoding]::new($true, $true, $true)
        return @{ Text = $encoding.GetString($bytes, 2, $bytes.Length - 2); Encoding = $encoding; Preamble = $encoding.GetPreamble() }
    }
    $encoding = [Text.UTF8Encoding]::new($false, $true)
    return @{ Text = $encoding.GetString($bytes); Encoding = $encoding; Preamble = [byte[]]@() }
}

function Write-TextFile([string]$path, [string]$text, $encoding, [byte[]]$preamble) {
    $content = $encoding.GetBytes($text)
    if ($preamble.Length -eq 0) {
        [IO.File]::WriteAllBytes($path, $content)
        return
    }
    $output = [byte[]]::new($preamble.Length + $content.Length)
    [Array]::Copy($preamble, 0, $output, 0, $preamble.Length)
    [Array]::Copy($content, 0, $output, $preamble.Length, $content.Length)
    [IO.File]::WriteAllBytes($path, $output)
}

$roots = @(
    (Join-Path $workspaceRoot "src"),
    (Join-Path $workspaceRoot "scripts"),
    (Join-Path $workspaceRoot "private")
)
$maintenanceRoot = Join-Path $workspaceRoot "scripts\maintenance"
$files = foreach ($root in $roots) {
    Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object {
        $_.Extension -in ".py", ".js", ".mjs", ".ps1" -and
        -not $_.FullName.StartsWith($maintenanceRoot, [StringComparison]::OrdinalIgnoreCase)
    }
}

$results = foreach ($file in $files) {
    try {
        $document = Read-TextFile $file.FullName
    }
    catch {
        [PSCustomObject]@{ Status = "SkippedEncoding"; Path = $file.FullName; Replacements = 0 }
        continue
    }
    $updated = $document.Text
    $count = 0
    foreach ($entry in $replacements.GetEnumerator()) {
        $before = $updated
        $updated = $updated.Replace($entry.Key, $entry.Value)
        if ($updated -ne $before) { $count++ }
    }

    if ($file.FullName.StartsWith((Join-Path $workspaceRoot "src"), [StringComparison]::OrdinalIgnoreCase) -and
        $updated.Contains("from 转债Parquet标准读写模块 import") -and
        -not $updated.Contains("_COMMON_MODULE_DIR =")) {
        $bootstrap = @'
import sys

_COMMON_MODULE_DIR = Path(__file__).resolve().parents[2]s[1] / "common"
if str(_COMMON_MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(_COMMON_MODULE_DIR))

from 转债Parquet标准读写模块 import
'@
        $updated = $updated.Replace("from 转债Parquet标准读写模块 import", $bootstrap.TrimEnd("`r", "`n"))
        $count++
    }

    if ($updated -eq $document.Text) {
        [PSCustomObject]@{ Status = "Unchanged"; Path = $file.FullName; Replacements = 0 }
        continue
    }
    Write-TextFile $file.FullName $updated $document.Encoding $document.Preamble
    [PSCustomObject]@{ Status = "Updated"; Path = $file.FullName; Replacements = $count }
}

$results | Where-Object Status -ne "Unchanged"

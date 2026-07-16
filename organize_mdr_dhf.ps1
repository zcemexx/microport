[CmdletBinding()]
param(
    [string]$SourceRoot = 'D:\ShareCache\LS_SH(上海生命科技)\1、研发部\2、CGM\CGM-欧盟注册-项目管理文件',
    [string]$GoldRoot = 'D:\ShareCache\LS_SH(上海生命科技)\1、研发部\2、CGM\CGM-MDR注册\Technical Documentation-final',
    [string]$ListPath = '',
    [string]$Destination = '',
    [switch]$Execute
)

$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 may not populate $PSScriptRoot while evaluating default parameters.
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ScriptDir)) { $ScriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ListPath)) {
    $ListPath = Join-Path $ScriptDir '项目管理文件编号.txt'
}
if ([string]::IsNullOrWhiteSpace($Destination)) {
    $DesktopDir = [Environment]::GetFolderPath('Desktop')
    if ([string]::IsNullOrWhiteSpace($DesktopDir)) { $DesktopDir = Join-Path $HOME 'Desktop' }
    $Destination = Join-Path $DesktopDir 'DHF'
}
$PreviewCsv = Join-Path $ScriptDir 'DHF_整理预览.csv'
$MissingCsv = Join-Path $ScriptDir 'DHF_未找到编号.csv'
$LogPath = Join-Path $ScriptDir 'DHF_整理日志.txt'

function Assert-Path([string]$Path, [string]$Label) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { throw "$Label 不存在：$Path" }
}
function Get-SafeTargetPath([string]$Folder, [string]$Name) {
    $candidate = Join-Path $Folder $Name
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    $base = [IO.Path]::GetFileNameWithoutExtension($Name)
    $ext = [IO.Path]::GetExtension($Name)
    $n = 2
    do {
        $candidate = Join-Path $Folder ("{0} [{1}]{2}" -f $base, $n, $ext)
        $n++
    } while (Test-Path -LiteralPath $candidate)
    return $candidate
}
function Test-DocumentNumber([string]$FileName, [string]$Number, [bool]$IsPrefix) {
    $escaped = [Regex]::Escape($Number)
    if ($IsPrefix) {
        return $FileName -match "(?i)(?<![A-Z0-9])$escaped"
    }
    # Prevent S0023-9008 from matching sub-number S0023-9008-01.
    return $FileName -match "(?i)(?<![A-Z0-9])$escaped(?![A-Z0-9]|-\d)"
}

Assert-Path $SourceRoot '项目管理文件目录'
Assert-Path $GoldRoot 'Gold Standard 目录'
Assert-Path $ListPath '项目管理文件编号清单'

$rows = Import-Csv -LiteralPath $ListPath -Delimiter "`t" -Encoding UTF8
if (-not $rows) { throw '编号清单为空。' }
if (-not ($rows[0].PSObject.Properties.Name -contains '文件编号')) {
    throw '清单中未找到“文件编号”列。'
}

# 目录只扫描一次，避免针对 187 个编号重复遍历磁盘。
Write-Host '正在一次性扫描两个目录……' -ForegroundColor Cyan
$sourceFiles = @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse | Where-Object { $_.Name -notlike '~$*' })
$goldFiles = @(Get-ChildItem -LiteralPath $GoldRoot -File -Recurse | Where-Object { $_.Name -notlike '~$*' })

# Windows PowerShell 5.1-native collections.
$plan = @()
$missing = @()
$seenSelection = @{}

foreach ($row in $rows) {
    $rawNumber = ([string]$row.'文件编号').Trim()
    if ([string]::IsNullOrWhiteSpace($rawNumber)) { continue }
    $isPrefix = $rawNumber.EndsWith('~')
    $number = $rawNumber.TrimEnd([char]'~').Trim()

    $src = @($sourceFiles | Where-Object { Test-DocumentNumber $_.Name $number $isPrefix })
    $gold = @($goldFiles | Where-Object { Test-DocumentNumber $_.Name $number $isPrefix })
    $extensionList = @()
    foreach ($matchedFile in $src) { $extensionList += $matchedFile.Extension.ToLowerInvariant() }
    foreach ($matchedFile in $gold) { $extensionList += $matchedFile.Extension.ToLowerInvariant() }
    $allExts = @($extensionList | Sort-Object -Unique)

    if ($allExts.Count -eq 0) {
        $missing += [pscustomobject][ordered]@{
            '文件名称' = $row.'文件名称'
            '文件编号' = $rawNumber
            '状态' = $row.'状态'
            '申请人' = $row.'申请人'
        }
        continue
    }

    foreach ($ext in $allExts) {
        $goldForExt = @($gold | Where-Object { $_.Extension.ToLowerInvariant() -eq $ext })
        $srcForExt = @($src | Where-Object { $_.Extension.ToLowerInvariant() -eq $ext })
        # 同一“编号+扩展名”只要 Gold 中存在，就完全采用 Gold；否则采用项目管理目录。
        $chosen = @()
        $origin = '项目管理文件'
        $suppressed = 0
        if ($goldForExt.Count -gt 0) {
            $chosen = $goldForExt
            $origin = 'Gold Standard'
            $suppressed = $srcForExt.Count
        } else {
            $chosen = $srcForExt
        }

        foreach ($file in $chosen) {
            $selectionKey = $file.FullName.ToLowerInvariant()
            if ($seenSelection.ContainsKey($selectionKey)) { continue }
            $seenSelection[$selectionKey] = $true
            $initialResult = '预览'
            if ($Execute) { $initialResult = '待复制' }
            $plan += [pscustomobject][ordered]@{
                '文件编号' = $rawNumber
                '清单文件名称' = $row.'文件名称'
                '选用来源' = $origin
                '文件名' = $file.Name
                '原路径' = $file.FullName
                '扩展名' = $file.Extension
                '被Gold替代的源文件数' = $suppressed
                '原路径长度' = $file.FullName.Length
                '目标路径' = (Join-Path $Destination $file.Name)
                '执行结果' = $initialResult
            }
        }
    }
}

if ($plan.Count -gt 0) {
    $plan | Sort-Object '文件编号', '选用来源', '文件名' | Export-Csv -LiteralPath $PreviewCsv -NoTypeInformation -Encoding UTF8
} else {
    '未找到任何匹配文件。' | Set-Content -LiteralPath $PreviewCsv -Encoding UTF8
}
if ($missing.Count -gt 0) {
    $missing | Sort-Object '文件编号' | Export-Csv -LiteralPath $MissingCsv -NoTypeInformation -Encoding UTF8
} else {
    '所有编号均至少匹配到一个文件。' | Set-Content -LiteralPath $MissingCsv -Encoding UTF8
}

if (-not $Execute) {
    Write-Host ''
    Write-Host '仅完成预览，未复制任何文件。' -ForegroundColor Yellow
    Write-Host "候选文件：$($plan.Count)；未找到编号：$($missing.Count)"
    Write-Host "预览表：$PreviewCsv"
    Write-Host "未找到：$MissingCsv"
    Write-Host '确认预览后，使用同一命令并追加 -Execute。' -ForegroundColor Green
    return
}

New-Item -ItemType Directory -Path $Destination -Force | Out-Null
$copied = 0
$copyErrors = 0
foreach ($item in $plan) {
    try {
        $target = Get-SafeTargetPath $Destination $item.'文件名'
        Copy-Item -LiteralPath $item.'原路径' -Destination $target -ErrorAction Stop
        $item.'目标路径' = $target
        $item.'执行结果' = '已复制'
        $copied++
    } catch {
        $item.'执行结果' = "失败：$($_.Exception.Message)"
        $copyErrors++
    }
}
$plan | Sort-Object '文件编号', '选用来源', '文件名' | Export-Csv -LiteralPath $PreviewCsv -NoTypeInformation -Encoding UTF8

@(
    "执行时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "项目管理目录：$SourceRoot"
    "Gold目录：$GoldRoot"
    "输出目录：$Destination"
    "清单编号数：$($rows.Count)"
    "成功复制：$copied"
    "复制失败：$copyErrors"
    "未找到编号：$($missing.Count)"
    "详细清单：$PreviewCsv"
    "未找到清单：$MissingCsv"
) | Set-Content -LiteralPath $LogPath -Encoding UTF8

Write-Host ''
Write-Host "完成：成功复制 $copied 个文件；失败 $copyErrors；未找到编号 $($missing.Count)。" -ForegroundColor Green
Write-Host "DHF目录：$Destination"
Write-Host "日志：$LogPath"

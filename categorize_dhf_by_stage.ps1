<#
.SYNOPSIS
    将桌面 DHF 文件夹中的文件按项目阶段复制到 DHF_categorized。
.DESCRIPTION
    从"项目管理文件编号(阶段).txt"读取文件编号与阶段的对应关系，
    按文件名中的编号将文件复制到以下五个目录：
      T1&2、T3、T4、T5、T6

    默认仅生成预览；添加 -Execute 后才实际创建目录并复制文件。
    源文件不会被移动、修改或删除。
.PARAMETER Source
    待分类的 DHF 文件夹。默认为桌面 DHF。
.PARAMETER ListPath
    带"阶段"和"文件编号"列的 TSV 清单。
    默认为脚本同目录的"项目管理文件编号(阶段).txt"。
.PARAMETER Destination
    分类输出目录。默认为桌面 DHF_categorized。
.PARAMETER Execute
    实际执行复制。未指定时仅生成预览。
.PARAMETER MaxPathLength
    目标完整路径最大长度，默认 259。
.EXAMPLE
    # 仅预览
    .\categorize_dhf_by_stage.ps1
.EXAMPLE
    # 确认预览后实际复制
    .\categorize_dhf_by_stage.ps1 -Execute
#>
[CmdletBinding()]
param(
    [string]$Source = '',
    [string]$ListPath = '',
    [string]$Destination = '',
    [switch]$Execute,
    [int]$MaxPathLength = 259
)

$ErrorActionPreference = 'Stop'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}

$DesktopDir = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($DesktopDir)) {
    $DesktopDir = Join-Path $HOME 'Desktop'
}
if ([string]::IsNullOrWhiteSpace($Source)) {
    $Source = Join-Path $DesktopDir 'DHF'
}
if ([string]::IsNullOrWhiteSpace($ListPath)) {
    $ListPath = Join-Path $ScriptDir '项目管理文件编号(阶段).txt'
}
if ([string]::IsNullOrWhiteSpace($Destination)) {
    $Destination = Join-Path $DesktopDir 'DHF_categorized'
}

$TimeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$PreviewCsv = Join-Path $ScriptDir "DHF_分类预览_$TimeStamp.csv"
$UnmatchedCsv = Join-Path $ScriptDir "DHF_分类未匹配_$TimeStamp.csv"
$SummaryLog = Join-Path $ScriptDir "DHF_分类汇总_$TimeStamp.txt"

function Assert-ExistingPath {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        throw "$Label 不存在：$Path"
    }
}

function Test-DocumentNumber {
    param(
        [string]$FileName,
        [string]$Number,
        [bool]$IsPrefix
    )
    $escaped = [Regex]::Escape($Number)
    if ($IsPrefix) {
        # 清单编号以 ~ 结尾时，允许编号后继续出现版本或子编号。
        return $FileName -match "(?i)(?<![A-Za-z0-9])$escaped"
    }
    # 普通编号不能误匹配更长编号，例如 9008 不匹配 9008-01。
    return $FileName -match "(?i)(?<![A-Za-z0-9])$escaped(?![A-Za-z0-9]|-\d)"
}

Assert-ExistingPath -Path $Source -Label '源 DHF 文件夹'
Assert-ExistingPath -Path $ListPath -Label '阶段清单'

$stageFolderMap = @{
    'T1&T2' = 'T1&2'
    'T1&2'  = 'T1&2'
    'T3'    = 'T3'
    'T4'    = 'T4'
    'T5'    = 'T5'
    'T6'    = 'T6'
}
$targetFolders = @('T1&2', 'T3', 'T4', 'T5', 'T6')

$rows = @(Import-Csv -LiteralPath $ListPath -Delimiter "`t" -Encoding UTF8)
if ($rows.Count -eq 0) {
    throw '阶段清单为空。'
}
$headers = @($rows[0].PSObject.Properties.Name)
if ($headers -notcontains '阶段' -or $headers -notcontains '文件编号') {
    throw '阶段清单必须包含"阶段"和"文件编号"列。'
}

$entries = @()
$invalidStages = @()
foreach ($row in $rows) {
    $stage = ([string]$row.'阶段').Trim()
    $rawNumber = ([string]$row.'文件编号').Trim()
    if ([string]::IsNullOrWhiteSpace($stage) -or [string]::IsNullOrWhiteSpace($rawNumber)) {
        continue
    }
    if (-not $stageFolderMap.ContainsKey($stage)) {
        $invalidStages += "$stage / $rawNumber"
        continue
    }
    $isPrefix = $rawNumber.EndsWith('~')
    $number = $rawNumber.TrimEnd([char]'~').Trim()
    $entries += [pscustomobject][ordered]@{
        RawNumber = $rawNumber
        Number = $number
        IsPrefix = $isPrefix
        Stage = $stage
        StageFolder = $stageFolderMap[$stage]
    }
}
if ($invalidStages.Count -gt 0) {
    throw "清单包含不支持的阶段：$($invalidStages -join '；')"
}

# 已整理完成的桌面 DHF 是平铺目录；这里只处理其顶层文件，避免意外遍历其他目录。
$files = @(Get-ChildItem -LiteralPath $Source -File -Force -ErrorAction Stop)
$plan = @()
$unmatched = @()

foreach ($file in $files) {
    $matches = @()
    foreach ($entry in $entries) {
        if (Test-DocumentNumber -FileName $file.Name -Number $entry.Number -IsPrefix $entry.IsPrefix) {
            $matches += $entry
        }
    }

    if ($matches.Count -eq 0) {
        $unmatched += [pscustomobject][ordered]@{
            '文件名' = $file.Name
            '原路径' = $file.FullName
            '状态' = '未在阶段清单中识别到文件编号'
            '匹配编号' = ''
            '匹配阶段' = ''
        }
        continue
    }

    $matchedStageFolders = @($matches | Select-Object -ExpandProperty StageFolder -Unique)
    if ($matchedStageFolders.Count -ne 1) {
        $unmatched += [pscustomobject][ordered]@{
            '文件名' = $file.Name
            '原路径' = $file.FullName
            '状态' = '同一文件匹配到多个阶段，未复制'
            '匹配编号' = (@($matches | Select-Object -ExpandProperty RawNumber -Unique) -join '；')
            '匹配阶段' = ($matchedStageFolders -join '；')
        }
        continue
    }

    $stageFolder = $matchedStageFolders[0]
    $targetFolder = Join-Path $Destination $stageFolder
    $targetPath = Join-Path $targetFolder $file.Name
    $pathCheck = ''
    if ($targetPath.Length -gt $MaxPathLength) {
        $pathCheck = "目标路径过长($($targetPath.Length)>$MaxPathLength)"
    }
    $initialResult = if ($Execute) { '待复制' } else { '预览' }

    $plan += [pscustomobject][ordered]@{
        '文件名' = $file.Name
        '匹配编号' = (@($matches | Select-Object -ExpandProperty RawNumber -Unique) -join '；')
        '阶段' = $stageFolder
        '原路径' = $file.FullName
        '文件大小' = $file.Length
        '目标路径' = $targetPath
        '目标路径长度' = $targetPath.Length
        '路径校验' = $pathCheck
        '执行结果' = $initialResult
    }
}

if (-not $Execute) {
    if ($plan.Count -gt 0) {
        $plan | Sort-Object '阶段', '匹配编号', '文件名' |
            Export-Csv -LiteralPath $PreviewCsv -NoTypeInformation -Encoding UTF8
    } else {
        '没有可分类的文件。' | Set-Content -LiteralPath $PreviewCsv -Encoding UTF8
    }
    if ($unmatched.Count -gt 0) {
        $unmatched | Sort-Object '文件名' |
            Export-Csv -LiteralPath $UnmatchedCsv -NoTypeInformation -Encoding UTF8
    } else {
        '所有文件均已匹配阶段。' | Set-Content -LiteralPath $UnmatchedCsv -Encoding UTF8
    }

    Write-Host ''
    Write-Host '仅生成分类预览，未创建目录、未复制文件。' -ForegroundColor Yellow
    Write-Host "源文件数：$($files.Count)；计划复制：$($plan.Count)；未匹配/歧义：$($unmatched.Count)"
    Write-Host "预览表：$PreviewCsv"
    Write-Host "未匹配：$UnmatchedCsv"
    Write-Host '确认后使用同一命令并添加 -Execute。' -ForegroundColor Green
    return
}

New-Item -ItemType Directory -Path $Destination -Force | Out-Null
foreach ($folderName in $targetFolders) {
    New-Item -ItemType Directory -Path (Join-Path $Destination $folderName) -Force | Out-Null
}

$copied = 0
$failed = 0
$skipped = 0
$index = 0
foreach ($item in $plan) {
    $index++
    if (-not [string]::IsNullOrWhiteSpace($item.'路径校验')) {
        $item.'执行结果' = "跳过：$($item.'路径校验')"
        $skipped++
        Write-Host "[$index/$($plan.Count)] 跳过：$($item.'文件名')" -ForegroundColor Yellow
        continue
    }

    try {
        # 分类目录是派生副本；重复执行时用当前 DHF 文件覆盖同名旧副本，避免产生 [2]。
        Copy-Item -LiteralPath $item.'原路径' -Destination $item.'目标路径' -Force -ErrorAction Stop
        $sourceLength = (Get-Item -LiteralPath $item.'原路径' -ErrorAction Stop).Length
        $targetLength = (Get-Item -LiteralPath $item.'目标路径' -ErrorAction Stop).Length
        if ($sourceLength -ne $targetLength) {
            throw "复制后文件大小不一致：源=$sourceLength，目标=$targetLength"
        }
        $item.'执行结果' = '已复制并通过大小校验'
        $copied++
        Write-Host "[$index/$($plan.Count)] 已复制：$($item.'阶段')\$($item.'文件名')"
    } catch {
        $item.'执行结果' = "失败：$($_.Exception.Message)"
        $failed++
        Write-Host "[$index/$($plan.Count)] 失败：$($item.'文件名')：$($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($plan.Count -gt 0) {
    $plan | Sort-Object '阶段', '匹配编号', '文件名' |
        Export-Csv -LiteralPath $PreviewCsv -NoTypeInformation -Encoding UTF8
} else {
    '没有可分类的文件。' | Set-Content -LiteralPath $PreviewCsv -Encoding UTF8
}
if ($unmatched.Count -gt 0) {
    $unmatched | Sort-Object '文件名' |
        Export-Csv -LiteralPath $UnmatchedCsv -NoTypeInformation -Encoding UTF8
} else {
    '所有文件均已匹配阶段。' | Set-Content -LiteralPath $UnmatchedCsv -Encoding UTF8
}

$stageSummary = @()
foreach ($folderName in $targetFolders) {
    $stageCount = @($plan | Where-Object { $_.'阶段' -eq $folderName }).Count
    $stageSummary += "${folderName}：$stageCount"
}

@(
    "执行时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "源目录：$Source"
    "阶段清单：$ListPath"
    "目标目录：$Destination"
    "源文件数：$($files.Count)"
    "计划分类数：$($plan.Count)"
    "成功复制：$copied"
    "复制失败：$failed"
    "路径过长跳过：$skipped"
    "未匹配/歧义：$($unmatched.Count)"
    "各阶段计划数：$($stageSummary -join '；')"
    "逐文件结果：$PreviewCsv"
    "未匹配清单：$UnmatchedCsv"
) | Set-Content -LiteralPath $SummaryLog -Encoding UTF8

Write-Host ''
Write-Host "分类完成：成功 $copied；失败 $failed；跳过 $skipped；未匹配/歧义 $($unmatched.Count)。" -ForegroundColor Green
Write-Host "分类目录：$Destination"
Write-Host "逐文件结果：$PreviewCsv"
Write-Host "汇总日志：$SummaryLog"

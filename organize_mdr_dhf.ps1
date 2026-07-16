<#
.SYNOPSIS
    DHF 项目管理文件整理脚本（v5 修正版）。
    改用 robocopy 扫描+进程级看门狗、staging 目录复制机制，解决长路径、
    ShareCache 离线文件阻塞、单文件复制卡死等问题。
.DESCRIPTION
    核心改进：
      - 扫描阶段：robocopy /L /S /FP 列文件清单（带超时看门狗），避免 Get-ChildItem -Recurse 被离线文件阻塞
      - 复制阶段：每文件先入 staging（GUID 子目录），再原子重命名到目标，防半成品
      - 新增诊断模式 -DiagnosticOnly：只跑 robocopy 扫描+输出清单，不匹配、不复制
      - 新增心跳/超时/taskkill 精确终止
      - 多日志文件：预览CSV、未找到CSV、扫描日志、复制日志、运行汇总
      - 路径长度预校验（目标路径 ≤ 259 字符，预留 1 字符余量）
.PARAMETER SourceRoot
    项目管理文件根目录。
.PARAMETER GoldRoot
    Gold Standard（MDR 技术文档）根目录。
.PARAMETER ListPath
    编号清单（TSV）。默认为脚本同目录的"项目管理文件编号.txt"。
.PARAMETER Destination
    输出目录。默认为桌面下的 DHF 文件夹。
.PARAMETER Execute
    实际复制。未指定时仅预览。
.PARAMETER DiagnosticOnly
    诊断模式：只执行 robocopy 扫描并输出扫描清单，不做编号匹配、不复制。
    用于排查"扫描卡死""路径过长""离线文件"等问题。
.PARAMETER ScanTimeoutSeconds
    robocopy 扫描单个根目录的超时秒数。默认 600（10 分钟）。
.PARAMETER CopyTimeoutSeconds
    robocopy 复制单个文件的超时秒数。默认 180（3 分钟），大 PDF 可适当调大。
.PARAMETER MaxPathLength
    目标路径最大长度（含文件名）。默认 259（Windows MAX_PATH - 1）。
.EXAMPLE
    # 仅预览
    .\organize_mdr_dhf.ps1
.EXAMPLE
    # 诊断模式（只扫描，排查卡死）
    .\organize_mdr_dhf.ps1 -DiagnosticOnly
.EXAMPLE
    # 实际复制
    .\organize_mdr_dhf.ps1 -Execute
.EXAMPLE
    # 自定义超时
    .\organize_mdr_dhf.ps1 -Execute -ScanTimeoutSeconds 900 -CopyTimeoutSeconds 300
#>
[CmdletBinding()]
param(
    [string]$SourceRoot = 'D:\ShareCache\LS_SH(上海生命科技)\1、研发部\2、CGM\CGM-欧盟注册-项目管理文件',
    [string]$GoldRoot = 'D:\ShareCache\LS_SH(上海生命科技)\1、研发部\2、CGM\CGM-MDR注册\Technical Documentation-final',
    [string]$ListPath = '',
    [string]$Destination = '',
    [switch]$Execute,
    [switch]$DiagnosticOnly,
    [int]$ScanTimeoutSeconds = 600,
    [int]$CopyTimeoutSeconds = 180,
    [int]$MaxPathLength = 259
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

# 时间戳后缀，避免覆盖历史日志；诊断模式单独命名以便区分
$TimeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if ($DiagnosticOnly) {
    $runTag = "diag_$TimeStamp"
} else {
    $runTag = $TimeStamp
}
$PreviewCsv   = Join-Path $ScriptDir ("DHF_整理预览_$runTag.csv")
$MissingCsv   = Join-Path $ScriptDir ("DHF_未找到编号_$runTag.csv")
$ScanLog      = Join-Path $ScriptDir ("DHF_扫描日志_$runTag.txt")
$CopyLog      = Join-Path $ScriptDir ("DHF_复制日志_$runTag.txt")
$SummaryLog   = Join-Path $ScriptDir ("DHF_运行汇总_$runTag.txt")
$StagingRoot  = Join-Path $ScriptDir ("DHF_staging_$runTag")
$SourceScanList  = Join-Path $ScriptDir ("DHF_源扫描清单_$runTag.txt")
$GoldScanList    = Join-Path $ScriptDir ("DHF_Gold扫描清单_$runTag.txt")

# ============================================================
# 工具函数
# ============================================================

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO', [string]$Path)
    $line = "[$([DateTime]::Now.ToString('HH:mm:ss'))] [$Level] $Message"
    Write-Host $line
    if ($Path) { Add-Content -LiteralPath $Path -Value $line -Encoding UTF8 }
}

function Assert-Path {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        throw "$Label 不存在：$Path"
    }
}

function Get-SafeTargetPath {
    param([string]$Folder, [string]$Name)
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

function Test-DocumentNumber {
    param([string]$FileName, [string]$Number, [bool]$IsPrefix)
    $escaped = [Regex]::Escape($Number)
    if ($IsPrefix) {
        # 系列编号（以 ~ 结尾）：编号后可跟任意字符（含 -数字）
        return $FileName -match "(?i)(?<![A-Za-z0-9])$escaped"
    }
    # 普通编号：编号后不能跟字母/数字，也不能跟 -数字（避免 S0023-9008 匹到 S0023-9008-01）
    return $FileName -match "(?i)(?<![A-Za-z0-9])$escaped(?![A-Za-z0-9]|-\d)"
}

<#
    通过 robocopy /L 列出目录下所有文件（不实际复制）。
    robocopy 对长路径、离线文件、ACL 异常有更强鲁棒性，且 /R:0 /W:0 避免卡死重试。
    返回文件完整路径数组。
    用 Start-Process -PassThru + 看门狗循环实现超时，超时后 taskkill /T /F 精确终止该进程树。
#>
function Invoke-RobocopyScan {
    param(
        [string]$Root,
        [string]$OutList,
        [string]$Label,
        [int]$TimeoutSeconds
    )
    Write-Log "开始扫描 $Label ：$Root" 'INFO' $ScanLog
    # /L  仅列出（不复制）
    # /S  含子目录（空目录不列）
    # /FP 显示完整路径
    # /NJH /NJS 去掉 robocopy 自带的 header/summary
    # /NDL 不列目录（只列文件）
    # /NC /NP 不显示类码和进度
    # /XJ 跳过联接点（junction），防止循环
    # /R:0 /W:0 不重试
    # /UNILOG:文件 以 Unicode 写日志（兼容中文路径）
    # 注意：$args 是 PS 自动变量，不能赋值，故用 $rcArgs
    $dummyTarget = Join-Path $([IO.Path]::GetTempPath()) 'DHF_DUMMY_TARGET'
    $rcArgs = @('/L', '/S', '/FP', '/NJH', '/NJS', '/NDL', '/NC', '/NP', '/XJ', '/R:0', '/W:0',
        "/UNILOG:`"$OutList`"", "`"$Root`"", "`"$dummyTarget`"")
    Write-Log "robocopy 参数：robocopy $($rcArgs -join ' ')" 'DEBUG' $ScanLog

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'robocopy.exe'
    # ProcessStartInfo.ArgumentList 在 .NET 5+ 才有；PS5.1 用 .NET Framework 4.x，
    # 这里仍用 Arguments 字符串，但参数已用引号包裹，路径中含空格/中文均可。
    $psi.Arguments = ($rcArgs -join ' ')
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $exitCode = $null
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $heartbeat = 0
    try {
        while (-not $proc.HasExited) {
            if ($watch.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
                Write-Log "扫描超时（${TimeoutSeconds}s），终止 robocopy 进程树 PID=$($proc.Id)" 'WARN' $ScanLog
                try {
                    # /T 终止整个进程树；/F 强制
                    $killProc = Start-Process -FilePath 'taskkill.exe' -ArgumentList "/PID $($proc.Id) /T /F" -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
                    Write-Log "taskkill 退出码：$($killProc.ExitCode)" 'DEBUG' $ScanLog
                } catch {
                    Write-Log "taskkill 异常：$($_.Exception.Message)" 'WARN' $ScanLog
                    try { $proc.Kill() } catch {}
                }
                Start-Sleep -Seconds 2
                if (-not $proc.HasExited) {
                    try { $proc.Kill() } catch {}
                }
                return $null
            }
            Start-Sleep -Seconds 5
            $heartbeat++
            # 每 30 秒输出一次心跳
            if ($heartbeat % 6 -eq 0) {
                $elapsed = [int]$watch.Elapsed.TotalSeconds
                Write-Log "扫描中… 已用 ${elapsed}s / ${TimeoutSeconds}s" 'INFO' $ScanLog
            }
        }
        $exitCode = $proc.ExitCode
    } finally {
        $watch.Stop()
        try { $proc.Dispose() } catch {}
    }

    # robocopy 退出码：0=无变化/无文件，1=成功列出文件，>=8 为错误
    # 注：/L 模式下 0 也可能是空目录，不算错误
    if ($exitCode -ge 8) {
        Write-Log "robocopy 扫描 $Label 失败，退出码=$exitCode（>=8 表示错误）" 'ERROR' $ScanLog
        return $null
    }
    Write-Log "robocopy 扫描 $Label 完成，退出码=$exitCode，用时 $([int]$watch.Elapsed.TotalSeconds)s" 'INFO' $ScanLog

    # 解析 UNILOG 文件，提取带完整路径的文件行
    # robocopy /FP 输出形如：
    #        新文件   D:\ShareCache\...\xxx.docx
    # 行首有 tab + 状态 + tab + 完整路径
    if (-not (Test-Path -LiteralPath $OutList)) {
        Write-Log "扫描日志未生成：$OutList" 'ERROR' $ScanLog
        return @()
    }
    # UNILOG 是 UTF-16 LE；PS5.1 的 Get-Content -Encoding Unicode 可正确读取
    $lines = Get-Content -LiteralPath $OutList -Encoding Unicode -ErrorAction SilentlyContinue
    $files = @()
    foreach ($line in $lines) {
        $t = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        # robocopy 文件行特征：以 tab 开头，含状态字（新文件/相同/较新…），后跟 tab 和路径
        # /FP 模式下路径必含盘符或 UNC 前缀
        $m = [Regex]::Match($t, '^[^\t]*\t+(.*)$')
        if ($m.Success) {
            $p = $m.Groups[1].Value.Trim()
            if ($p -and ($p -match '^[A-Za-z]:\\' -or $p -match '^\\\\')) {
                $files += $p
            }
        } else {
            # 兜底：整行就是路径
            if ($t -match '^[A-Za-z]:\\' -or $t -match '^\\\\') {
                $files += $t
            }
        }
    }
    Write-Log "$Label 解析到文件数：$($files.Count)" 'INFO' $ScanLog
    return $files
}

<#
    通过 robocopy 复制单个文件到 staging 目录。
    原子提交：复制成功后，staging 文件 Rename 到最终目标。
    超时则 taskkill 终止，确保不卡死。
#>
function Copy-FileWithWatchdog {
    param(
        [string]$Source,
        [string]$StagingDir,
        [int]$TimeoutSeconds
    )
    if (-not (Test-Path -LiteralPath $StagingDir)) {
        New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null
    }

    # robocopy 语法：robocopy <源目录> <目标目录> [<文件>...]
    # 复制单文件时需拆分为 源目录 + 目标目录 + 文件名
    $srcDir = Split-Path -Parent $Source
    $srcName = Split-Path -Leaf $Source
    $stagingDir = $StagingDir
    $rcArgs = @("`"$srcDir`"", "`"$stagingDir`"", "`"$srcName`"", '/R:0', '/W:0', '/NJH', '/NJS', '/NDL', '/NP', '/NS', '/NC', "/UNILOG:`"$StagingFile.log`"")
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'robocopy.exe'
    $psi.Arguments = ($rcArgs -join ' ')
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        while (-not $proc.HasExited) {
            if ($watch.Elapsed.TotalSeconds -gt $TimeoutSeconds) {
                try { Start-Process -FilePath 'taskkill.exe' -ArgumentList "/PID $($proc.Id) /T /F" -Wait -NoNewWindow -ErrorAction SilentlyContinue | Out-Null } catch {}
                Start-Sleep -Seconds 1
                if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
                return @{ Success = $false; Error = "复制超时（${TimeoutSeconds}s）" }
            }
            Start-Sleep -Milliseconds 500
        }
        $code = $proc.ExitCode
    } finally {
        $watch.Stop()
        try { $proc.Dispose() } catch {}
    }

    # robocopy 0=无操作,1=复制成功,2=多余文件,3=1+2；>=8 错误
    if ($code -ge 8) {
        return @{ Success = $false; Error = "robocopy 退出码 $code"; StagingFile = '' }
    }
    # robocopy 把文件复制到 $stagingDir，文件名保持原名 $srcName
    $actualStagingFile = Join-Path $stagingDir $srcName
    if (-not (Test-Path -LiteralPath $actualStagingFile)) {
        return @{ Success = $false; Error = "robocopy 完成（码 $code）但 staging 文件不存在：$actualStagingFile"; StagingFile = '' }
    }
    return @{ Success = $true; Error = ''; StagingFile = $actualStagingFile }
}

# ============================================================
# 参数校验
# ============================================================
Assert-Path $SourceRoot '项目管理文件目录'
Assert-Path $GoldRoot 'Gold Standard 目录'

# ============================================================
# 阶段 0：诊断模式
# ============================================================
if ($DiagnosticOnly) {
    Write-Host ''
    Write-Host '=== 诊断模式：仅执行 robocopy 扫描，不匹配、不复制 ===' -ForegroundColor Magenta
    Write-Host "源目录：$SourceRoot"
    Write-Host "Gold目录：$GoldRoot"
    Write-Host "扫描超时：${ScanTimeoutSeconds}s"
    Write-Host "扫描日志：$ScanLog"
    Write-Host ''

    $srcFiles = Invoke-RobocopyScan -Root $SourceRoot -OutList $SourceScanList -Label '项目管理目录' -TimeoutSeconds $ScanTimeoutSeconds
    $goldFiles = Invoke-RobocopyScan -Root $GoldRoot -OutList $GoldScanList -Label 'Gold Standard 目录' -TimeoutSeconds $ScanTimeoutSeconds

    Write-Host ''
    Write-Host '=== 诊断结果 ===' -ForegroundColor Cyan
    Write-Host "源目录文件数：$(@($srcFiles).Count)"
    Write-Host "Gold目录文件数：$(@($goldFiles).Count)"
    Write-Host "源扫描清单：$SourceScanList"
    Write-Host "Gold扫描清单：$GoldScanList"
    Write-Host "扫描日志：$ScanLog"
    Write-Host ''
    Write-Host '诊断完成。请检查上述清单与日志，确认扫描无卡死后再运行预览或 -Execute。' -ForegroundColor Green
    return
}

# ============================================================
# 阶段 1：扫描
# ============================================================
Assert-Path $ListPath '项目管理文件编号清单'
Write-Host '正在通过 robocopy 扫描两个目录（带超时看门狗）……' -ForegroundColor Cyan
$srcFiles = Invoke-RobocopyScan -Root $SourceRoot -OutList $SourceScanList -Label '项目管理目录' -TimeoutSeconds $ScanTimeoutSeconds
$goldFiles = Invoke-RobocopyScan -Root $GoldRoot -OutList $GoldScanList -Label 'Gold Standard 目录' -TimeoutSeconds $ScanTimeoutSeconds

if ($null -eq $srcFiles -or $null -eq $goldFiles) {
    Write-Log '扫描阶段超时或失败，已终止。请检查扫描日志，或增大 -ScanTimeoutSeconds，或先用 -DiagnosticOnly 排查。' 'ERROR' $ScanLog
    throw '扫描阶段失败，详见扫描日志。'
}

# 将路径转为轻量对象（仅保留 Name/FullName/Extension），便于后续匹配
function ConvertTo-FileInfoLite {
    param([string[]]$Paths)
    $list = @()
    foreach ($p in $Paths) {
        $name = [IO.Path]::GetFileName($p)
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -like '~$*') { continue }  # 跳过 Office 临时锁文件
        $list += [pscustomobject]@{
            Name = $name
            FullName = $p
            Extension = [IO.Path]::GetExtension($p)
        }
    }
    return $list
}
$sourceFiles = ConvertTo-FileInfoLite $srcFiles
$goldFiles = ConvertTo-FileInfoLite $goldFiles
Write-Log "扫描完成。源文件 $($sourceFiles.Count)；Gold 文件 $($goldFiles.Count)" 'INFO' $ScanLog

# ============================================================
# 阶段 2：读取清单 + 编号匹配
# ============================================================
$rows = Import-Csv -LiteralPath $ListPath -Delimiter "`t" -Encoding UTF8
if (-not $rows) { throw '编号清单为空。' }
if (-not ($rows[0].PSObject.Properties.Name -contains '文件编号')) {
    throw '清单中未找到"文件编号"列。'
}

$plan = @()
$missing = @()
$seenSelection = @{}
$pathWarn = 0

foreach ($row in $rows) {
    $rawNumber = ([string]$row.'文件编号').Trim()
    if ([string]::IsNullOrWhiteSpace($rawNumber)) { continue }
    $isPrefix = $rawNumber.EndsWith('~')
    $number = $rawNumber.TrimEnd([char]'~').Trim()

    $src = @($sourceFiles | Where-Object { Test-DocumentNumber $_.Name $number $isPrefix })
    $gold = @($goldFiles | Where-Object { Test-DocumentNumber $_.Name $number $isPrefix })
    $extensionList = @()
    foreach ($f in $src) { $extensionList += $f.Extension.ToLowerInvariant() }
    foreach ($f in $gold) { $extensionList += $f.Extension.ToLowerInvariant() }
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

            $targetPath = (Join-Path $Destination $file.Name)
            $initialResult = if ($Execute) { '待复制' } else { '预览' }

            # 目标路径长度预校验
            $pathFlag = ''
            if ($targetPath.Length -gt $MaxPathLength) {
                $pathFlag = "目标路径过长($($targetPath.Length)>${MaxPathLength})"
                $pathWarn++
                if ($Execute) { $initialResult = $pathFlag }
            }

            $plan += [pscustomobject][ordered]@{
                '文件编号' = $rawNumber
                '清单文件名称' = $row.'文件名称'
                '选用来源' = $origin
                '文件名' = $file.Name
                '原路径' = $file.FullName
                '扩展名' = $file.Extension
                '被Gold替代的源文件数' = $suppressed
                '原路径长度' = $file.FullName.Length
                '目标路径' = $targetPath
                '目标路径长度' = $targetPath.Length
                '路径校验' = $pathFlag
                '执行结果' = $initialResult
            }
        }
    }
}

# 导出预览/未找到
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

if ($pathWarn -gt 0) {
    Write-Log "警告：有 $pathWarn 个文件目标路径超过 ${MaxPathLength} 字符，已标记路径校验列。执行模式将跳过这些文件。" 'WARN' $ScanLog
}

if (-not $Execute) {
    Write-Host ''
    Write-Host '仅完成预览，未复制任何文件。' -ForegroundColor Yellow
    Write-Host "候选文件：$($plan.Count)；未找到编号：$($missing.Count)；路径过长：$pathWarn"
    Write-Host "预览表：$PreviewCsv"
    Write-Host "未找到：$MissingCsv"
    Write-Host '确认预览后，使用同一命令并追加 -Execute。' -ForegroundColor Green
    return
}

# ============================================================
# 阶段 3：复制（staging + 原子提交）
# ============================================================
New-Item -ItemType Directory -Path $Destination -Force | Out-Null
if (Test-Path -LiteralPath $StagingRoot) { Remove-Item -LiteralPath $StagingRoot -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Path $StagingRoot -Force | Out-Null

$copied = 0
$copyErrors = 0
$skipped = 0
$total = $plan.Count
$idx = 0

foreach ($item in $plan) {
    $idx++
    # 跳过路径过长的文件
    if ($item.'路径校验' -ne '') {
        $item.'执行结果' = "跳过：$($item.'路径校验')"
        $skipped++
        Write-Log "[$idx/$total] 跳过（路径过长）：$($item.'文件名')" 'WARN' $CopyLog
        continue
    }

    $target = Get-SafeTargetPath $Destination $item.'文件名'
    # 每文件一个独立的 staging 子目录（GUID），避免同名文件冲突
    $stagingSubDir = Join-Path $StagingRoot ([Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stagingSubDir -Force | Out-Null
    Write-Log "[$idx/$total] 复制：$($item.'文件名') -> staging($stagingSubDir) -> $target" 'INFO' $CopyLog

    $r = Copy-FileWithWatchdog -Source $item.'原路径' -StagingFile $stagingSubDir -TimeoutSeconds $CopyTimeoutSeconds
    if (-not $r.Success) {
        $item.'执行结果' = "失败：$($r.Error)"
        $copyErrors++
        Write-Log "[$idx/$total] 失败：$($r.Error) | 源：$($item.'原路径')" 'ERROR' $CopyLog
        if (Test-Path -LiteralPath $stagingSubDir) { Remove-Item -LiteralPath $stagingSubDir -Recurse -Force -ErrorAction SilentlyContinue }
        continue
    }

    # 原子提交：staging 实际文件 -> 目标（同卷 Rename 为原子操作）
    $actualStagingFile = $r.StagingFile
    try {
        # 若目标已存在（Get-SafeTargetPath 已避免，但并发/历史残留时兜底）
        if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force -ErrorAction Stop }
        Move-Item -LiteralPath $actualStagingFile -Destination $target -ErrorAction Stop
        $item.'目标路径' = $target
        $item.'执行结果' = '已复制'
        $copied++
        Write-Log "[$idx/$total] 成功" 'INFO' $CopyLog
    } catch {
        $item.'执行结果' = "提交失败：$($_.Exception.Message)"
        $copyErrors++
        Write-Log "[$idx/$total] 提交失败：$($_.Exception.Message)" 'ERROR' $CopyLog
    } finally {
        # 无论成功失败，清理该文件的 staging 子目录（含 robocopy 日志等残留）
        if (Test-Path -LiteralPath $stagingSubDir) { Remove-Item -LiteralPath $stagingSubDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# 清理 staging
if (Test-Path -LiteralPath $StagingRoot) {
    Remove-Item -LiteralPath $StagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# 重新导出预览（含执行结果）
$plan | Sort-Object '文件编号', '选用来源', '文件名' | Export-Csv -LiteralPath $PreviewCsv -NoTypeInformation -Encoding UTF8

# 汇总日志
@(
    "执行时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "运行标签：$runTag"
    "项目管理目录：$SourceRoot"
    "Gold目录：$GoldRoot"
    "输出目录：$Destination"
    "扫描超时：${ScanTimeoutSeconds}s / 复制超时：${CopyTimeoutSeconds}s / 最大路径：${MaxPathLength}"
    "清单编号数：$($rows.Count)"
    "候选文件数：$($plan.Count)"
    "成功复制：$copied"
    "复制失败：$copyErrors"
    "路径过长跳过：$skipped"
    "未找到编号：$($missing.Count)"
    "预览表：$PreviewCsv"
    "未找到清单：$MissingCsv"
    "扫描日志：$ScanLog"
    "复制日志：$CopyLog"
) | Set-Content -LiteralPath $SummaryLog -Encoding UTF8

Write-Host ''
Write-Host "完成：成功 $copied；失败 $copyErrors；路径过长跳过 $skipped；未找到编号 $($missing.Count)。" -ForegroundColor Green
Write-Host "DHF目录：$Destination"
Write-Host "汇总日志：$SummaryLog"
Write-Host "复制日志：$CopyLog"

<#
xae_cli.ps1 — TwinCAT XAE 部署 CLI（Windows PowerShell 5.1）

目标：把"打开方案 → 编译 → 生成 boot 工程 → 目标切 Config → 激活配置 → 重启进 RUN → 轮询 851 → 只读校验"
固化为可复用、可判定、可留证的 CLI；示例按 TwinCAT 3 solution/tsproj 结构编写，验证符号需按你的 PLC 工程适配。

必须用 Windows PowerShell 5.1 运行（PS7 没有 [Marshal]::GetActiveObject）：
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File src\xae_cli.ps1 <cmd> -Solution <sln> [-NetId <id>] [-Json] [-Timeout 180] [-LogPath <file>]

子命令：
  status    只读：DTE/方案/工程就绪、tsproj TargetNetId、TMC 时间戳、目标 :10000/:851 状态、XAE 窗口（含弹窗）
  fingerprint  只读磁盘：solution 指纹（绝对路径/仓库/tsproj GUID/plcproj/SHA256），不碰 XAE/目标
  compile   编译（Build.RebuildSolution，TMC 变化 + LastBuildInfo + 错误列表三判据）
  boot      ITcPlcProject.CompileProject + BootProjectAutostart=TRUE + GenerateBootProject
  config    目标切 Config(15)（ADS 10000 WriteControl RECONFIG，轮询确认；本 CX 实测可用）
  activate  SetTargetNetId + (AutoSaveProjectToTarget) + ActivateConfiguration（COM 忙重试；模态弹窗判阻塞）
  run       StartRestartTwinCAT（COM 忙重试）
  wait      轮询 :10000→RUN(5) 再 :851→RUN(5)（超时给明确失败点）
  verify    只读 851：read_state + 关键符号 + 心跳递增进（成功判据=851 RUN 且符号可读且心跳递增）
  all       依次 status→compile→boot→config→activate→run→wait→verify，任一失败即停

退出码（全命令统一）：
  0  成功
  2  参数/文件错误（-Solution 不存在、未知子命令）
  3  XAE/DTE 不可用（无法附加/启动、COM 环境异常）
  4  方案未打开 / 工程未就绪（含"打开的方案不是指定方案"）
  5  编译失败（TMC 未刷新且有错误 / LastBuildInfo≠0 / 错误列表非空）
  6  boot 工程生成失败（CompileProject/GenerateBootProject 报错或 TMC 缺失）
  7  目标切 Config 失败（ADS 10000 不可达 / 写控制被拒 / 超时未到 15）
  8  激活调用被拒/异常（ActivateConfiguration 失败或工程对象异常）
  9  许可/模态弹窗阻塞（COM 持续 RPC_E_CALL_REJECTED/E_ABORT，窗口文本含许可/试用/错误/警告）
  10 目标 :10000 超时未到 RUN(5)
  11 PLC :851 未注册/超时未到 RUN(5)（error 6 属此类的典型表现）
  12 verify 失败（851 可连但符号不可读/心跳不递增/关键值不符）
  13 串行约束被占用（检测到其它项目 XAE/部署工具正在活动，拒绝操作）
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("status", "fingerprint", "compile", "boot", "config", "activate", "run", "wait", "verify", "all")]
    [string]$Command = "status",
    [string]$Solution = "",
    [string]$NetId = "",
    [switch]$Json,
    [int]$Timeout = 180,
    [string]$LogPath = "",
    [switch]$NoLaunch,
    [string]$ExpectedRepoRoot = "",
    [string]$ExpectedProjectName = "",
    [int]$ExpectedAxes = 0,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$script:Result = [ordered]@{ command = $Command; startedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss"); steps = @(); ok = $false; exitCode = 0 }
$script:StepRows = New-Object System.Collections.ArrayList
$script:AdsAsm = $null
$script:InteropAsm = $null

function Say([string]$m) {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m
    Write-Host $line
    if ($LogPath) {
        try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
    }
}
function Step([string]$name, [bool]$ok, [string]$detail) {
    $row = [ordered]@{ name = $name; ok = $ok; detail = $detail }
    [void]$script:StepRows.Add($row)
    Say ("  {0} {1}: {2}" -f $(if ($ok) { "[OK]" } else { "[FAIL]" }), $name, $detail)
}
function Finish([int]$code, [string]$why) {
    $script:Result.ok = ($code -eq 0)
    $script:Result.exitCode = $code
    $script:Result.steps = @($script:StepRows)
    $script:Result.endedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    if ($code -ne 0) { $script:Result.error = $why; Say ("EXIT {0}: {1}" -f $code, $why) } else { Say "EXIT 0: OK" }
    if ($Json) { $script:Result | ConvertTo-Json -Depth 6 | Write-Output }
    exit $code
}

# ───────────────────────────── COM MessageFilter（RPC_E_CALL_REJECTED 自动重试）─────────────────────────────
if (-not ("XaeComMessageFilter" -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}
public class XaeComMessageFilter : IOleMessageFilter {
    [DllImport("Ole32.dll")] static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new XaeComMessageFilter(), out old); }
    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) { return 100; }
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 1; }
}
'@
}
try { [XaeComMessageFilter]::Register() } catch { Say ("WARN MessageFilter 注册失败: " + $_.Exception.Message) }

function Invoke-ComRetry([scriptblock]$sb, [int]$n = 30, [int]$delay = 2, [string]$what = "COM") {
    for ($i = 1; $i -le $n; $i++) {
        try { return (& $sb) } catch {
            $msg = $_.Exception.Message
            Say ("  {0} 重试 {1}/{2}: {3}" -f $what, $i, $n, $msg.Substring(0, [Math]::Min(160, $msg.Length)))
            Start-Sleep -Seconds $delay
        }
    }
    throw ("{0} 重试 {1} 次仍失败" -f $what, $n)
}

# ───────────────────────────── UIA：枚举 XAE 顶层窗口/弹窗（模态阻塞证据）─────────────────────────────
function Get-XaeWindows {
    $rows = New-Object System.Collections.ArrayList
    try {
        Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
        $xe = Get-Process TcXaeShell -ErrorAction SilentlyContinue
        if (-not $xe) { return @() }
        $pids = @($xe | ForEach-Object { $_.Id })
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $children = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($c in $children) {
            try {
                $procId = $c.Current.ProcessId
                if ($procIds -contains $procId) {
                    [void]$rows.Add([ordered]@{ pid = $procId; name = $c.Current.Name; controlType = $c.Current.ControlType.ProgrammaticName })
                }
            } catch { }
        }
    } catch { Say ("WARN UIA 枚举失败: " + $_.Exception.Message) }
    return @($rows)
}
function Test-BlockedWindows([object[]]$windows) {
    $kw = @("许可", "试用", "License", "license", "Activate Configuration", "激活配置", "Error", "错误", "警告", "Warning", "TwinCAT Project")
    $hits = @()
    foreach ($w in $windows) {
        foreach ($k in $kw) { if ($w.name -and $w.name.Contains($k)) { $hits += $w.name; break } }
    }
    return @($hits | Select-Object -Unique)
}

# ───────────────────────────── 串行约束：侦察其它 XAE/部署活动 ─────────────────────────────
function Get-OtherProjectActivity([string]$sln) {
    $hits = New-Object System.Collections.ArrayList
    $procs = Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='python.exe' OR Name='node.exe' OR Name='TcXaeShell.exe' OR Name='devenv.exe'" -ErrorAction SilentlyContinue
    $mine = $sln.ToLowerInvariant()
    # 排除自身/父进程链，避免把自己的命令行误判为“其它项目活动”。
    $selfPids = @($PID)
    $curPid = $PID
    for ($i = 0; $i -lt 6; $i++) {
        $p = Get-CimInstance Win32_Process -Filter "ProcessId=$curPid" -ErrorAction SilentlyContinue
        if (-not $p -or -not $p.ParentProcessId) { break }
        $curPid = [int]$p.ParentProcessId
        $selfPids += $curPid
    }
    $selfScript = ([string]$PSCommandPath).ToLowerInvariant()
    foreach ($pr in $procs) {
        if ($selfPids -contains [int]$pr.ProcessId) { continue }
        $cmd = [string]$pr.CommandLine
        if (-not $cmd) { continue }
        $lc = $cmd.ToLowerInvariant()
        if ($selfScript -and $lc.Contains($selfScript)) { continue }
        $isOther = $false
        if ($lc -match 'tcxae\.ps1') { $isOther = $true }
        foreach ($m in [regex]::Matches($lc, '[a-z]:\\[^"]*?\.sln')) {
            if ($m.Groups[0].Value -ne $mine) { $isOther = $true }
        }
        if ($isOther) { [void]$hits.Add([ordered]@{ pid = $pr.ProcessId; name = $pr.Name; cmd = $cmd.Substring(0, [Math]::Min(240, $cmd.Length)) }) }
    }
    return @($hits)
}

# ───────────────────────────── ADS（v4.3 TcAdsClient，只用于读状态/切模式/只读符号）─────────────────────────────
function Load-AdsV4 {
    if ($script:AdsAsm) { return }
    $cands = @(
        "C:\Program Files (x86)\Beckhoff\TwinCAT\3.1\Components\Base\v150\TwinCAT.Ads.dll",
        "C:\Program Files (x86)\Beckhoff\TwinCAT\3.1\Components\Base\TwinCAT.Ads.dll"
    )
    foreach ($p in $cands) {
        if (Test-Path $p) {
            $script:AdsAsm = [Reflection.Assembly]::LoadFrom($p)
            Say ("  ADS 程序集: " + $p + " (" + $script:AdsAsm.GetName().Version + ")")
            return
        }
    }
    throw "未找到 TwinCAT.Ads.dll（v4.3）"
}
function New-AdsClient([string]$netid, [int]$port) {
    Load-AdsV4
    $t = $script:AdsAsm.GetType("TwinCAT.Ads.TcAdsClient")
    $c = [Activator]::CreateInstance($t)
    $c.Timeout = 5000
    $c.Connect($netid, $port)
    return $c
}
function Get-AdsStateInt([object]$client) {
    $si = $client.ReadState()
    return [int]$si.AdsState
}
function Read-AdsStateSafe([string]$netid, [int]$port) {
    $c = $null
    try {
        $c = New-AdsClient $netid $port
        $si = $c.ReadState()
        return [ordered]@{ ok = $true; state = [int]$si.AdsState; device = [int]$si.DeviceState; error = $null }
    } catch {
        $msg = $_.Exception.Message
        return [ordered]@{ ok = $false; state = $null; device = $null; error = $msg.Substring(0, [Math]::Min(200, $msg.Length)) }
    } finally { if ($c) { try { $c.Dispose() } catch { } } }
}
function Set-AdsModeWait([string]$netid, [int]$wantedState, [int]$timeoutSec) {
    Load-AdsV4
    $t = $script:AdsAsm.GetType("TwinCAT.Ads.TcAdsClient")
    $adsStateType = $script:AdsAsm.GetType("TwinCAT.Ads.AdsState")
    $target = $null
    foreach ($n in @("Reconfig", "Reset", "Run", "Stop", "Config")) {
        if ((($wantedState -eq 15) -and ($n -eq "Reconfig")) -or (($wantedState -eq 5) -and ($n -eq "Reset"))) {
            $target = [Enum]::Parse($adsStateType, $n); break
        }
    }
    if ($null -eq $target) { throw ("不支持的目标状态 " + $wantedState) }
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    $last = $null
    $lastErr = ""
    while ((Get-Date) -lt $deadline) {
        try {
            $c = New-AdsClient $netid 10000
            try {
                $cur = Get-AdsStateInt $c
                $last = $cur
                if ($cur -ne $wantedState) {
                    $si = New-Object TwinCAT.Ads.StateInfo
                    $si.AdsState = $target
                    $si.DeviceState = 0
                    $c.WriteControl($si)
                    Say ("  10000 WriteControl(" + $target + ") 已发出（当前 " + $cur + "）")
                }
                for ($i = 0; $i -lt 20; $i++) {
                    Start-Sleep -Milliseconds 800
                    try { $cur = Get-AdsStateInt $c; $last = $cur } catch { continue }
                    if ($cur -eq $wantedState) { break }
                }
            } finally { try { $c.Dispose() } catch { } }
        } catch { $lastErr = $_.Exception.Message }
        if ($last -eq $wantedState) { return $true }
        Start-Sleep -Seconds 2
    }
    Say ("  切模式超时：最后状态=" + $last + " 错误=" + $lastErr)
    return $false
}
function Read-AdsSymbols([string]$netid, [int]$port, [object[]]$defs) {
    # defs: @(@{name=...; type="bool"/"udint"/"int"/"real"; }...)
    $out = [ordered]@{}
    $c = $null
    try {
        $c = New-AdsClient $netid $port
        foreach ($d in $defs) {
            $h = 0
            try {
                $h = $c.CreateVariableHandle($d.name)
                $dot = $d.type
                switch ($dot) {
                    "bool" { $out[$d.name] = [bool]$c.ReadAny($h, [bool]) }
                    "udint" { $out[$d.name] = [uint32]$c.ReadAny($h, [uint32]) }
                    "dint" { $out[$d.name] = [int32]$c.ReadAny($h, [int32]) }
                    "int" { $out[$d.name] = [int16]$c.ReadAny($h, [int16]) }
                    "real" { $out[$d.name] = [single]$c.ReadAny($h, [single]) }
                    default { $out[$d.name] = $c.ReadAny($h, [int16]) }
                }
            } catch {
                $out[$d.name] = "ERR: " + $_.Exception.Message.Substring(0, [Math]::Min(120, $_.Exception.Message.Length))
            } finally { if ($h -ne 0) { try { $c.ReleaseVariableHandle($h) } catch { } } }
        }
        return $out
    } finally { if ($c) { try { $c.Dispose() } catch { } } }
}

# ───────────────────────────── 工程解析 / DTE / 系统管理器 ─────────────────────────────
function Resolve-SolutionPath([string]$sln) {
    if (-not $sln) { return $null }
    try { $full = (Resolve-Path -LiteralPath $sln -ErrorAction Stop).Path } catch { return $null }
    if ($full -notmatch "\.sln$") { return $null }
    return $full
}
function Get-SolutionFingerprint([string]$sln) {
    # solution 指纹守卫（2026-11-17 误激活事故后新增）：
    # 多个 TwinCAT 仓库可能都有同名 solution，窗口标题/文件名相同 → 仅看标题无法区分。
    # 指纹 = 绝对路径 + tsproj ProjectGUID + plcproj 路径 + git 仓库根 + 仓库标记 + SHA256。
    $slnFull = (Resolve-Path -LiteralPath $sln -ErrorAction Stop).Path
    $dir = Split-Path $slnFull -Parent
    $tsproj = Get-ChildItem $dir -Filter *.tsproj -File -ErrorAction SilentlyContinue | Select-Object -First 1
    $guid = $null; $plcRel = $null; $tmcRel = $null
    if ($tsproj) {
        $txt = Get-Content $tsproj.FullName -Raw -Encoding UTF8
        $m = [regex]::Match($txt, 'ProjectGUID="([^"]+)"'); if ($m.Success) { $guid = $m.Groups[1].Value }
        $m2 = [regex]::Match($txt, 'PrjFilePath="([^"]+)"'); if ($m2.Success) { $plcRel = $m2.Groups[1].Value }
        $m3 = [regex]::Match($txt, 'TmcFilePath="([^"]+)"'); if ($m3.Success) { $tmcRel = $m3.Groups[1].Value }
    }
    $repo = $null
    $cands = @($dir, (Split-Path $dir -Parent), (Split-Path (Split-Path $dir -Parent) -Parent))
    foreach ($cand in $cands) {
        try {
            $out = & git -C $cand rev-parse --show-toplevel 2>$null
            if ($LASTEXITCODE -eq 0 -and $out) { $repo = ([string]($out | Select-Object -First 1)).Trim(); break }
        } catch { }
    }
    $plcproj = $null
    if ($plcRel -and $tsproj) { $plcproj = Join-Path (Split-Path $tsproj.FullName -Parent) ($plcRel -replace '\\', '\') }
    $tmc = $null
    if ($tmcRel -and $tsproj) { $tmc = Join-Path (Split-Path $tsproj.FullName -Parent) ($tmcRel -replace '\\', '\') }
    $lc = $slnFull.ToLowerInvariant()
    $marker = "unknown"
    if ($repo) { $marker = (Split-Path $repo -Leaf).ToLowerInvariant() }
        $raw = @($slnFull.ToLowerInvariant(), ([string]$tsproj.FullName).ToLowerInvariant(), ([string]$guid).ToLowerInvariant(), ([string]$plcproj).ToLowerInvariant(), ([string]$repo).ToLowerInvariant(), $marker) -join "|"
    $sha = [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($raw))).Replace("-", "").ToLowerInvariant()
    return [ordered]@{
        marker = $marker
        repo = $repo
        sln = $slnFull
        tsproj = $(if ($tsproj) { $tsproj.FullName } else { $null })
        projectGuid = $guid
        plcproj = $plcproj
        tmc = $tmc
        sha256 = $sha
    }
}
function Assert-SolutionFingerprint([string]$sln) {
    # 只读磁盘校验：sln 存在、tsproj 存在、plcproj 存在、仓库标记可判定；失败返回错误文本
    try {
        $fp = Get-SolutionFingerprint $sln
    } catch { return ("指纹计算失败: " + $_.Exception.Message) }
    if (-not $fp.tsproj -or -not (Test-Path $fp.tsproj)) { return "tsproj 不存在（工程不完整）" }
    if (-not $fp.plcproj -or -not (Test-Path $fp.plcproj)) { return ("plcproj 不存在: " + $fp.plcproj) }
    if ($fp.marker -eq "unknown") { return ("无法判定仓库标记（sln 路径不在已知仓库内）: " + $fp.sln) }
    return $null
}

function Get-TsprojPath([string]$sln) {
    $dir = Split-Path $sln -Parent
    $ts = Get-ChildItem $dir -Filter *.tsproj -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ts) { return $ts.FullName }
    return $null
}
function Get-TsprojInfo([string]$sln) {
    $ts = Get-TsprojPath $sln
    $info = [ordered]@{ tsproj = $ts; targetNetId = $null; tmcPath = $null }
    if (-not $ts) { return $info }
    $txt = Get-Content $ts -Raw -Encoding UTF8
    $m = [regex]::Match($txt, 'TargetNetId="([^"]+)"')
    if ($m.Success) { $info.targetNetId = $m.Groups[1].Value }
    $m2 = [regex]::Match($txt, 'TmcFilePath="([^"]+)"')
    if ($m2.Success) {
        $info.tmcPath = Join-Path (Split-Path $ts -Parent) ($m2.Groups[1].Value -replace '\\', '\')
    } else {
        $guess = Get-ChildItem (Split-Path $sln -Parent) -Filter *.tmc -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($guess) { $info.tmcPath = $guess.FullName }
    }
    return $info
}
function Get-TmcStamp([string]$path) {
    if (-not $path -or -not (Test-Path $path)) { return [ordered]@{ exists = $false; bytes = 0; mtime = $null } }
    $f = Get-Item $path
    return [ordered]@{ exists = $true; bytes = $f.Length; mtime = $f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") }
}
function Get-InteropAsm {
    if ($script:InteropAsm) { return $script:InteropAsm }
    $interop = Join-Path $env:TEMP "interop-TCatSysManager.dll"
    $tlb = "C:\Program Files (x86)\Beckhoff\TwinCAT\3.1\Components\Base\TypeLib\3.1\TCatSysManager.tlb"
    if (-not (Test-Path $interop)) {
        $tlbimp = "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\TlbImp.exe"
        if (-not (Test-Path $tlbimp)) { throw "缺少 TlbImp.exe" }
        & $tlbimp $tlb /out:$interop /silent | Out-Null
    }
    if (-not (Test-Path $interop)) { throw "TlbImp 生成互操作程序集失败: $interop" }
    $script:InteropAsm = [Reflection.Assembly]::LoadFrom($interop)
    return $script:InteropAsm
}
function Test-DteUsable([object]$dte) {
    # 只读校验：DTO 代理可能指向已退出的进程（GetActiveObject 残留），属性会全部为空 → 必须判活
    if ($null -eq $dte) { return $false }
    try {
        if (-not $dte.Solution.IsOpen) { return $false }
        if ([string]::IsNullOrWhiteSpace([string]$dte.Solution.FullName)) { return $false }
        if ($dte.Solution.Projects.Count -lt 1) { return $false }
    } catch { return $false }
    return $true
}

function Get-Dte([string]$sln) {
    $dte = $null
    try { $dte = [Runtime.InteropServices.Marshal]::GetActiveObject("TcXaeShell.DTE.17.0") } catch { $dte = $null }
    if ($dte -and (Test-DteUsable $dte)) {
        Say ("  附加已运行 XAE（pid=" + ((Get-Process TcXaeShell -ErrorAction SilentlyContinue | Select-Object -First 1).Id) + "）")
        return $dte
    }
    if ($dte) { Say "  WARN GetActiveObject 返回的 XAE 代理不可用（方案未开/无工程/进程已退出），按未附加处理" }
    if ($NoLaunch) { return $null }
    $exes = @(
        "C:\Program Files\Beckhoff\TcXaeShell\Common7\IDE\TcXaeShell.exe",
        "C:\Program Files (x86)\Beckhoff\TcXaeShell\Common7\IDE\TcXaeShell.exe"
    )
    $exe = $exes | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $exe) { throw "未找到 TcXaeShell.exe" }
    Say ("  启动 XAE: " + $exe + " " + $sln)
    Start-Process -FilePath $exe -ArgumentList ('"' + $sln + '"') | Out-Null
    $deadline = (Get-Date).AddSeconds(180)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 3
        $dte = $null
        try { $dte = [Runtime.InteropServices.Marshal]::GetActiveObject("TcXaeShell.DTE.17.0") } catch { }
        if ($dte -and (Test-DteUsable $dte)) { return $dte }
    }
    return $null
}
function Assert-Solution([object]$dte, [string]$sln) {
    try {
        if (-not $dte.Solution.IsOpen) { return "方案未打开" }
        $cur = [string]$dte.Solution.FullName
        if ([string]::IsNullOrWhiteSpace($cur)) { return "方案路径为空（死代理）" }
        $curResolved = Resolve-Path -LiteralPath $cur -ErrorAction SilentlyContinue
        if (-not $curResolved) { return ("方案路径不可解析: " + $cur) }
        if ($curResolved.Path -ne $sln) { return ("打开的方案不是指定方案: " + $cur) }
        $fp = Get-SolutionFingerprint $sln
        if ($fp.tsproj) {
            $proj = $dte.Solution.Projects.Item(1)
            $pf = [string]$proj.FullName
            if ($pf) {
                $pfResolved = Resolve-Path -LiteralPath $pf -ErrorAction SilentlyContinue
                if (-not $pfResolved) { return ("工程路径不可解析: " + $pf) }
                if ($pfResolved.Path -ne $fp.tsproj) { return ("打开的工程不是期望 tsproj（跨仓库风险）: " + $pf) }
            }
        }
    } catch { return ("读取方案状态失败: " + $_.Exception.Message) }
    return $null
}
function Get-SysManager([object]$dte) {
    $obj = $null
    for ($i = 0; $i -lt 90; $i++) {
        try { $obj = $dte.Solution.Projects.Item(1).Object; if ($null -ne $obj) { break } } catch { }
        Start-Sleep -Seconds 2
    }
    if ($null -eq $obj) { return $null }
    $asm = Get-InteropAsm
    $sysType = $asm.GetType("interop-TCatSysManager.ITcSysManager")
    $punk = [Runtime.InteropServices.Marshal]::GetComInterfaceForObject($obj, $sysType)
    return [Runtime.InteropServices.Marshal]::GetTypedObjectForIUnknown($punk, $sysType)
}
function Get-PlcProjectNode([object]$sys) {
    $asm = Get-InteropAsm
    $tpc = $sys.LookupTreeItem("TIPC")
    $start = $tpc.Child(1)
    $iface = $asm.GetType("interop-TCatSysManager.ITcPlcProject")
    $pp = $null
    try { $pp = $start -as $iface } catch { }
    if ($null -eq $pp) {
        $punk = [Runtime.InteropServices.Marshal]::GetComInterfaceForObject($start, $iface)
        $pp = [Runtime.InteropServices.Marshal]::GetTypedObjectForIUnknown($punk, $iface)
    }
    return [ordered]@{ node = $start; plc = $pp }
}
function Read-DteErrors([object]$dte) {
    $items = @()
    try {
        $el = $dte.ToolWindows.ErrorList
        $n = $el.ErrorItems.Count
        for ($i = 1; $i -le $n; $i++) {
            $it = $el.ErrorItems.Item($i)
            $items += [ordered]@{
                description = [string]$it.Description
                file = [string]$it.FileName
                line = [int]$it.Line
            }
        }
    } catch { }
    return @($items)
}

# ───────────────────────────── 子命令实现 ─────────────────────────────
function Invoke-Fingerprint([object]$ctx) {
    Say "=== fingerprint（只读磁盘；不碰 XAE/目标）==="
    $bad = Assert-SolutionFingerprint $ctx.sln
    if ($bad) { Step "fingerprint" $false $bad; Finish 4 $bad }
    $fp = Get-SolutionFingerprint $ctx.sln
    foreach ($k in $fp.Keys) { Say ("  " + $k + " = " + $fp[$k]) }
    Step "fingerprint" $true ($fp.marker + " / sha256=" + $fp.sha256.Substring(0, 16) + "…")
    if ($Json) { $fp | ConvertTo-Json -Depth 4 | Write-Output }
    Finish 0 ""
}

# ───────────────────────────── 身份守卫（2026-11-17 同名 solution 误激活后强制）─────────────────────────────
function Get-OpenedXaeInfo {
    # 只读：返回实际打开的 XAE 方案/工程；绝不写、不激活
    $procs = @(Get-Process TcXaeShell -ErrorAction SilentlyContinue)
    if ($procs.Count -eq 0) { return [ordered]@{ state = "none"; detail = "无 TcXaeShell 进程"; sln = $null; tsproj = $null; processes = 0 } }
    try {
        $dte = [Runtime.InteropServices.Marshal]::GetActiveObject("TcXaeShell.DTE.17.0")
    } catch {
        return [ordered]@{ state = "com_unavailable"; detail = ("COM 附加失败（可能 elevated/完整性级别不同）: " + $_.Exception.Message); sln = $null; tsproj = $null; processes = $procs.Count }
    }
    if (-not (Test-DteUsable $dte)) {
        return [ordered]@{ state = "com_unavailable"; detail = "DTE 代理不可用（方案未开/工程未载/进程已退出）"; sln = $null; tsproj = $null; processes = $procs.Count }
    }
    $sln = $null; $ts = $null
    try { $sln = [string]$dte.Solution.FullName } catch { }
    try { $ts = [string]$dte.Solution.Projects.Item(1).FullName } catch { }
    return [ordered]@{ state = "ok"; detail = "COM 可读"; sln = $sln; tsproj = $ts; processes = $procs.Count }
}
function Assert-IdentityGuard([object]$ctx) {
    # 强制身份守卫：失败返回错误文本（调用方 exit 13），不触发任何下载/激活/切模式。
    $errs = New-Object System.Collections.ArrayList
    $detail = [ordered]@{
        expectedRepoRoot = $ctx.repoRoot
        expectedProjectName = $ctx.expectedProjectName
        solution = $ctx.sln
        requestedNetId = $ctx.netid
        dryRun = [bool]$ctx.dryRun
    }
    # 1) solution 必须存在且位于 ExpectedRepoRoot 下（防同名跨仓）
    if (-not (Test-Path -LiteralPath $ctx.sln)) { [void]$errs.Add("solution 不存在: " + $ctx.sln) }
    else {
        $slnCanon = (Resolve-Path -LiteralPath $ctx.sln).Path
        $rootCanon = $ctx.repoRoot
        $under = ($slnCanon.ToLowerInvariant().StartsWith($rootCanon.ToLowerInvariant() + [System.IO.Path]::DirectorySeparatorChar))
        if (-not $under) { [void]$errs.Add("solution 不在 ExpectedRepoRoot 下（跨仓库拒绝）: " + $slnCanon + " 不在 " + $rootCanon) }
        $detail.solutionCanonical = $slnCanon
    }
    # 2) sln 引用的 tsproj 存在且在仓库下
    $tsproj = $null
    if (Test-Path -LiteralPath $ctx.sln) {
        $txt = Get-Content $ctx.sln -Raw -Encoding UTF8
        $m = [regex]::Match($txt, 'Project\("[^"]+"\)\s*=\s*"([^"]+)"\s*,\s*"([^"]+\.tsproj)"')
        if (-not $m.Success) { [void]$errs.Add("sln 内未找到 .tsproj 引用") }
        else {
            $slnProjName = $m.Groups[1].Value
            $tsRel = $m.Groups[2].Value -replace '\\', [System.IO.Path]::DirectorySeparatorChar
            $tsproj = Join-Path (Split-Path $ctx.sln -Parent) $tsRel
            $detail.slnProjectName = $slnProjName
            $detail.tsproj = $tsproj
            if (-not (Test-Path -LiteralPath $tsproj)) { [void]$errs.Add("tsproj 不存在: " + $tsproj) }
            $rootCanon = $ctx.repoRoot
            if ((Resolve-Path -LiteralPath $tsproj -ErrorAction SilentlyContinue) -and (-not ((Resolve-Path -LiteralPath $tsproj).Path.ToLowerInvariant().StartsWith($rootCanon.ToLowerInvariant() + [System.IO.Path]::DirectorySeparatorChar)))) {
                [void]$errs.Add("tsproj 不在 ExpectedRepoRoot 下: " + $tsproj)
            }
            if ($ctx.expectedProjectName) {
                if ($slnProjName -ne $ctx.expectedProjectName) { [void]$errs.Add("项目名不符: 实际=" + $slnProjName + " 期望=" + $ctx.expectedProjectName) }
            }
        }
    }
    # 3) tsproj TargetNetId == 请求 NetId
    if ($tsproj -and (Test-Path -LiteralPath $tsproj)) {
        $tj = Get-Content $tsproj -Raw -Encoding UTF8
        $mn = [regex]::Match($tj, 'TargetNetId="([^"]+)"')
        $tz = $null
        if ($mn.Success) { $tz = $mn.Groups[1].Value }
        $detail.tsprojTargetNetId = $tz
        if (-not $tz) { [void]$errs.Add("tsproj 缺 TargetNetId") }
        elseif ($tz -ne $ctx.netid) { [void]$errs.Add("NetId 不符: tsproj=" + $tz + " 请求=" + $ctx.netid) }
        $mp = [regex]::Match($tj, 'PrjFilePath="([^"]+)"')
        if ($mp.Success) {
            $plcRel = $mp.Groups[1].Value -replace '\\', [System.IO.Path]::DirectorySeparatorChar
            $plcproj = Join-Path (Split-Path $tsproj -Parent) $plcRel
            $detail.plcproj = $plcproj
            if (-not (Test-Path -LiteralPath $plcproj)) { [void]$errs.Add("plcproj 不存在: " + $plcproj) }
        } else { [void]$errs.Add("tsproj 缺 PrjFilePath") }
    }
    # 4) 实际打开的 XAE（COM 优先；不可用则拒绝，绝不猜）
    $xae = Get-OpenedXaeInfo
    $detail.xaeState = $xae.state
    $detail.xaeDetail = $xae.detail
    if ($xae.state -eq "ok") {
        $detail.xaeSolution = $xae.sln
        $detail.xaeTsproj = $xae.tsproj
        $expSln = (Resolve-Path -LiteralPath $ctx.sln -ErrorAction SilentlyContinue)
        if ($expSln -and $xae.sln) {
            $actSln = (Resolve-Path -LiteralPath $xae.sln -ErrorAction SilentlyContinue)
            if (-not $actSln) { [void]$errs.Add("实际 XAE 方案路径不可解析: " + $xae.sln) }
            elseif ($actSln.Path -ne $expSln.Path) { [void]$errs.Add("实际打开的 XAE 方案不是期望方案（跨仓库风险）: " + $actSln.Path + " ≠ " + $expSln.Path) }
        }
        if ($tsproj -and $xae.tsproj) {
            $actTs = (Resolve-Path -LiteralPath $xae.tsproj -ErrorAction SilentlyContinue)
            $expTs = (Resolve-Path -LiteralPath $tsproj -ErrorAction SilentlyContinue)
            if (-not $actTs -or -not $expTs -or $actTs.Path -ne $expTs.Path) { [void]$errs.Add("实际打开的工程不是期望 tsproj: " + $xae.tsproj) }
        }
    } elseif ($xae.state -eq "none") {
        if (-not $ctx.dryRun) { [void]$errs.Add("XAE 未运行，无法校验实际打开方案（禁止盲激活）；请用 PS5.1 打开正确 solution 后再执行") }
        else { $detail.note = "dry-run：XAE 未运行，仅做磁盘身份校验" }
    } else {
        [void]$errs.Add("XAE COM 不可用（不猜）：" + $xae.detail + "；恢复步骤：用 Windows PowerShell 5.1 以相同权限启动正确 solution，或由人在状态栏核对 repository 后再执行")
    }
    $script:GuardDetail = $detail
    foreach ($k in $detail.Keys) { Say ("  guard." + $k + " = " + $detail[$k]) }
    if ($errs.Count -gt 0) { return ($errs -join "；") }
    return $null
}

function Invoke-Status([object]$ctx) {
    Say "=== status ==="
    $sln = $ctx.sln
    Step "solution" $true $sln
    $info = Get-TsprojInfo $sln
    Step "tsproj" ([bool]$info.tsproj) ([string]$info.tsproj)
    if ($info.targetNetId) { Step "tsproj.TargetNetId" $true $info.targetNetId } else { Step "tsproj.TargetNetId" $false "缺 TargetNetId" }
    if ($info.tmcPath) {
        $st = Get-TmcStamp $info.tmcPath
        Step "TMC" $st.exists ("{0} bytes @ {1}" -f $st.bytes, $st.mtime)
    } else { Step "TMC" $false "未找到 TMC 路径" }
    $dte = $null
    try { $dte = [Runtime.InteropServices.Marshal]::GetActiveObject("TcXaeShell.DTE.17.0") } catch { }
    if ($dte) {
        try { Step "dte" $true ("TcXaeShell " + $dte.Version + " / " + $dte.Solution.FullName) } catch { Step "dte" $false "DTE 存在但读方案失败" }
        try { Step "dte.projects" ($dte.Solution.Projects.Count -ge 1) ([string]$dte.Solution.Projects.Count) } catch { }
    } else {
        Step "dte" $false "没有运行中的 TcXaeShell（可用不带 -NoLaunch 的 compile/activate 自动打开）"
    }
    $sysState = Read-AdsStateSafe $ctx.netid 10000
    Step "target.10000" $sysState.ok ("state=" + $sysState.state + " device=" + $sysState.device + " err=" + $sysState.error)
    $plcState = Read-AdsStateSafe $ctx.netid 851
    Step "target.851" ($plcState.ok -and $plcState.state -eq 5) ("state=" + $plcState.state + " err=" + $plcState.error)
    $programMatch = $true
    if ($plcState.ok -and $plcState.state -eq 5) {
        $disc = Read-AdsSymbols $ctx.netid 851 @(@{ name = "GVL_Diag.nSelfTestTick"; type = "udint" }, @{ name = "GVL_Heartbeat.ControlTaskHeartbeat"; type = "udint" })
        $programMatch = -not ([string]$disc["GVL_Diag.nSelfTestTick"]).StartsWith("ERR")
        Step "target.851.program" $programMatch ($(if ($programMatch) { "示例验证符号可读（GVL_Diag）" } else { "851 RUN 但示例验证符号不可读 → 可能运行其它 PLC 程序；请按你的工程调整验证符号" }))
        if (-not $programMatch) { Say ("  程序身份判别: " + [string]$disc["GVL_Diag.nSelfTestTick"]) }
    }
    $wins = @(Get-XaeWindows)
    if ($wins.Count -gt 0) {
        foreach ($w in $wins) { Say ("  window: pid=" + $w.pid + " type=" + $w.controlType + " name=" + $w.name) }
        $blocked = @(Test-BlockedWindows $wins)
        Step "xae.windows" ($blocked.Count -eq 0) ($(if ($blocked.Count -gt 0) { "疑似弹窗: " + ($blocked -join " | ") } else { "无许可/模态弹窗关键词" }))
    }
    $other = @(Get-OtherProjectActivity $ctx.sln)
    Step "otherProject.activity" ($other.Count -eq 0) ($(if ($other.Count -gt 0) { "其它项目活动: " + (($other | ForEach-Object { "pid=" + $_.pid + " " + $_.name }) -join "; ") } else { "无（串行约束满足）" }))
    if ($other.Count -gt 0) { Finish 13 ("其它项目/工具正在活动，按串行约束停手: " + (($other | Select-Object -First 3 | ForEach-Object { $_.cmd }) -join " | ")) }
    if ($plcState.ok -and $plcState.state -eq 5 -and $programMatch) { Finish 0 "" }
    elseif ($plcState.ok -and $plcState.state -eq 5 -and -not $programMatch) { Finish 12 "851 RUN 但示例验证符号不可读；不能当作预期 PLC 已激活" }
    elseif (-not $sysState.ok) { Finish 10 ("目标 :10000 不可达: " + $sysState.error) }
    elseif ($sysState.state -ne 5) { Finish 10 ("目标 :10000 状态=" + $sysState.state + "（期望 5=RUN）") }
    else { Finish 11 ("PLC :851 未 RUN（error=" + $plcState.error + "）") }
}

function Invoke-Compile([object]$ctx) {
    Say "=== compile ==="
    $dte = Get-Dte $ctx.sln
    if ($null -eq $dte) { Finish 3 "无法附加/启动 TcXaeShell" }
    $bad = Assert-Solution $dte $ctx.sln
    if ($bad) { Finish 4 $bad }
    $info = Get-TsprojInfo $ctx.sln
    $before = Get-TmcStamp $info.tmcPath
    Say ("  TMC before: " + $before.bytes + " B @ " + $before.mtime)
    $t0 = Get-Date
    try {
        $dte.ExecuteCommand("Build.RebuildSolution")
        Say "  Build.RebuildSolution 已发出"
    } catch {
        Say ("  Build.RebuildSolution 失败: " + $_.Exception.Message + "；回退 SolutionBuild.Rebuild")
        try { $dte.Solution.SolutionBuild.Rebuild($true) } catch { Say ("  SolutionBuild.Rebuild 也失败: " + $_.Exception.Message) }
    }
    $deadline = (Get-Date).AddSeconds($ctx.timeout)
    $sawChange = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        $now = Get-TmcStamp $info.tmcPath
        if ($now.exists -and ($now.bytes -ne $before.bytes -or $now.mtime -ne $before.mtime)) { $sawChange = $true; break }
        try {
            $bs = [int]$dte.Solution.SolutionBuild.BuildState
            if ($bs -ge 2 -and ((Get-Date) - $t0).TotalSeconds -gt 5) { break }
        } catch { }
    }
    $after = Get-TmcStamp $info.tmcPath
    $lastBuild = -1
    try { $lastBuild = [int]$dte.Solution.SolutionBuild.LastBuildInfo } catch { }
    $all = @(Read-DteErrors $dte)
    $errs = @($all | Where-Object { $_.description -match "^(C|E|A|W)\d{4}\." -or $_.description -match "(?i)\berror\b" -or $_.description -match "错误" })
    $dur = [int]((Get-Date) - $t0).TotalSeconds
    Say ("  TMC after: " + $after.bytes + " B @ " + $after.mtime + " / lastBuildInfo=" + $lastBuild + " / errors=" + $errs.Count + " / 耗时=" + $dur + "s")
    foreach ($e in $errs | Select-Object -First 10) { Say ("    err: " + $e.description + " @ " + $e.file + ":" + $e.line) }
    $blocked = @(Test-BlockedWindows @(Get-XaeWindows))
    Step "tmc.refreshed" ($sawChange -or $lastBuild -eq 0) ($(if ($sawChange) { "TMC 已刷新" } else { "TMC 未变化（源码无变更时正常）" }))
    Step "lastBuildInfo" ($lastBuild -eq 0 -or $lastBuild -eq -1) ([string]$lastBuild)
    Step "errors" ($errs.Count -eq 0) ([string]$errs.Count)
    if ($errs.Count -gt 0) { Finish 5 ("编译错误 " + $errs.Count + " 条（首条: " + $errs[0].description + "）") }
    if ($lastBuild -ne 0 -and $lastBuild -ne -1) { Finish 5 ("LastBuildInfo=" + $lastBuild) }
    if (-not $sawChange -and $lastBuild -eq -1) { Finish 5 "无法判定编译完成（TMC 未刷新且 LastBuildInfo 不可读）" }
    if ($blocked.Count -gt 0) { Finish 9 ("编译后出现疑似模态弹窗: " + ($blocked -join " | ")) }
    Finish 0 ""
}

function Invoke-Boot([object]$ctx) {
    Say "=== boot ==="
    $dte = Get-Dte $ctx.sln
    if ($null -eq $dte) { Finish 3 "无法附加/启动 TcXaeShell" }
    $bad = Assert-Solution $dte $ctx.sln
    if ($bad) { Finish 4 $bad }
    $sys = $null
    try { $sys = Invoke-ComRetry { Get-SysManager $dte } 20 2 "取 ITcSysManager" } catch { Finish 4 $_.Exception.Message }
    if ($null -eq $sys) { Finish 4 "ITcSysManager 为 null（工程未就绪）" }
    $nodes = $null
    try { $nodes = Invoke-ComRetry { Get-PlcProjectNode $sys } 15 2 "取 ITcPlcProject" } catch { Finish 4 ("TIPC/Start 节点不可用: " + $_.Exception.Message) }
    $plc = $nodes.plc
    Say ("  PLC 节点: " + $nodes.node.Name + " type=" + $nodes.node.ItemType)
    $compileOk = $false
    try { $compileOk = [bool](Invoke-ComRetry { $plc.CompileProject() } 10 2 "CompileProject") } catch { Say ("  CompileProject 异常: " + $_.Exception.Message) }
    Step "CompileProject" $compileOk ([string]$compileOk)
    try { $plc.BootProjectAutostart = $true; Step "BootProjectAutostart" $true "TRUE" } catch { Step "BootProjectAutostart" $false $_.Exception.Message }
    $genOk = $false
    try { Invoke-ComRetry { $plc.GenerateBootProject() } 10 2 "GenerateBootProject" | Out-Null; $genOk = $true } catch { Say ("  GenerateBootProject 异常: " + $_.Exception.Message) }
    Step "GenerateBootProject" $genOk ([string]$genOk)
    $info = Get-TsprojInfo $ctx.sln
    $tmc = Get-TmcStamp $info.tmcPath
    Step "tmc.exists" $tmc.exists ("{0} B @ {1}" -f $tmc.bytes, $tmc.mtime)
    if (-not $genOk -or -not $tmc.exists) { Finish 6 "boot 工程生成失败（GenerateBootProject/TMC）" }
    if (-not $compileOk) {
        # 2026-11-17 D4 实测：本机 XAE 17.0 的 CompileProject() 返回 False，但
        # GenerateBootProject 成功且激活后新代码生效（FaultCode 3→5）。故此处只告警，
        # 仍以 GenerateBootProject 成功 + TMC 存在为 boot 硬判据；不得静默忽略 GenerateBootProject 失败。
        Say "  [WARN] CompileProject 返回 False；本机以 GenerateBootProject 结果为准（激活后必须 verify）"
    }
    Finish 0 ""
}

function Invoke-Config([object]$ctx) {
    Say "=== config（目标切 Config=15）==="
    $other = @(Get-OtherProjectActivity $ctx.sln)
    if ($other.Count -gt 0) { Finish 13 ("其它项目/工具正在活动，拒绝切 Config（串行约束）: " + (($other | Select-Object -First 3 | ForEach-Object { $_.cmd }) -join " | ")) }
    $st = Read-AdsStateSafe $ctx.netid 10000
    Step "target.10000.before" $st.ok ("state=" + $st.state + " err=" + $st.error)
    if (-not $st.ok) { Finish 7 ("目标 :10000 不可达: " + $st.error) }
    if ($st.state -eq 15) { Step "target.10000.config" $true "已在 Config(15)"; Finish 0 "" }
    $ok = Set-AdsModeWait $ctx.netid 15 $ctx.timeout
    $st2 = Read-AdsStateSafe $ctx.netid 10000
    Step "target.10000.config" ($ok -and $st2.state -eq 15) ("state=" + $st2.state + " err=" + $st2.error)
    if (-not ($ok -and $st2.state -eq 15)) { Finish 7 ("切 Config 失败（state=" + $st2.state + "）") }
    Finish 0 ""
}

function Invoke-Activate([object]$ctx) {
    Say "=== activate ==="
    $other = @(Get-OtherProjectActivity $ctx.sln)
    if ($other.Count -gt 0) { Finish 13 ("其它项目/工具正在活动，拒绝激活（串行约束）: " + (($other | Select-Object -First 3 | ForEach-Object { $_.cmd }) -join " | ")) }
    $dte = Get-Dte $ctx.sln
    if ($null -eq $dte) { Finish 3 "无法附加/启动 TcXaeShell" }
    $bad = Assert-Solution $dte $ctx.sln
    if ($bad) { Finish 4 $bad }
    $sys = $null
    try { $sys = Invoke-ComRetry { Get-SysManager $dte } 20 2 "取 ITcSysManager" } catch { Finish 4 $_.Exception.Message }
    if ($null -eq $sys) { Finish 4 "ITcSysManager 为 null" }
    try { [void]$dte.SuppressUI } catch { }
    try { $dte.UserControl = $false } catch { }
    try { Invoke-ComRetry { $sys.SetTargetNetId($ctx.netid) } 10 2 "SetTargetNetId" | Out-Null; Step "SetTargetNetId" $true $ctx.netid } catch { Step "SetTargetNetId" $false $_.Exception.Message }
    try {
        $sys.AutoSaveProjectToTarget = $true
        Step "AutoSaveProjectToTarget" $true "TRUE"
    } catch { Step "AutoSaveProjectToTarget" $false "属性不可用（不影响激活）" }
    $actOk = $false
    try {
        Invoke-ComRetry { $sys.ActivateConfiguration() } 60 2 "ActivateConfiguration" | Out-Null
        $actOk = $true
    } catch {
        Say ("  ActivateConfiguration 异常: " + $_.Exception.Message)
    }
    Step "ActivateConfiguration" $actOk ([string]$actOk)
    if (-not $actOk) {
        $blocked = @(Test-BlockedWindows @(Get-XaeWindows))
        if ($blocked.Count -gt 0) { Finish 9 ("激活被弹窗阻塞: " + ($blocked -join " | ")) }
        Finish 8 "ActivateConfiguration 失败（无弹窗关键词；需查 XAE 状态/许可）"
    }
    Finish 0 ""
}

function Invoke-Run([object]$ctx) {
    Say "=== run（StartRestartTwinCAT）==="
    $other = @(Get-OtherProjectActivity $ctx.sln)
    if ($other.Count -gt 0) { Finish 13 ("其它项目/工具正在活动，拒绝重启（串行约束）: " + (($other | Select-Object -First 3 | ForEach-Object { $_.cmd }) -join " | ")) }
    $dte = Get-Dte $ctx.sln
    if ($null -eq $dte) { Finish 3 "无法附加/启动 TcXaeShell" }
    $bad = Assert-Solution $dte $ctx.sln
    if ($bad) { Finish 4 $bad }
    $sys = $null
    try { $sys = Invoke-ComRetry { Get-SysManager $dte } 15 2 "取 ITcSysManager" } catch { Finish 4 $_.Exception.Message }
    if ($null -eq $sys) { Finish 4 "ITcSysManager 为 null" }
    $ok = $false
    for ($i = 1; $i -le 6; $i++) {
        try { $sys.StartRestartTwinCAT(); $ok = $true; break } catch { Say ("  StartRestartTwinCAT 重试 " + $i + "/6: " + $_.Exception.Message); Start-Sleep -Seconds 3 }
    }
    Step "StartRestartTwinCAT" $ok ([string]$ok)
    if (-not $ok) {
        $blocked = @(Test-BlockedWindows @(Get-XaeWindows))
        if ($blocked.Count -gt 0) { Finish 9 ("重启被弹窗阻塞: " + ($blocked -join " | ")) }
        Finish 8 "StartRestartTwinCAT 失败"
    }
    Finish 0 ""
}

function Invoke-Wait([object]$ctx) {
    Say "=== wait（10000→RUN, 851→RUN）==="
    $deadline = (Get-Date).AddSeconds($ctx.timeout)
    $sysState = $null; $plcState = $null
    while ((Get-Date) -lt $deadline) {
        $s = Read-AdsStateSafe $ctx.netid 10000
        if ($s.ok -and $s.state -eq 5) { $sysState = 5 } else { $sysState = $s.state }
        if ($sysState -eq 5) {
            $p = Read-AdsStateSafe $ctx.netid 851
            if ($p.ok -and $p.state -eq 5) { $plcState = 5; break } else { $plcState = $p.error }
        }
        Start-Sleep -Seconds 2
    }
    Step "target.10000.run" ($sysState -eq 5) ([string]$sysState)
    Step "target.851.run" ($plcState -eq 5) ([string]$plcState)
    if ($sysState -ne 5) { Finish 10 ("目标 :10000 超时未到 RUN（state=" + $sysState + "）") }
    if ($plcState -ne 5) { Finish 11 ("PLC :851 超时未到 RUN（" + $plcState + "）") }
    Finish 0 ""
}

function Invoke-Verify([object]$ctx) {
    Say "=== verify（只读 851）==="
    $p = Read-AdsStateSafe $ctx.netid 851
    Step "target.851" ($p.ok -and $p.state -eq 5) ("state=" + $p.state + " err=" + $p.error)
    if (-not ($p.ok -and $p.state -eq 5)) { Finish 11 ("851 不可用: " + $p.error) }
    $defs = @(
        @{ name = "GVL_Diag.nSelfTestTick"; type = "udint" },
        @{ name = "GVL_Diag.bSelfTestDone"; type = "bool" },
        @{ name = "GVL_IoMap.N_IO_AXES"; type = "int" },
        @{ name = "GVL_Heartbeat.ControlTaskHeartbeat"; type = "udint" },
        @{ name = "GVL_Command.eCommand"; type = "int" }
    )
    $a = Read-AdsSymbols $ctx.netid 851 $defs
    Start-Sleep -Seconds 1
    $b = Read-AdsSymbols $ctx.netid 851 $defs
    $readable = 0
    foreach ($k in $a.Keys) { if (-not ([string]$a[$k]).StartsWith("ERR")) { $readable++ } }
    Step "symbols.readable" ($readable -eq $defs.Count) ([string]$readable + "/" + $defs.Count)
    $hbA = 0; $hbB = 0
    [void][uint32]::TryParse([string]$a["GVL_Heartbeat.ControlTaskHeartbeat"], [ref]$hbA)
    [void][uint32]::TryParse([string]$b["GVL_Heartbeat.ControlTaskHeartbeat"], [ref]$hbB)
    $hbInc = ($hbB -gt $hbA)
    Step "heartbeat.increasing" $hbInc ([string]$hbA + "→" + [string]$hbB)
    $axes = $a["GVL_IoMap.N_IO_AXES"]
    if ($ctx.expectedAxes -gt 0) { Step "GVL_IoMap.N_IO_AXES" ($axes -eq $ctx.expectedAxes) ([string]$axes) }
    $tick = $a["GVL_Diag.nSelfTestTick"]
    $done = $a["GVL_Diag.bSelfTestDone"]
    $cmd = $a["GVL_Command.eCommand"]
    foreach ($kv in $a.GetEnumerator()) { Say ("  " + $kv.Key + " = " + $kv.Value) }
    if (-not ($hbInc -and ($ctx.expectedAxes -le 0 -or $axes -eq $ctx.expectedAxes) -and $readable -eq $defs.Count)) { Finish 12 "verify 判据不满足（心跳/轴数/符号可读性）" }
    Say ("  自检: tick=" + $tick + " done=" + $done + " eCommand=" + $cmd + "（只读，不驱动）")
    Finish 0 ""
}

# ───────────────────────────── 主流程 ─────────────────────────────
$sln = Resolve-SolutionPath $Solution
if (-not $sln) {
    if ($Command -in @("boot", "config", "activate", "run", "all")) {
        Say ("IDENTITY-REJECT(13): -Solution 无效/不是 .sln/不存在: " + $Solution)
        $script:Result.error = "identity: invalid solution"
        $script:Result.exitCode = 13
        $script:Result.steps = @($script:StepRows)
        if ($Json) { $script:Result | ConvertTo-Json -Depth 6 | Write-Output }
        exit 13
    }
    Say ("INVALID: -Solution 无效或不存在: " + $Solution)
    $script:Result.error = "solution invalid"
    $script:Result.exitCode = 2
    if ($Json) { $script:Result | ConvertTo-Json -Depth 6 | Write-Output }
    exit 2
}
$info = Get-TsprojInfo $sln
if (-not $NetId) {
    if ($info.targetNetId) { $NetId = [string]$info.targetNetId } else { Finish 2 "未指定 -NetId，且 tsproj 无 TargetNetId" }
}
if (-not $ExpectedRepoRoot) { $ExpectedRepoRoot = Split-Path $PSScriptRoot -Parent }
$repoRootResolved = (Resolve-Path -LiteralPath $ExpectedRepoRoot -ErrorAction SilentlyContinue)
if ($repoRootResolved) { $ExpectedRepoRoot = $repoRootResolved.Path } else { Say ("WARN ExpectedRepoRoot 不可解析: " + $ExpectedRepoRoot) }
$ctx = [ordered]@{ sln = $sln; netid = $NetId; timeout = $Timeout; repoRoot = $ExpectedRepoRoot; expectedProjectName = $ExpectedProjectName; expectedAxes = $ExpectedAxes; dryRun = [bool]$DryRun }
Say ("xae_cli {0} | solution={1} | netid={2} | PS={3}" -f $Command, $sln, $NetId, $PSVersionTable.PSVersion)
Say ("  tsproj={0} targetNetId={1} tmc={2}" -f $info.tsproj, $info.targetNetId, $info.tmcPath)
Say ("  窗口: " + ((@(Get-XaeWindows) | ForEach-Object { $_.name }) -join " | "))

if ($Command -in @("boot", "config", "activate", "run", "all")) {
    $deny = Assert-IdentityGuard $ctx
    if ($deny) { Finish 13 ("身份守卫拒绝: " + $deny) }
    if ($DryRun) { Say "DRY-RUN: 身份守卫通过；未执行任何编译/激活/切模式/重启。"; Finish 0 "" }
}
switch ($Command) {
    "fingerprint" { Invoke-Fingerprint $ctx }
    "status" { Invoke-Status $ctx }
    "compile" { Invoke-Compile $ctx }
    "boot" { Invoke-Boot $ctx }
    "config" { Invoke-Config $ctx }
    "activate" { Invoke-Activate $ctx }
    "run" { Invoke-Run $ctx }
    "wait" { Invoke-Wait $ctx }
    "verify" { Invoke-Verify $ctx }
    "all" {
        Invoke-Status $ctx
        Invoke-Compile $ctx
        Invoke-Boot $ctx
        Invoke-Config $ctx
        Invoke-Activate $ctx
        Invoke-Run $ctx
        Invoke-Wait $ctx
        Invoke-Verify $ctx
    }
}

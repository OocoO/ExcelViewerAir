# dedup-smoke.ps1 - 真机验证「同一个表只开一个窗口」
#
# 真的开窗口跑一遍（不是离屏）：
#   1) 打开 A            → 1 个进程 1 个窗口
#   2) 再打开 A          → 还是 1 个进程 1 个窗口，第二次启动的进程自己退出，A 的窗口被提到前台
#   3) 打开 B            → 还是 1 个进程，多一个窗口
#   4) 再打开 A          → 窗口数不变，前台切回 A 的窗口
#
# 用法（默认验 viewer\publish 里的 exe）：
#   pwsh -File tools\dedup-smoke.ps1
#   pwsh -File tools\dedup-smoke.ps1 -Exe viewer\publish-test\ExcelViewer.exe
#
# 注意：会真的弹出查看器窗口；脚本只关掉自己启动的那些进程，已经开着的窗口不动。
# 会话不在交互式桌面上（比如某些远程/服务环境）时，前台窗口那两项会报"环境不支持"，
# 其余项仍然有效。

param(
    [string]$Exe = '',
    [string]$FileA = '',
    [string]$FileB = ''
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Exe)) { $Exe = Join-Path $root 'viewer\publish\ExcelViewer.exe' }
elseif (-not [System.IO.Path]::IsPathRooted($Exe)) { $Exe = Join-Path $root $Exe }
if ([string]::IsNullOrWhiteSpace($FileA)) { $FileA = Join-Path $root 'test\Item.xlsx' }
if ([string]::IsNullOrWhiteSpace($FileB)) { $FileB = Join-Path $root 'test\Level.xlsx' }

foreach ($p in @($Exe, $FileA, $FileB)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "找不到：$p" }
}

if (-not ('ExcelViewerSmokeWin' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExcelViewerSmokeWin
{
    private delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowTextW(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassNameW(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

    /// <summary>最小化（SW_MINIMIZE = 6）。</summary>
    public static void Minimize(IntPtr hwnd) { ShowWindow(hwnd, 6); }

    public static bool IsMinimized(IntPtr hwnd) { return IsIconic(hwnd); }

    /// <summary>给窗口发 WM_CLOSE，等同于用户点右上角的叉。</summary>
    public static void CloseWindow(IntPtr hwnd) { PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero); }

    /// <summary>某个进程的可见顶层窗口（带标题的非工具窗口）。</summary>
    public static List<IntPtr> WindowsOf(int processId)
    {
        var found = new List<IntPtr>();
        EnumWindows((hwnd, _) =>
        {
            if (!IsWindowVisible(hwnd)) return true;
            GetWindowThreadProcessId(hwnd, out var pid);
            if (pid != (uint)processId) return true;
            if (ClassOf(hwnd) == "Shell_TrayWnd") return true;
            if (TitleOf(hwnd).Length == 0) return true;
            found.Add(hwnd);
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static string TitleOf(IntPtr hwnd)
    {
        var sb = new StringBuilder(512);
        GetWindowTextW(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static string ClassOf(IntPtr hwnd)
    {
        var sb = new StringBuilder(256);
        GetClassNameW(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static int ForegroundProcessId
    {
        get
        {
            var hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return 0;
            GetWindowThreadProcessId(hwnd, out var pid);
            return (int)pid;
        }
    }

    public static string ForegroundTitle
    {
        get
        {
            var hwnd = GetForegroundWindow();
            return hwnd == IntPtr.Zero ? string.Empty : TitleOf(hwnd);
        }
    }
}
'@
}

$failures = 0
function Check([string]$what, [bool]$ok, [string]$detail = '') {
    $note = if ($ok -or [string]::IsNullOrEmpty($detail)) { '' } else { "  → $detail" }
    Write-Host "  [$(if ($ok) { 'OK' } else { 'FAIL' })] $what$note"
    if (-not $ok) { $script:failures++ }
}

$started = New-Object System.Collections.Generic.List[int]

# 启动一律走 CreateProcess（UseShellExecute = $false），不用 Start-Process：
# Start-Process 走的是 ShellExecute，而这个 exe 没有数字签名 —— 从受限/沙箱进程里
# 通过 shell 启动时，Windows 会先弹「打开文件 - 安全警告：无法验证发布者」，
# 并且**卡住调用方**，直到有人去点「运行」。直接 CreateProcess 不经过这一步。
function Start-Viewer([string]$file) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.Arguments = '"' + $file + '"'
    $psi.UseShellExecute = $false
    $proc = [System.Diagnostics.Process]::Start($psi)
    $started.Add($proc.Id)
    return $proc
}

function Start-BareViewer {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.UseShellExecute = $false
    $proc = [System.Diagnostics.Process]::Start($psi)
    $started.Add($proc.Id)
    return $proc
}

function Wait-Windows([int]$processId, [int]$want, [int]$timeoutSec = 20) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    while ((Get-Date) -lt $deadline) {
        $w = [ExcelViewerSmokeWin]::WindowsOf($processId)
        if ($w.Count -ge $want) { return $w }
        Start-Sleep -Milliseconds 150
    }
    return [ExcelViewerSmokeWin]::WindowsOf($processId)
}

function Get-HostProcess {
    # 主持进程＝还在跑、且有可见窗口的那个
    $alive = @(Get-Process -Name 'ExcelViewer' -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -eq $Exe -and $_.HasExited -eq $false })
    return @($alive | Where-Object { ([ExcelViewerSmokeWin]::WindowsOf($_.Id)).Count -gt 0 })
}

Write-Host "查看器: $Exe"
Write-Host "测试表: $FileA / $FileB"
Write-Host ''

# ---- 1) 第一次打开 ----
Write-Host '1) 打开第一个表'
$first = Start-Viewer $FileA
$windows = Wait-Windows $first.Id 1
Check '开出一个窗口' ($windows.Count -eq 1) "实际 $($windows.Count) 个"
$hostId = $first.Id

# ---- 2) 再打开同一个表 ----
Write-Host '2) 再打开同一个表'
$second = Start-Viewer $FileA
Start-Sleep -Milliseconds 1500
$second.Refresh()
Check '第二次启动的进程把请求交出去后自己退出' $second.HasExited "仍然在跑（pid $($second.Id)）"

$hosts = Get-HostProcess
Check '仍然只有一个查看器进程' ($hosts.Count -eq 1) "实际 $($hosts.Count) 个"

$windows = [ExcelViewerSmokeWin]::WindowsOf($hostId)
Check '窗口数还是 1 个（没有重复开窗）' ($windows.Count -eq 1) "实际 $($windows.Count) 个"

$foregroundOk = [ExcelViewerSmokeWin]::ForegroundProcessId -eq $hostId
if ($foregroundOk) {
    Check '前台窗口就是这扇窗口' $true
}
else {
    Write-Host "  [--] 前台窗口检查跳过（当前桌面拿不到焦点：前台是 '$([ExcelViewerSmokeWin]::ForegroundTitle)'）"
}

# ---- 3) 打开另一个表 ----
Write-Host '3) 打开另一个表'
$third = Start-Viewer $FileB
$windows = Wait-Windows $hostId 2
Check '同一个进程里多开了一个窗口' ($windows.Count -eq 2) "实际 $($windows.Count) 个"
Check '第二个表没有再起一个进程' (@(Get-HostProcess).Count -eq 1)

# ---- 4) 再打开第一个表 ----
Write-Host '4) 再打开第一个表'
$fourth = Start-Viewer $FileA
Start-Sleep -Milliseconds 1500
$windows = [ExcelViewerSmokeWin]::WindowsOf($hostId)
Check '窗口数不变（定位到已有窗口）' ($windows.Count -eq 2) "实际 $($windows.Count) 个"

if ([ExcelViewerSmokeWin]::ForegroundProcessId -eq $hostId) {
    Check '前台窗口还在这个进程里' $true
}
else {
    Write-Host "  [--] 前台窗口检查跳过（当前桌面拿不到焦点：前台是 '$([ExcelViewerSmokeWin]::ForegroundTitle)'）"
}

# ---- 5) 最小化以后再打开同一个表：要还原并提到前台 ----
Write-Host '5) 把第一个表的窗口最小化，再打开同一个表'
$itemWindows = @([ExcelViewerSmokeWin]::WindowsOf($hostId) |
    Where-Object { [ExcelViewerSmokeWin]::TitleOf($_) -like '*Item.xlsx*' })
if ($itemWindows.Count -eq 0) {
    Check '找得到第一个表的窗口' $false '窗口标题里没有 Item.xlsx'
}
else {
    $itemWindow = [IntPtr]$itemWindows[0]
    [ExcelViewerSmokeWin]::Minimize($itemWindow)
    Start-Sleep -Milliseconds 400
    Check '窗口已被最小化（验证起点有效）' ([ExcelViewerSmokeWin]::IsMinimized($itemWindow))

    $null = Start-Viewer $FileA
    $deadline = (Get-Date).AddSeconds(5)
    while ((Get-Date) -lt $deadline -and [ExcelViewerSmokeWin]::IsMinimized($itemWindow)) {
        Start-Sleep -Milliseconds 120
    }

    Check '再打开同一个表会把最小化的窗口还原' (-not [ExcelViewerSmokeWin]::IsMinimized($itemWindow))
}

# ---- 6) 同一个表换个写法（.. 绕一圈）：仍然算同一个窗口 ----
Write-Host '6) 用 test\..\test\Item.xlsx 这种写法再打开一次'
$indirect = [System.IO.Path]::Combine($root, 'test', '..', 'test', 'Item.xlsx')
$null = Start-Viewer $indirect
Start-Sleep -Milliseconds 1500
Check '另一种路径写法没有多开窗口' (([ExcelViewerSmokeWin]::WindowsOf($hostId)).Count -eq 2) `
    "实际 $(([ExcelViewerSmokeWin]::WindowsOf($hostId)).Count) 个"

# ---- 7) 不带文件空跑：定位到已有窗口，不新开空窗口 ----
Write-Host '7) 不带任何文件启动'
$bare = Start-BareViewer
Start-Sleep -Milliseconds 1500
$bare.Refresh()
Check '空跑的那个进程把自己退掉了' $bare.HasExited '仍然在跑'
Check '没有多出空窗口' (([ExcelViewerSmokeWin]::WindowsOf($hostId)).Count -eq 2) `
    "实际 $(([ExcelViewerSmokeWin]::WindowsOf($hostId)).Count) 个"

# ---- 8) 关掉一扇窗口：查看器不能整个退出（ShutdownMode 必须按"最后一扇窗口"算） ----
Write-Host '8) 关掉其中一扇窗口'
$closeTarget = @([ExcelViewerSmokeWin]::WindowsOf($hostId) |
    Where-Object { [ExcelViewerSmokeWin]::TitleOf($_) -like '*Item.xlsx*' })
if ($closeTarget.Count -eq 0) {
    Check '找得到要关掉的那扇窗口' $false '窗口标题里没有 Item.xlsx'
}
else {
    [ExcelViewerSmokeWin]::CloseWindow([IntPtr]$closeTarget[0])
    Start-Sleep -Milliseconds 900

    $hostProcess = Get-Process -Id $hostId -ErrorAction SilentlyContinue
    Check '关掉一扇窗口后查看器进程还在' ($null -ne $hostProcess) '进程跟着退掉了'
    $left = @([ExcelViewerSmokeWin]::WindowsOf($hostId))
    Check '另一扇窗口没有被一起关掉' ($left.Count -eq 1) "剩 $($left.Count) 扇"
}

# ---- 清理：只关掉本次启动的进程 ----
foreach ($procId in $started) {
    $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
    if ($null -ne $p) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }
}

Write-Host ''
if ($failures -eq 0) { Write-Host 'DEDUP SMOKE OK' } else { Write-Host "DEDUP SMOKE FAILED: $failures 项不通过" }
exit $failures

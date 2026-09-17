# install-open-with.ps1 - make sure the HKCU registry entries exist, then send
# the user to the Windows "Default apps" page.
#
# Chinese text is kept out of the .cmd files entirely: cmd.exe parses a .cmd
# with the console's OEM codepage, and even a few UTF-8 Chinese bytes in a
# comment or a path can desync that parser (symptoms: "'xxx' is not recognized",
# "else was unexpected at this time"). PowerShell reads .ps1 as UTF-8, so it is
# the safe place for both the Chinese messages and the Chinese file name.

param([switch]$SkipSettings)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot     # project root (parent of tools\)

# ---- 1) import the .reg file (found by pattern, so no Chinese in the .cmd) ----
$regFiles = @(Get-ChildItem -LiteralPath $root -Filter '*.reg' -File |
    Where-Object { Select-String -LiteralPath $_.FullName -Pattern 'ExcelViewer\.Sheet' -Quiet })

if ($regFiles.Count -gt 0) {
    foreach ($f in $regFiles) {
        Write-Host "导入注册表: $($f.Name)"
        & reg.exe import $f.FullName | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "reg import 返回 $LASTEXITCODE（文件可能已经导入过，可忽略）"
        }
    }
} else {
    Write-Warning "没有在 $root 下找到包含 ExcelViewer.Sheet 的 .reg 文件"
}

# ---- 2) verify the entries really landed ----
$probe = 'HKCU:\Software\Classes\ExcelViewer.Sheet\shell\open\command'
$ok = Test-Path $probe
Write-Host $(if ($ok) { '注册表项已就绪。' } else { '注册表项缺失，导入似乎没有生效。' })

# ---- 3) explain the UserChoice limitation, then open Settings ----
Add-Type -AssemblyName System.Windows.Forms

if (-not $SkipSettings) {
    $message = @'
接下来会打开系统的「默认应用」设置页。

请把 .xlsx 的默认打开方式改成「Excel 只读查看器」。

更简单的做法：
  右键任意 .xlsx - 打开方式 - 选择其他应用 -
  选「Excel 只读查看器」并勾选「始终」

（Windows 不允许程序自己改默认打开方式，必须你确认一次。）
'@

    [System.Windows.Forms.MessageBox]::Show($message, '设为默认打开方式', 'OK', 'Information') | Out-Null
    Start-Process 'ms-settings:defaultapps'
}

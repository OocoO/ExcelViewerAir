# launch.ps1 - launcher logic for ExcelViewer (Chinese text is safe here:
# PowerShell reads .ps1 as UTF-8, unlike cmd.exe which parses .cmd with the
# console's OEM codepage).
#
# Called by ExcelViewer.cmd with the file to open (optional).

param([string]$File = '')

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot          # project root (parent of tools\)
$exe = Join-Path $root 'viewer\publish\ExcelViewer.exe'

if (-not (Test-Path -LiteralPath $exe)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "没有找到查看器：`n`n$exe`n`n请先在 excel2csv 目录执行：`n    dotnet publish viewer\ExcelViewer.csproj -c Release -o viewer\publish",
        'ExcelViewer',
        'OK',
        'Warning') | Out-Null
    Start-Process explorer.exe $root
    exit 1
}

if ([string]::IsNullOrWhiteSpace($File)) {
    Start-Process -FilePath $exe
} else {
    Start-Process -FilePath $exe -ArgumentList @($File)
}

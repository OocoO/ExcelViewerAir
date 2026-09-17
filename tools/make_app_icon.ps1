#Requires -Version 7.0
<#
  make_app_icon.ps1 — 把 AI 生成的方形图标原图，做成可以直接塞进 WPF 工程的图标资产。

  做三件事：
    1. 裁掉原图四周的纯色背景边（AI 出图常见的大片留白），只留圆角方块本身；
    2. 按 Windows 图标安全区重排画布：内容约占 80%，外圈留透明边，
       否则小尺寸（16/24px）下图标会顶满格子、比旁边的系统图标大一圈；
    3. 生成多尺寸 .ico（每档内嵌 PNG）+ 一份 256px PNG 预览。

  用法：
    pwsh -File tools\make_app_icon.ps1 -Source assets\icon\app-icon.png `
         -Ico viewer\app.ico -Png assets\icon\app-icon-256.png
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Ico,
    [string]$Png,
    [int[]]$Sizes = @(256, 128, 64, 48, 32, 24, 16),
    # 内容（裁完的圆角方块）占最终画布的比例，Windows 图标习惯 0.8 左右
    [double]$ContentRatio = 0.80,
    # 判成背景的亮度阈值：亮度 >= 该值的像素视为透明
    [int]$BackgroundLevel = 246,
    # 背景淡出区间宽度（BackgroundLevel 往下这么多像素内线性过渡，避免硬边）
    [int]$BackgroundFade = 10)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-TransparentBitmap([int]$w, [int]$h) {
    $bmp = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $bmp.SetResolution(96, 96)
    return $bmp
}

# ---------- 1. 读源图，转成 32bppArgb 并解开像素 ----------
$srcPath = (Resolve-Path -LiteralPath $Source).Path
$raw = [System.Drawing.Image]::FromFile($srcPath)
$width = $raw.Width
$height = $raw.Height
$flat = New-TransparentBitmap $width $height
$g = [System.Drawing.Graphics]::FromImage($flat)
$g.DrawImage($raw, 0, 0, $width, $height)
$g.Dispose()
$raw.Dispose()

$rect = New-Object System.Drawing.Rectangle 0, 0, $width, $height
$data = $flat.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadWrite, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$stride = $data.Stride
$bytes = New-Object byte[] ($stride * $height)
[System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $bytes, 0, $bytes.Length)

# BGRA 顺序，逐像素：把接近背景色的像素变透明（按亮度做一段渐变，保留抗锯齿边）
$minX = $width; $minY = $height; $maxX = -1; $maxY = -1
for ($y = 0; $y -lt $height; $y++) {
    $row = $y * $stride
    for ($x = 0; $x -lt $width; $x++) {
        $i = $row + $x * 4
        $b = $bytes[$i]; $gg = $bytes[$i + 1]; $r = $bytes[$i + 2]; $a = $bytes[$i + 3]
        $lum = [int](0.299 * $r + 0.587 * $gg + 0.114 * $b)
        if ($lum -ge $BackgroundLevel) {
            $bytes[$i] = 0; $bytes[$i + 1] = 0; $bytes[$i + 2] = 0; $bytes[$i + 3] = 0
            continue
        }
        if ($lum -gt ($BackgroundLevel - $BackgroundFade)) {
            $t = ($BackgroundLevel - $lum) / [double]$BackgroundFade   # 0..1
            $bytes[$i + 3] = [byte]([math]::Round($a * $t))
        }
        if ($bytes[$i + 3] -gt 0) {
            if ($x -lt $minX) { $minX = $x }
            if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }
            if ($y -gt $maxY) { $maxY = $y }
        }
    }
}
[System.Runtime.InteropServices.Marshal]::Copy($bytes, 0, $data.Scan0, $bytes.Length)
$flat.UnlockBits($data)

if ($maxX -lt 0) { throw "整张图都被判成了背景，检查 -BackgroundLevel（当前 $BackgroundLevel）" }

# 外圈前景像素大多来自柔和投影，用 2% 的分位裁剪，别让阴影把内容挤小
$cropW = $maxX - $minX + 1
$cropH = $maxY - $minY + 1
$cropX = [int]($minX + $cropW * 0.02)
$cropY = [int]($minY + $cropH * 0.02)
$cropW = [int]($cropW * 0.96)
$cropH = [int]($cropH * 0.96)
Write-Host ("源图 {0}x{1} → 内容区 {2},{3} {4}x{5}（背景 {6}% 面积）" -f `
    $width, $height, $cropX, $cropY, $cropW, $cropH, [math]::Round(100 * (1 - ($cropW * $cropH) / [double]($width * $height))))

# ---------- 2. 重排到带透明安全区的正方形画布 ----------
$side = [int]([math]::Round([math]::Max($cropW, $cropH) / $ContentRatio))
$canvas = New-TransparentBitmap $side $side
$g = [System.Drawing.Graphics]::FromImage($canvas)
$g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
$g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$dst = New-Object System.Drawing.Rectangle ([int](($side - $cropW) / 2)), ([int](($side - $cropH) / 2)), $cropW, $cropH
$srcRect = New-Object System.Drawing.Rectangle $cropX, $cropY, $cropW, $cropH
$g.DrawImage($flat, $dst, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
$g.Dispose()
$flat.Dispose()
Write-Host "重排画布: ${side}x${side}（内容占 $([int]($ContentRatio * 100))%）"

# ---------- 3. 输出各尺寸 PNG + .ico ----------
$blobs = New-Object System.Collections.Generic.List[byte[]]
$ordered = @($Sizes | Sort-Object -Descending)
foreach ($size in $ordered) {
    $bmp = New-TransparentBitmap $size $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.DrawImage($canvas, (New-Object System.Drawing.Rectangle 0, 0, $size, $size))
    $g.Dispose()

    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $blobs.Add($ms.ToArray())
    $ms.Dispose()

    if ($Png -and $size -eq 256) {
        $pngPath = [System.IO.Path]::GetFullPath($Png)
        New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($pngPath)) | Out-Null
        $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host "预览 PNG: $pngPath"
    }
    $bmp.Dispose()
}
$canvas.Dispose()

$icoPath = [System.IO.Path]::GetFullPath($Ico)
New-Item -ItemType Directory -Force -Path ([System.IO.Path]::GetDirectoryName($icoPath)) | Out-Null
$stream = [System.IO.File]::Create($icoPath)
$bw = New-Object System.IO.BinaryWriter($stream)
try {
    $count = $ordered.Count
    $bw.Write([uint16]0)       # reserved
    $bw.Write([uint16]1)       # type = icon
    $bw.Write([uint16]$count)
    $offset = 6 + 16 * $count
    for ($i = 0; $i -lt $count; $i++) {
        $size = $ordered[$i]
        # ICO 目录项的宽高是单字节，256 记 0
        $dim = if ($size -ge 256) { 0 } else { $size }
        $bw.Write([byte]$dim)
        $bw.Write([byte]$dim)
        $bw.Write([byte]0)     # 调色板数
        $bw.Write([byte]0)     # reserved
        $bw.Write([uint16]1)   # color planes
        $bw.Write([uint16]32)  # bit depth
        $bw.Write([uint32]$blobs[$i].Length)
        $bw.Write([uint32]$offset)
        $offset += $blobs[$i].Length
    }
    foreach ($blob in $blobs) { $bw.Write($blob) }
} finally {
    $bw.Dispose(); $stream.Dispose()
}
Write-Host ("ICO: {0}（{1:N0} 字节，尺寸 {2}）" -f $icoPath, (Get-Item -LiteralPath $icoPath).Length, ($ordered -join '/'))

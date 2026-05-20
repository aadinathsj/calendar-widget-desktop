# Creates a desktop shortcut + custom calendar icon for Calendar Widget
# Run via: npm run shortcut

param(
  [string]$Day = (Get-Date -Format "%d")
)

$projectDir = Split-Path -Parent $PSScriptRoot
$assetsDir  = Join-Path $projectDir "assets"
$icoPath    = Join-Path $assetsDir "icon.ico"
$pngPath    = Join-Path $assetsDir "icon.png"
$desktop    = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "Calendar Widget.lnk"
$electronExe  = Join-Path $projectDir "node_modules\electron\dist\electron.exe"

# ── 1. Create assets folder ────────────────────────────────────────────────
if (-not (Test-Path $assetsDir)) {
  New-Item -ItemType Directory -Path $assetsDir | Out-Null
}

# ── 2. Draw calendar icon with System.Drawing ──────────────────────────────
Add-Type -AssemblyName System.Drawing

$size = 256
$bmp = New-Object System.Drawing.Bitmap($size, $size)
$g   = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.Clear([System.Drawing.Color]::Transparent)

# Body (blue card)
$blue     = [System.Drawing.Color]::FromArgb(255, 0, 120, 212)
$darkBlue = [System.Drawing.Color]::FromArgb(255, 0, 80, 160)
$white    = [System.Drawing.Color]::White

$bodyBrush   = New-Object System.Drawing.SolidBrush($blue)
$headerBrush = New-Object System.Drawing.SolidBrush($darkBlue)
$whiteBrush  = New-Object System.Drawing.SolidBrush($white)
$blueBrush   = New-Object System.Drawing.SolidBrush($blue)

# Shadow
$shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60, 0, 0, 0))
$g.FillRectangle($shadowBrush, 28, 28, 212, 212)

# Card body
$g.FillRectangle($bodyBrush, 20, 20, 212, 212)

# Header strip (darker blue top)
$g.FillRectangle($headerBrush, 20, 20, 212, 64)

# Calendar rings (white circles with blue holes)
foreach ($rx in @(72, 184)) {
  $g.FillEllipse($whiteBrush, $rx - 14, 8, 28, 36)
  $g.FillEllipse($blueBrush,  $rx - 8,  18, 16, 16)
}

# Day number centred in lower portion
$font = New-Object System.Drawing.Font("Segoe UI", 96, [System.Drawing.FontStyle]::Bold)
$sf   = New-Object System.Drawing.StringFormat
$sf.Alignment     = [System.Drawing.StringAlignment]::Center
$sf.LineAlignment = [System.Drawing.StringAlignment]::Center
$textRect = New-Object System.Drawing.RectangleF(20, 84, 212, 148)
$g.DrawString($Day, $font, $whiteBrush, $textRect, $sf)

# Month label (small) - "MAY"
$monthFont = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Regular)
$monthRect  = New-Object System.Drawing.RectangleF(20, 24, 212, 46)
$g.DrawString((Get-Date -Format "MMM").ToUpper(), $monthFont, $whiteBrush, $monthRect, $sf)

$g.Dispose()

# ── 3. Save PNG ────────────────────────────────────────────────────────────
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "PNG saved: $pngPath"

# ── 4. Wrap PNG in ICO (Vista+ supports PNG-inside-ICO) ────────────────────
$pngBytes = [System.IO.File]::ReadAllBytes($pngPath)
$stream   = New-Object System.IO.MemoryStream
$bw       = New-Object System.IO.BinaryWriter($stream)

# ICONDIR
$bw.Write([uint16]0)   # reserved
$bw.Write([uint16]1)   # type: 1 = icon
$bw.Write([uint16]1)   # image count

# ICONDIRENTRY
$bw.Write([byte]0)     # width  0 => 256
$bw.Write([byte]0)     # height 0 => 256
$bw.Write([byte]0)     # color count
$bw.Write([byte]0)     # reserved
$bw.Write([uint16]1)   # planes
$bw.Write([uint16]32)  # bit count
$bw.Write([uint32]$pngBytes.Length)
$bw.Write([uint32]22)  # image offset (6 ICONDIR + 16 ICONDIRENTRY)

# PNG payload
$bw.Write($pngBytes)
$bw.Flush()

[System.IO.File]::WriteAllBytes($icoPath, $stream.ToArray())
$stream.Close()
Write-Host "ICO saved:  $icoPath"

# ── 5. Create / update desktop shortcut ────────────────────────────────────
if (-not (Test-Path $electronExe)) {
  Write-Warning "electron.exe not found at: $electronExe"
  Write-Warning "Run 'npm install' first, then try again."
  exit 1
}

$wsh      = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath      = $electronExe
$shortcut.Arguments       = "`"$projectDir`""
$shortcut.WorkingDirectory= $projectDir
$shortcut.IconLocation    = $icoPath
$shortcut.Description     = "Calendar Widget - Outlook Calendar"
$shortcut.Save()

Write-Host ""
Write-Host "Desktop shortcut created:" -ForegroundColor Green
Write-Host "  $shortcutPath"

# ── 6. Attempt to pin to Taskbar ───────────────────────────────────────────
Write-Host ""
Write-Host "Attempting to pin to Taskbar..." -ForegroundColor Cyan
try {
  $shell   = New-Object -ComObject "Shell.Application"
  $folder  = $shell.Namespace([System.IO.Path]::GetDirectoryName($shortcutPath))
  $file    = $folder.ParseName([System.IO.Path]::GetFileName($shortcutPath))
  $pinVerb = $file.Verbs() | Where-Object { ($_.Name -replace '&','') -match 'Pin to [Tt]askbar' }
  if ($pinVerb) {
    $pinVerb | Select-Object -First 1 | ForEach-Object { $_.DoIt() }
    Write-Host "Pinned to Taskbar successfully." -ForegroundColor Green
  } else {
    Write-Warning "Auto-pin unavailable on this Windows version."
    Write-Host "  To pin manually: right-click the desktop shortcut -> 'Pin to taskbar'" -ForegroundColor Yellow
  }
} catch {
  Write-Warning "Taskbar pin failed: $_"
  Write-Host "  To pin manually: right-click the desktop shortcut -> 'Pin to taskbar'" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Double-click the desktop shortcut to launch Calendar Widget."

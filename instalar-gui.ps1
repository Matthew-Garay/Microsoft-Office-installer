<#
.SYNOPSIS
    Microsoft Office Installer - GUI Edition
.DESCRIPTION
    Graphical tool to download and install Microsoft Office LTSC editions.
    Supports multiple editions, architectures, languages and app selection.
.NOTES
    Requirements: Administrator, PowerShell 5.0+, .NET Framework 4.5+
    Usage: irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-gui.ps1 | iex
#>

#Requires -RunAsAdministrator

# ---- LOAD ASSEMBLIES ----
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---- THEME ----
# Two full palettes: light (default) and dark. Apply-Theme repaints every
# registered control, so switching mode never leaves stale colors behind.
$script:themes = @{
    Light = @{
        Name = "Light"; Bg = [System.Drawing.Color]::FromArgb(243, 244, 246)
        Panel = [System.Drawing.Color]::White
        Text = [System.Drawing.Color]::FromArgb(22, 27, 34)
        Dim = [System.Drawing.Color]::FromArgb(100, 112, 126)
        Accent = [System.Drawing.Color]::FromArgb(27, 111, 220)
        AccentH = [System.Drawing.Color]::FromArgb(48, 130, 240)
        Ok = [System.Drawing.Color]::FromArgb(31, 141, 59)
        Bad = [System.Drawing.Color]::FromArgb(207, 54, 44)
        Warn = [System.Drawing.Color]::FromArgb(172, 114, 15)
        Track = [System.Drawing.Color]::FromArgb(223, 228, 234)
        Border = [System.Drawing.Color]::FromArgb(208, 215, 222)
        CardBd = [System.Drawing.Color]::FromArgb(208, 215, 222)
        CardOn = [System.Drawing.Color]::FromArgb(221, 240, 255)
        CardOnBd = [System.Drawing.Color]::FromArgb(27, 111, 220)
        ComboBg = [System.Drawing.Color]::White
        PanelH = [System.Drawing.Color]::FromArgb(232, 234, 237)
        Shadow = [System.Drawing.Color]::FromArgb(200, 208, 216)
    }
    Dark = @{
        Name = "Dark"; Bg = [System.Drawing.Color]::FromArgb(13, 17, 23)
        Panel = [System.Drawing.Color]::FromArgb(22, 27, 34)
        Text = [System.Drawing.Color]::FromArgb(230, 237, 243)
        Dim = [System.Drawing.Color]::FromArgb(139, 148, 158)
        Accent = [System.Drawing.Color]::FromArgb(47, 129, 247)
        AccentH = [System.Drawing.Color]::FromArgb(79, 151, 255)
        Ok = [System.Drawing.Color]::FromArgb(63, 185, 80)
        Bad = [System.Drawing.Color]::FromArgb(248, 81, 73)
        Warn = [System.Drawing.Color]::FromArgb(210, 153, 34)
        Track = [System.Drawing.Color]::FromArgb(30, 37, 49)
        Border = [System.Drawing.Color]::FromArgb(48, 58, 72)
        CardBd = [System.Drawing.Color]::FromArgb(48, 58, 72)
        CardOn = [System.Drawing.Color]::FromArgb(22, 45, 76)
        CardOnBd = [System.Drawing.Color]::FromArgb(47, 129, 247)
        ComboBg = [System.Drawing.Color]::FromArgb(22, 27, 34)
        PanelH = [System.Drawing.Color]::FromArgb(30, 37, 49)
        Shadow = [System.Drawing.Color]::FromArgb(0, 0, 0)
    }
}
$script:themeName = "Light"
function Get-Th { return $script:themes[$script:themeName] }
$script:themed = New-Object System.Collections.ArrayList
function Register-Theme { param($Control, [string]$Role)
    [void]$script:themed.Add(@{ C = $Control; Role = $Role }) }
function Apply-Theme {
    $t = Get-Th
    foreach ($e in $script:themed) {
        try {
            $c = $e.C
            switch ($e.Role) {
                "form" { $c.BackColor = $t.Bg }
                "panel" { $c.BackColor = $t.Panel; try { $c.ForeColor = $t.CardBd } catch {} }
                "label" { $c.ForeColor = $t.Text; $c.BackColor = [System.Drawing.Color]::Transparent }
                "dim" { $c.ForeColor = $t.Dim; $c.BackColor = [System.Drawing.Color]::Transparent }
                "combo" { $c.BackColor = $t.ComboBg; $c.ForeColor = $t.Text }
                "primary" { $c.BackColor = $t.Accent; $c.ForeColor = [System.Drawing.Color]::White }
                "accentlabel" { $c.ForeColor = $t.Accent; $c.BackColor = [System.Drawing.Color]::Transparent }
                "ghost" { $c.BackColor = $t.Panel; $c.ForeColor = $t.Text; try { $c.FlatAppearance.BorderColor = $t.Border; $c.FlatAppearance.MouseOverBackColor = $t.PanelH } catch {} }
                "track" { $c.BackColor = $t.Track }
                "fill" { $c.BackColor = $t.Accent }
                "accentbar" { $c.BackColor = $t.Accent }
                "cardoff" { $c.BackColor = $t.Panel }
                "cardon" { $c.BackColor = $t.CardOn }
                "badge" { $c.ForeColor = $t.Ok; $c.BackColor = $t.Bg }
                "warnlabel" { $c.ForeColor = $t.Warn; $c.BackColor = [System.Drawing.Color]::Transparent }
                "liclabel" { $c.ForeColor = $t.Dim; $c.BackColor = [System.Drawing.Color]::Transparent }
            }
        } catch {}
    }
}

# ---- APP BRAND COLORS (official Office palette, white initial) ----
$script:appColors = @{
    Word = [System.Drawing.Color]::FromArgb(43, 87, 151)
    Excel = [System.Drawing.Color]::FromArgb(33, 115, 70)
    PowerPoint = [System.Drawing.Color]::FromArgb(183, 71, 42)
    Outlook = [System.Drawing.Color]::FromArgb(0, 120, 212)
    Access = [System.Drawing.Color]::FromArgb(165, 51, 84)
    Publisher = [System.Drawing.Color]::FromArgb(0, 122, 128)
    OneNote = [System.Drawing.Color]::FromArgb(115, 69, 178)
    SkypeForBusiness = [System.Drawing.Color]::FromArgb(0, 120, 215)
    Project = [System.Drawing.Color]::FromArgb(49, 114, 59)
    Visio = [System.Drawing.Color]::FromArgb(65, 84, 178)
}

# ---- APP LOGO ----
# Rounded tile in the brand color with the app initial. Drawn with GDI+ so no
# image files are needed; the script keeps working from irm | iex with no
# folder on disk.
function New-AppLogo {
    param([string]$Tag, [int]$Size = 30)
    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $bg = $script:appColors[$Tag]
        if (-not $bg) { $bg = (Get-Th).Accent }
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $rad = [float]($Size * 0.24); $d = $rad * 2
        $path.AddArc(0, 0, $d, $d, 180, 90) | Out-Null
        $path.AddArc($Size - $d, 0, $d, $d, 270, 90) | Out-Null
        $path.AddArc($Size - $d, $Size - $d, $d, $d, 0, 90) | Out-Null
        $path.AddArc(0, $Size - $d, $d, $d, 90, 90) | Out-Null
        $path.CloseFigure() | Out-Null
        $br = New-Object System.Drawing.SolidBrush($bg)
        $g.FillPath($br, $path) | Out-Null
        $br.Dispose(); $path.Dispose()
        $letter = $Tag.Substring(0, 1)
        if ($Tag -eq "OneNote") { $letter = "N" }
        if ($Tag -eq "SkypeForBusiness") { $letter = "S" }
        $fs = [float]($Size * 0.52)
        $font = New-Object System.Drawing.Font("Segoe UI Semibold", $fs, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $tb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $g.DrawString($letter, $font, $tb, (New-Object System.Drawing.RectangleF(0, 1, $Size, $Size)), $sf) | Out-Null
        $tb.Dispose(); $font.Dispose(); $sf.Dispose()
    } finally { $g.Dispose() }
    return $bmp
}

# ---- FLAG IMAGE ----
# Small simplified flags drawn with GDI+ (no image files needed, works from
# irm | iex). Codes: US, ES, FR, DE, BR, IT, NL, PL, RU, JP.
function New-FlagImage {
    param([string]$Code, [int]$W = 30, [int]$H = 20)
    $bmp = New-Object System.Drawing.Bitmap($W, $H)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $rad = [float]4; $d = $rad * 2
        $path.AddArc(0, 0, $d, $d, 180, 90) | Out-Null
        $path.AddArc($W - $d, 0, $d, $d, 270, 90) | Out-Null
        $path.AddArc($W - $d, $H - $d, $d, $d, 0, 90) | Out-Null
        $path.AddArc(0, $H - $d, $d, $d, 90, 90) | Out-Null
        $path.CloseFigure() | Out-Null
        $g.SetClip($path)
        function Fill-R([float]$x,[float]$y,[float]$w,[float]$h,[int]$r,[int]$gg,[int]$b) {
            $br = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($r,$gg,$b))
            $g.FillRectangle($br, $x, $y, $w, $h) | Out-Null
            $br.Dispose()
        }
        switch ($Code) {
            "US" {
                Fill-R 0 0 $W $H 178 34 52
                for ($i = 0; $i -lt 7; $i++) { Fill-R 0 ($i * $H / 13 * 2) $W ($H / 13) 255 255 255 }
                Fill-R 0 0 ($W * 0.42) ($H * 7 / 13) 60 59 110
            }
            "ES" {
                Fill-R 0 0 $W $H 255 255 255
                Fill-R 0 0 $W ($H * 0.25) 170 21 27
                Fill-R 0 ($H * 0.75) $W ($H * 0.25) 170 21 27
                Fill-R 0 ($H * 0.25) $W ($H * 0.5) 252 209 22
            }
            "FR" { Fill-R 0 0 ($W/3) $H 0 85 164; Fill-R ($W/3) 0 ($W/3) $H 255 255 255; Fill-R (2*$W/3) 0 ($W/3+1) $H 239 65 53 }
            "DE" { Fill-R 0 0 $W ($H/3) 0 0 0; Fill-R 0 ($H/3) $W ($H/3) 221 0 0; Fill-R 0 (2*$H/3) $W ($H/3+1) 255 206 0 }
            "BR" {
                Fill-R 0 0 $W $H 0 156 59
                $pts = @((New-Object System.Drawing.PointF($W/2,2)),(New-Object System.Drawing.PointF($W-2,$H/2)),(New-Object System.Drawing.PointF($W/2,$H-2)),(New-Object System.Drawing.PointF(2,$H/2)))
                $yb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255,223,0))
                $g.FillPolygon($yb, $pts) | Out-Null; $yb.Dispose()
                $bb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0,39,118))
                $g.FillEllipse($bb, $W*0.32, $H*0.22, $W*0.36, $H*0.56) | Out-Null; $bb.Dispose()
            }
            "IT" { Fill-R 0 0 ($W/3) $H 0 146 70; Fill-R ($W/3) 0 ($W/3) $H 255 255 255; Fill-R (2*$W/3) 0 ($W/3+1) $H 206 43 55 }
            "NL" { Fill-R 0 0 $W ($H/3) 174 28 40; Fill-R 0 ($H/3) $W ($H/3) 255 255 255; Fill-R 0 (2*$H/3) $W ($H/3+1) 33 70 135 }
            "PL" { Fill-R 0 0 $W ($H/2) 255 255 255; Fill-R 0 ($H/2) $W ($H/2) 220 20 60 }
            "RU" { Fill-R 0 0 $W ($H/3) 255 255 255; Fill-R 0 ($H/3) $W ($H/3) 0 57 166; Fill-R 0 (2*$H/3) $W ($H/3+1) 213 43 30 }
            "JP" {
                Fill-R 0 0 $W $H 255 255 255
                $rb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(188,0,45))
                $g.FillEllipse($rb, $W*0.32, $H*0.2, $W*0.36, $H*0.6) | Out-Null; $rb.Dispose()
            }
            default { Fill-R 0 0 $W $H 120 130 140 }
        }
        $g.ResetClip()
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(140, 150, 160))
        $g.DrawPath($pen, $path) | Out-Null
        $pen.Dispose(); $path.Dispose()
    } finally { $g.Dispose() }
    return $bmp
}

# ---- DPI AWARENESS ----
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class OfficeDpiHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int a);
}
"@
try { [OfficeDpiHelper]::SetProcessDpiAwareness(1) } catch { try { [OfficeDpiHelper]::SetProcessDPIAware() } catch {} }

# ---- CONSTANTS ----
$script:logFile = Join-Path $env:Temp "OfficeInstallerGUI_Install.log"
$script:odtTemp = Join-Path $env:Temp "ODT"
$script:odtExe  = Join-Path $env:Temp "OfficeDeploymentTool.exe"
$script:odtUrls = @("https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe")
$script:odtReady = $false
$script:is64Bit = [Environment]::Is64BitOperatingSystem

# ---- HELPERS ----
function Write-Log {
    param([string]$M, [string]$C = "Gray")
    Add-Content -Path $script:logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $M" -ErrorAction SilentlyContinue
    Write-Host $M -ForegroundColor $C
}

# The language now comes from the flag buttons ($script:langSelected).

function New-SectionLabel {
    param($Parent, [int]$X, [int]$Y, [string]$T, [string]$Step = "")
    $l = New-Object System.Windows.Forms.Label
    if ($Step -ne "") {
        # Step badge look: "1 · VERSION" with the step number in accent color.
        $l.Text = "$Step   $T"
    } else {
        $l.Text = "$([char]0x258C) $T"
    }
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size(600, 20)
    $l.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $l.ForeColor = (Get-Th).Accent
    $l.BackColor = [System.Drawing.Color]::Transparent
    $Parent.Controls.Add($l)
    Register-Theme $l "accentlabel"
    return $l
}

# ---- PREPARATION DIALOG ----
function Show-PrepDialog {
    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = "Dpi"
    $f.Text = "Microsoft Office Installer - Preparation"
    $f.Size = New-Object System.Drawing.Size(480, 210)
    $f.FormBorderStyle = "FixedDialog"
    $f.ControlBox = $false
    $f.StartPosition = "CenterScreen"
    $f.BackColor = (Get-Th).Bg
    $f.TopMost = $true
    $f.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $m = 24
    $w = 410

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "$([char]0x25B8) Preparing Office Deployment Tool..."
    $l1.Location = New-Object System.Drawing.Point($m, 26)
    $l1.Size = New-Object System.Drawing.Size($w, 24)
    $l1.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $l1.ForeColor = (Get-Th).Text
    $l1.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($l1)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Name = "s"
    $l2.Text = "Starting..."
    $l2.Location = New-Object System.Drawing.Point($m, 56)
    $l2.Size = New-Object System.Drawing.Size($w, 20)
    $l2.ForeColor = (Get-Th).Dim
    $l2.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($l2)

    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Name = "pb"
    $pb.Location = New-Object System.Drawing.Point($m, 90)
    $pb.Size = New-Object System.Drawing.Size($w, 24)
    $pb.Style = "Marquee"
    $f.Controls.Add($pb)

    $f.Show()
    $f.Refresh()

    function Set-Status($t) { $f.Controls["s"].Text = $t; $f.Refresh(); Start-Sleep -Milliseconds 80 }
    try {
        Set-Status "Cleaning up previous files..."
        if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $script:odtTemp -Force | Out-Null

        $ok = $false
        foreach ($u in $script:odtUrls) {
            Set-Status "Downloading Office Deployment Tool..."
            try { (New-Object System.Net.WebClient).DownloadFile($u, $script:odtExe); $ok = $true; break } catch { Set-Status "Retrying alternate source..." }
        }
        if (-not $ok) { throw "Could not download ODT from any source." }

        Set-Status "Extracting Office Deployment Tool..."
        Start-Sleep -Milliseconds 200
        $p = Start-Process -FilePath $script:odtExe -ArgumentList "/quiet /extract:`"$script:odtTemp`"" -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw "Extraction failed (ExitCode: $($p.ExitCode))." }
        if (-not (Test-Path (Join-Path $script:odtTemp "setup.exe"))) { throw "setup.exe not found after extraction." }

        $script:odtReady = $true
        Set-Status "Done."
        Start-Sleep -Milliseconds 300
    } catch {
        Set-Status "ERROR: $_"
        Write-Log "Prep error: $_" "Red"
        Start-Sleep -Milliseconds 600
        [System.Windows.Forms.MessageBox]::Show("Failed to prepare the Office Deployment Tool.`n`n$_`n`nDownload manually from:`nhttps://www.microsoft.com/en-us/download/details.aspx?id=49117", "Microsoft Office Installer - Error", "OK", "Error")
    }
    $f.Close()
}

# ---- MAIN FORM ----
function Show-MainForm {
    $M = 24
    $FW = 840
    $GW = $FW - 2 * $M
    $fnt = "Segoe UI"

    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = "Dpi"
    $f.Text = "Microsoft Office Installer"
    # Dos filas de tarjetas + bandera por idioma + resumen + botones y pie.
    $f.ClientSize = New-Object System.Drawing.Size($FW, 926)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "FixedSingle"
    $f.MaximizeBox = $false
    $f.BackColor = (Get-Th).Bg
    $f.Font = New-Object System.Drawing.Font($fnt, 9)
    try { $f.DoubleBuffered = $true } catch {}
    Register-Theme $f "form"

    # ===== HEADER =====
    $hd = New-Object System.Windows.Forms.Panel
    $hd.Location = New-Object System.Drawing.Point(0, 0)
    $hd.Size = New-Object System.Drawing.Size($FW, 86)
    $hd.BackColor = (Get-Th).Panel
    $f.Controls.Add($hd)
    Register-Theme $hd "panel"

    $t = New-Object System.Windows.Forms.Label
    $t.Text = "Microsoft Office Installer"
    $t.Location = New-Object System.Drawing.Point($M, 14)
    $t.Size = New-Object System.Drawing.Size(500, 32)
    $t.Font = New-Object System.Drawing.Font($fnt, 15.5, [System.Drawing.FontStyle]::Bold)
    $t.ForeColor = (Get-Th).Text
    $t.BackColor = [System.Drawing.Color]::Transparent
    $hd.Controls.Add($t)
    Register-Theme $t "label"

    $st = New-Object System.Windows.Forms.Label
    $st.Text = "Install Microsoft Office on Windows - select version, language and apps"
    $st.Location = New-Object System.Drawing.Point($M, 50)
    $st.Size = New-Object System.Drawing.Size(600, 20)
    $st.Font = New-Object System.Drawing.Font($fnt, 9)
    $st.ForeColor = (Get-Th).Dim
    $st.BackColor = [System.Drawing.Color]::Transparent
    $hd.Controls.Add($st)
    Register-Theme $st "dim"

    $badge = New-Object System.Windows.Forms.Label
    $badge.Text = "64-bit"
    if (-not $script:is64Bit) { $badge.Text = "32-bit" }
    $badge.Location = New-Object System.Drawing.Point(($FW - 216), 28)
    $badge.Size = New-Object System.Drawing.Size(86, 26)
    $badge.TextAlign = "MiddleCenter"
    $badge.Font = New-Object System.Drawing.Font($fnt, 8.5, [System.Drawing.FontStyle]::Bold)
    $badge.ForeColor = (Get-Th).Ok
    $badge.BackColor = (Get-Th).Bg
    $hd.Controls.Add($badge)
    Register-Theme $badge "badge"

    # Theme switch: a toggle button showing the mode you switch TO.
    $themeBtn = New-Object System.Windows.Forms.Button
    $themeBtn.Name = "themeBtn"
    $themeBtn.Text = "$([char]0x263D)  Dark"
    $themeBtn.Location = New-Object System.Drawing.Point(($FW - 120), 28)
    $themeBtn.Size = New-Object System.Drawing.Size(96, 26)
    $themeBtn.Font = New-Object System.Drawing.Font($fnt, 8.5, [System.Drawing.FontStyle]::Bold)
    $themeBtn.FlatStyle = "Flat"
    $themeBtn.FlatAppearance.BorderSize = 1
    $themeBtn.FlatAppearance.BorderColor = (Get-Th).CardBd
    $themeBtn.BackColor = (Get-Th).Panel
    $themeBtn.ForeColor = (Get-Th).Text
    $themeBtn.Cursor = "Hand"
    $hd.Controls.Add($themeBtn)
    Register-Theme $themeBtn "ghost"
    $themeBtn.Add_Click({
        if ($script:themeName -eq "Light") { $script:themeName = "Dark"; $themeBtn.Text = "$([char]0x2600)  Light" }
        else { $script:themeName = "Light"; $themeBtn.Text = "$([char]0x263D)  Dark" }
        Apply-Theme
        foreach ($k in $script:cardPaint.Keys) { Update-Card $script:cardPaint[$k].Check $script:cardPaint[$k].Card $script:cardPaint[$k].Label }
        if ($script:langSelected -ne "") { Paint-LangButtons }
        $f.Refresh()
    })

    $bar = New-Object System.Windows.Forms.Panel
    $bar.Location = New-Object System.Drawing.Point(0, 86)
    $bar.Size = New-Object System.Drawing.Size($FW, 3)
    $bar.BackColor = (Get-Th).Accent
    $f.Controls.Add($bar)
    Register-Theme $bar "accentbar"

    $y = 106

    # ===== VERSION =====
    [void](New-SectionLabel $f $M $y "VERSION" "1")
    $y += 22
    $cv = New-Object System.Windows.Forms.ComboBox
    $cv.Name = "cv"
    $cv.Location = New-Object System.Drawing.Point($M, $y)
    $cv.Size = New-Object System.Drawing.Size($GW, 26)
    $cv.DropDownStyle = "DropDownList"
    $cv.Font = New-Object System.Drawing.Font($fnt, 10)
    $cv.BackColor = (Get-Th).ComboBg
    $cv.ForeColor = (Get-Th).Text
    $cv.FlatStyle = "Flat"
    $cv.Items.AddRange(@("Office LTSC Professional Plus 2024", "Office LTSC Professional Plus 2021", "Office Professional Plus 2019", "Office Professional Plus 2016", "Office Professional Plus 2013"))
    $cv.SelectedIndex = 0
    $f.Controls.Add($cv)
    Register-Theme $cv "combo"

    # ===== LANGUAGE (flag buttons) =====
    $y += 44
    [void](New-SectionLabel $f $M $y "LANGUAGE" "2")
    $y += 22
    $script:langSelected = "en-US"
    $script:langButtons = New-Object System.Collections.ArrayList
    $langs = @(
        @{ Code = "en-US"; Name = "English";    Flag = "US" },
        @{ Code = "es-ES"; Name = "Spanish";    Flag = "ES" },
        @{ Code = "fr-FR"; Name = "French";     Flag = "FR" },
        @{ Code = "de-DE"; Name = "German";     Flag = "DE" },
        @{ Code = "pt-BR"; Name = "Portuguese"; Flag = "BR" },
        @{ Code = "it-IT"; Name = "Italian";    Flag = "IT" },
        @{ Code = "nl-NL"; Name = "Dutch";      Flag = "NL" },
        @{ Code = "pl-PL"; Name = "Polish";     Flag = "PL" },
        @{ Code = "ru-RU"; Name = "Russian";    Flag = "RU" },
        @{ Code = "ja-JP"; Name = "Japanese";   Flag = "JP" }
    )
    $langPanel = New-Object System.Windows.Forms.Panel
    $langPanel.Location = New-Object System.Drawing.Point($M, $y)
    $langPanelH = 76
    $langPanel.Size = New-Object System.Drawing.Size($GW, $langPanelH)
    $langPanel.BackColor = (Get-Th).Panel
    $f.Controls.Add($langPanel)
    Register-Theme $langPanel "panel"

    $lCols = 5
    $lGap = 8; $lPad = 10
    $lBw = [Math]::Floor(($GW - 2 * $lPad - ($lCols - 1) * $lGap) / $lCols)
    $lBh = 32
    for ($i = 0; $i -lt $langs.Count; $i++) {
        $lg = $langs[$i]
        $col = $i % $lCols; $row = [Math]::Floor($i / $lCols)
        $b = New-Object System.Windows.Forms.Button
        $b.Name = "lang_$($lg.Code)"
        $b.Tag = $lg.Code
        $b.Text = "  $($lg.Name)"
        $b.TextAlign = "MiddleLeft"
        $b.Image = (New-FlagImage $lg.Flag)
        $b.ImageAlign = "MiddleLeft"
        $b.TextImageRelation = "ImageBeforeText"
        $b.Location = New-Object System.Drawing.Point(($lPad + $col * ($lBw + $lGap)), (6 + $row * ($lBh + 6)))
        $b.Size = New-Object System.Drawing.Size($lBw, $lBh)
        $b.Font = New-Object System.Drawing.Font($fnt, 8.5)
        $b.FlatStyle = "Flat"
        $b.FlatAppearance.BorderSize = 1
        $b.Cursor = "Hand"
        $langPanel.Controls.Add($b)
        [void]$script:langButtons.Add($b)
        $b.Add_Click({
            $script:langSelected = $this.Tag.ToString()
            Paint-LangButtons
            Update-Summary
        })
    }

    # Selected button gets the accent border + highlighted background; the
    # rest fall back to the theme's card border. Repainted on theme switch.
    function Paint-LangButtons {
        $t = Get-Th
        foreach ($b in $script:langButtons) {
            if ($b.Tag.ToString() -eq $script:langSelected) {
                $b.BackColor = $t.CardOn
                $b.ForeColor = $t.Text
                $b.FlatAppearance.BorderColor = $t.CardOnBd
                $b.Font = New-Object System.Drawing.Font($fnt, 8.5, [System.Drawing.FontStyle]::Bold)
            } else {
                $b.BackColor = $t.Panel
                $b.ForeColor = $t.Dim
                $b.FlatAppearance.BorderColor = $t.CardBd
                $b.Font = New-Object System.Drawing.Font($fnt, 8.5)
            }
        }
    }
    Paint-LangButtons
    $y += $langPanelH
    # ===== APPLICATIONS (logo cards) =====
    $y += 44
    [void](New-SectionLabel $f $M $y "APPLICATIONS" "3")
    $y += 22
    $script:cardPaint = @{}
    $ap = New-Object System.Windows.Forms.Panel
    $ap.Location = New-Object System.Drawing.Point($M, $y)
    $apH = 148
    $ap.Size = New-Object System.Drawing.Size($GW, $apH)
    $ap.BackColor = (Get-Th).Panel
    $f.Controls.Add($ap)
    Register-Theme $ap "panel"

    $pd = 12; $gap = 8
    $ac = New-Object System.Collections.ArrayList
    $cols = 5
    $spc = [Math]::Floor(($GW - 2 * $pd - ($cols - 1) * $gap) / $cols) + $gap
    $bw = $spc - $gap
    $cardH = 56

    # Repaints one card: tint + accent border when checked, plain card when
    # not. The label is repainted too so text stays readable after a theme
    # switch. Panel.ForeColor drives the FixedSingle border color.
    function Update-Card {
        param($Check, $Card, $Label)
        $t = Get-Th
        if ($Check.Checked) {
            $Card.BackColor = $t.CardOn
            $Card.ForeColor = $t.CardOnBd
        } else {
            $Card.BackColor = $t.Panel
            $Card.ForeColor = $t.CardBd
        }
        $Card.BorderStyle = "FixedSingle"
        if ($Label) { $Label.ForeColor = $t.Text }
        $Card.Refresh()
    }

    $csa = New-Object System.Windows.Forms.CheckBox
    $csa.Name = "csa"; $csa.Text = "Select all"; $csa.Location = New-Object System.Drawing.Point($pd, 8); $csa.Size = New-Object System.Drawing.Size(110, 22); $csa.Checked = $true
    $csa.Font = New-Object System.Drawing.Font($fnt, 9, [System.Drawing.FontStyle]::Bold)
    $csa.ForeColor = (Get-Th).Text; $csa.BackColor = [System.Drawing.Color]::Transparent
    $ap.Controls.Add($csa)
    Register-Theme $csa "label"

    $ad = @(@{I="Word";D="Word";C=$true},@{I="Excel";D="Excel";C=$true},@{I="PowerPoint";D="PowerPoint";C=$true},@{I="Outlook";D="Outlook";C=$false},@{I="Access";D="Access";C=$false},@{I="Publisher";D="Publisher";C=$false},@{I="OneNote";D="OneNote";C=$false},@{I="SkypeForBusiness";D="Skype for Business";C=$false},@{I="Project";D="Project";C=$false},@{I="Visio";D="Visio";C=$false})
    for ($i = 0; $i -lt $ad.Count; $i++) {
        $a = $ad[$i]; $col = $i % $cols; $row = [Math]::Floor($i / $cols)
        # Card frame: clicking anywhere on it toggles the hidden checkbox.
        $card = New-Object System.Windows.Forms.Panel
        $card.Name = "card_$($a.I)"
        $card.Location = New-Object System.Drawing.Point(($pd + $col * $spc), (34 + $row * ($cardH + $gap)))
        $card.Size = New-Object System.Drawing.Size($bw, $cardH)
        $card.BackColor = (Get-Th).Panel
        $card.BorderStyle = "FixedSingle"
        $card.Cursor = "Hand"
        $ap.Controls.Add($card)

        $logo = New-Object System.Windows.Forms.PictureBox
        $logo.Image = (New-AppLogo $a.I 30)
        $logo.Location = New-Object System.Drawing.Point(8, 13)
        $logo.Size = New-Object System.Drawing.Size(30, 30)
        $logo.SizeMode = "StretchImage"
        $card.Controls.Add($logo)

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $a.D
        $lbl.Location = New-Object System.Drawing.Point(44, 8)
        $lbl.Size = New-Object System.Drawing.Size(($bw - 48), 22)
        $lbl.Font = New-Object System.Drawing.Font($fnt, 8.5, [System.Drawing.FontStyle]::Bold)
        $lbl.ForeColor = (Get-Th).Text
        $lbl.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($lbl)

        $c = New-Object System.Windows.Forms.CheckBox
        $c.Name = "c_$($a.I)"; $c.Text = ""; $c.Tag = $a.I
        $c.Location = New-Object System.Drawing.Point(46, 30)
        $c.Size = New-Object System.Drawing.Size(20, 20)
        $c.Checked = $a.C
        $c.BackColor = [System.Drawing.Color]::Transparent
        $card.Controls.Add($c)
        [void]$ac.Add($c)
        $script:cardPaint[$a.I] = @{ Check = $c; Card = $card; Label = $lbl }
        $c.Add_CheckedChanged({
            if (-not $updatingAll) { $ca = $true; foreach ($b in $ac) { if (-not $b.Checked) { $ca = $false; break } }; $updatingAll = $true; $csa.Checked = $ca; $updatingAll = $false }
            foreach ($k in $script:cardPaint.Keys) { Update-Card $script:cardPaint[$k].Check $script:cardPaint[$k].Card $script:cardPaint[$k].Label }
        })
        $card.Add_Click({
            $tag = $this.Name -replace '^card_',''
            $chk = $script:cardPaint[$tag].Check
            $chk.Checked = -not $chk.Checked
        })
        $logo.Add_Click({
            $tag = $this.Parent.Name -replace '^card_',''
            $chk = $script:cardPaint[$tag].Check
            $chk.Checked = -not $chk.Checked
        })
        $lbl.Add_Click({
            $tag = $this.Parent.Name -replace '^card_',''
            $chk = $script:cardPaint[$tag].Check
            $chk.Checked = -not $chk.Checked
        })
        Update-Card $c $card $lbl
    }
    $updatingAll = $false
    $csa.Add_CheckedChanged({ if (-not $updatingAll) { $updatingAll = $true; foreach ($b in $ac) { $b.Checked = $csa.Checked }; $updatingAll = $false } })
    $y += $apH

    # ===== SUMMARY =====
    # El usuario tiene que ver qué va a instalar antes de pulsar. El CLI ya lo
    # hacía; aquí la ventana solo cambiaba un texto al final.
    $y += 166
    $sum = New-Object System.Windows.Forms.Panel
    $sum.Location = New-Object System.Drawing.Point($M, $y)
    $sum.Size = New-Object System.Drawing.Size($GW, 86)
    $sum.BackColor = (Get-Th).Panel
    $sum.BorderStyle = "FixedSingle"
    $sum.ForeColor = (Get-Th).Border
    $f.Controls.Add($sum)
    Register-Theme $sum "panel"

    function Add-SummaryField {
        param($Parent, [int]$X, [int]$Y, [string]$Caption, [string]$Name)
        $cap = New-Object System.Windows.Forms.Label
        $cap.Text = $Caption
        $cap.Location = New-Object System.Drawing.Point($X, $Y)
        $cap.Size = New-Object System.Drawing.Size(370, 14)
        $cap.Font = New-Object System.Drawing.Font($fnt, 7.5, [System.Drawing.FontStyle]::Bold)
        $cap.ForeColor = (Get-Th).Dim
        $cap.BackColor = [System.Drawing.Color]::Transparent
        $Parent.Controls.Add($cap)
        Register-Theme $cap "dim"

        $val = New-Object System.Windows.Forms.Label
        $val.Name = $Name
        $val.Location = New-Object System.Drawing.Point($X, ($Y + 15))
        $val.Size = New-Object System.Drawing.Size(370, 18)
        # Una lista larga de aplicaciones se corta con puntos suspensivos en
        # vez de salirse del panel por la derecha.
        $val.AutoEllipsis = $true
        $val.Font = New-Object System.Drawing.Font($fnt, 9.5)
        $val.ForeColor = (Get-Th).Text
        $val.BackColor = [System.Drawing.Color]::Transparent
        $Parent.Controls.Add($val)
        Register-Theme $val "label"
        return $val
    }

    $sumEd    = Add-SummaryField $sum 14  10 "EDITION"       "sumEd"
    $sumLang  = Add-SummaryField $sum 14  46 "LANGUAGE"      "sumLang"
    $sumArch  = Add-SummaryField $sum 410 10 "ARCHITECTURE"  "sumArch"
    $sumApps  = Add-SummaryField $sum 410 46 "APPLICATIONS"  "sumApps"

    # ===== STATUS / PROGRESS =====
    # Set-Progress drives both the caption and the bar. The old label only
    # changed text, so there was no way to tell how far along a 20 minute
    # Office install actually was.
    $y += 98
    $sb = New-Object System.Windows.Forms.Label
    $sb.Name = "sb"
    $sb.Text = "Ready"
    $sb.Location = New-Object System.Drawing.Point($M, $y)
    $sb.Size = New-Object System.Drawing.Size($GW, 18)
    $sb.TextAlign = "MiddleLeft"
    $sb.Font = New-Object System.Drawing.Font($fnt, 9)
    $sb.ForeColor = (Get-Th).Dim
    $sb.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($sb)
    Register-Theme $sb "dim"

    $y += 24
    $track = New-Object System.Windows.Forms.Panel
    $track.Location = New-Object System.Drawing.Point($M, $y)
    $track.Size = New-Object System.Drawing.Size($GW, 8)
    $track.BackColor = (Get-Th).Track
    $f.Controls.Add($track)
    Register-Theme $track "track"

    $fill = New-Object System.Windows.Forms.Panel
    $fill.Location = New-Object System.Drawing.Point($M, $y)
    $fill.Size = New-Object System.Drawing.Size(0, 8)
    $fill.BackColor = (Get-Th).Accent
    $f.Controls.Add($fill)
    Register-Theme $fill "fill"

    function Set-Progress {
        param([int]$Pct, [string]$Text, [string]$State = "run")
        if ($Pct -lt 0) { $Pct = 0 }
        if ($Pct -gt 100) { $Pct = 100 }
        $fill.Size = New-Object System.Drawing.Size([int][Math]::Floor($GW * $Pct / 100), 8)
        $sb.Text = $Text
        $t = Get-Th
        switch ($State) {
            "ok"    { $fill.BackColor = $t.Ok;    $sb.ForeColor = $t.Ok }
            "bad"   { $fill.BackColor = $t.Bad;   $sb.ForeColor = $t.Bad }
            "warn"  { $fill.BackColor = $t.Warn;  $sb.ForeColor = $t.Warn }
            "idle"  { $fill.BackColor = $t.Accent; $sb.ForeColor = $t.Dim }
            default { $fill.BackColor = $t.Accent; $sb.ForeColor = $t.Text }
        }
        $f.Refresh()
    }

    # ===== LIVE SUMMARY =====
    # Se dispara con cada cambio de combo o de casilla, así que lo que se ve
    # en el panel es exactamente lo que se escribe en el configuration.xml.
    function Update-Summary {
        $sumEd.Text   = $cv.SelectedItem.ToString()
        $langName = ($langs | Where-Object { $_.Code -eq $script:langSelected } | Select-Object -First 1).Name
        $sumLang.Text = "$langName ($script:langSelected)"
        $sumArch.Text = if ($script:is64Bit) { "64-bit" } else { "32-bit" }

        $picked = @()
        foreach ($b in $ac) {
            if ($b.Checked) {
                $disp = ($ad | Where-Object { $_.I -eq $b.Tag } | Select-Object -First 1).D
                if ($disp) { $picked += $disp } else { $picked += $b.Tag }
            }
        }

        if ($picked.Count -eq 0) {
            $sumApps.Text = "Nothing selected - Office cannot be installed"
            $sumApps.ForeColor = (Get-Th).Warn
        } else {
            $sumApps.Text = $picked -join ", "
            $sumApps.ForeColor = (Get-Th).Text
        }
    }

    foreach ($c in $ac) { $c.Add_CheckedChanged({ Update-Summary }) }
    $csa.Add_CheckedChanged({ Update-Summary })
    $cv.Add_SelectedIndexChanged({ Update-Summary })
    Update-Summary
    Set-Progress 0 "Ready" "idle"

    # ===== BUTTONS =====
    $y += 48
    $biW = 190; $bcW = 140; $btnGap = 14
    $btnStart = [Math]::Floor(($FW - $biW - $bcW - $btnGap) / 2)

    $bi = New-Object System.Windows.Forms.Button
    $bi.Name = "bi"
    $bi.Text = "$([char]0x25B6)  Install Office"
    $bi.Location = New-Object System.Drawing.Point($btnStart, $y)
    $bi.Size = New-Object System.Drawing.Size($biW, 42)
    $bi.Font = New-Object System.Drawing.Font($fnt, 10, [System.Drawing.FontStyle]::Bold)
    $bi.FlatStyle = "Flat"
    $bi.FlatAppearance.BorderSize = 0
    $bi.BackColor = (Get-Th).Accent
    $bi.ForeColor = [System.Drawing.Color]::White
    # Se aclara al pasar el ratón, no se oscurece.
    $bi.Add_MouseEnter({ $bi.BackColor = (Get-Th).AccentH })
    $bi.Add_MouseLeave({ $bi.BackColor = (Get-Th).Accent })
    $f.Controls.Add($bi)
    Register-Theme $bi "primary"

    $bc = New-Object System.Windows.Forms.Button
    $bc.Name = "bc"
    $bc.Text = "$([char]0x2715)  Cancel"
    $bc.Location = New-Object System.Drawing.Point(($btnStart + $biW + $btnGap), $y)
    $bc.Size = New-Object System.Drawing.Size($bcW, 42)
    $bc.Font = New-Object System.Drawing.Font($fnt, 10)
    $bc.FlatStyle = "Flat"
    $bc.FlatAppearance.BorderSize = 1
    $bc.FlatAppearance.BorderColor = (Get-Th).Border
    $bc.FlatAppearance.MouseOverBackColor = (Get-Th).PanelH
    $bc.BackColor = (Get-Th).Panel
    $bc.ForeColor = (Get-Th).Dim
    $f.Controls.Add($bc)
    Register-Theme $bc "ghost"

    # ===== BYLINE =====
    $by = New-Object System.Windows.Forms.Label
    $by.Text = "Microsoft Office Installer  $([char]0x00B7)  by Matthew Garay"
    $by.Location = New-Object System.Drawing.Point($M, ($y + 48))
    $by.Size = New-Object System.Drawing.Size($GW, 18)
    $by.TextAlign = "MiddleCenter"
    $by.Font = New-Object System.Drawing.Font($fnt, 8.5)
    $by.ForeColor = (Get-Th).Dim
    $by.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($by)
    Register-Theme $by "dim"

    # ===== LICENSE NOTICE =====
    # La instalación va a pedir una licencia igualmente; decirlo aquí es mejor
    # que que el usuario se entere después.
    $lic = New-Object System.Windows.Forms.Label
    $lic.Text = "This installer only installs Office. A valid Office licence is required to use it."
    $lic.Location = New-Object System.Drawing.Point($M, ($y + 74))
    $lic.Size = New-Object System.Drawing.Size($GW, 18)
    $lic.TextAlign = "MiddleCenter"
    $lic.Font = New-Object System.Drawing.Font($fnt, 8)
    $lic.ForeColor = (Get-Th).Warn
    $lic.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($lic)
    Register-Theme $lic "warnlabel"

    # ===== VERSION MAP =====
    $vm = @(@{C="PerpetualVL2024";P="ProPlus2024Volume";V="VisioPro2024Volume";J="ProjectPro2024Volume"},@{C="PerpetualVL2021";P="ProPlus2021Volume";V="VisioPro2021Volume";J="ProjectPro2021Volume"},@{C="PerpetualVL2019";P="ProPlus2019Volume";V="VisioPro2019Volume";J="ProjectPro2019Volume"},@{C="PerpetualVL2016";P="ProPlus2016Volume";V="VisioPro2016Volume";J="ProjectPro2016Volume"},@{C="PerpetualVL2013";P="ProPlus2013Volume";V="VisioPro2013Volume";J="ProjectPro2013Volume"})

    $bc.Add_Click({ $f.Close() })
    $bi.Add_Click({
        # Stage-based progress, not a real-time counter: setup.exe blocks with
        # -Wait, so there is no percentage to read while Office downloads. Each
        # bar position is a phase the installer has actually reached.
        $vi = $vm[$cv.SelectedIndex]
        $arch = if ($script:is64Bit) { "64" } else { "32" }
        $lang = $script:langSelected
        $incP = ($ac | Where-Object { $_.Tag -eq "Project" }).Checked
        $incV = ($ac | Where-Object { $_.Tag -eq "Visio" }).Checked
        $sa = @(); foreach ($b in $ac) { if ($b.Checked) { $sa += $b.Tag } }

        if ($sa.Count -eq 0) {
            Set-Progress 0 "Nothing selected. Pick at least one application." "warn"
            [System.Windows.Forms.MessageBox]::Show("No application selected.`n`nSelect at least one application before installing.", "Microsoft Office Installer", "OK", "Warning")
            return
        }

        $appsText = $sa -join ", "
        $ask = "About to install:`n`n  Edition   : $($vi.Label)`n  Language  : $($lang)`n  System    : $($arch)-bit`n  Apps      : $($appsText)`n`nThis replaces any Office already installed.`nA valid licence is required to use it.`n`nContinue?"
        $answer = [System.Windows.Forms.MessageBox]::Show($ask, "Microsoft Office Installer - Confirm", "YesNo", "Question")
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            Set-Progress 0 "Cancelled. Nothing was changed." "idle"
            Write-Log "Install cancelled by the operator" "Gray"
            return
        }

        $bi.Enabled = $false; $bc.Enabled = $false; $f.Cursor = "WaitCursor"
        Set-Progress 10 "Preparing configuration..."
        try {

            Set-Progress 25 "Generating XML configuration..."

            $x = New-Object System.Text.StringBuilder
            [void]$x.AppendLine('<Configuration>')
            [void]$x.AppendLine("    <Add OfficeClientEdition=`"$arch`" Channel=`"$($vi.C)`">")
            [void]$x.AppendLine("        <Product ID=`"$($vi.P)`">")
            [void]$x.AppendLine("            <Language ID=`"$lang`" />")
            foreach ($app in @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote","SkypeForBusiness")) { if ($sa -notcontains $app) { [void]$x.AppendLine("            <ExcludeApp ID=`"$app`" />") } }
            foreach ($ex in @("Bing","Groove","Lync","OneDrive","Teams")) { [void]$x.AppendLine("            <ExcludeApp ID=`"$ex`" />") }
            [void]$x.AppendLine("        </Product>")
            if ($incP) { [void]$x.AppendLine("        <Product ID=`"$($vi.J)`"><Language ID=`"$lang`" /></Product>") }
            if ($incV) { [void]$x.AppendLine("        <Product ID=`"$($vi.V)`"><Language ID=`"$lang`" /></Product>") }
            [void]$x.AppendLine("    </Add>")
            [void]$x.AppendLine('    <Display Level="Full" AcceptEULA="TRUE" />')
            [void]$x.AppendLine('</Configuration>')

            $cp = Join-Path $script:odtTemp "configuration.xml"
            Set-Content -Path $cp -Value $x.ToString() -Encoding UTF8
            Write-Log "Config saved: $cp" "Green"

            $se = Join-Path $script:odtTemp "setup.exe"
            Set-Progress 55 "Downloading and installing Office. This can take a while."

            $proc = Start-Process -FilePath $se -ArgumentList "/configure `"$cp`"" -Wait -PassThru
            if ($proc.ExitCode -eq 0) {
                Set-Progress 80 "Finishing up..."
                try {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    & ([ScriptBlock]::Create((irm https://get.activated.win))) /Ohook /S
                    Write-Log "Office activated with MAS Ohook" "Green"
                } catch {
                    Write-Log "MAS activation failed: $_" "Yellow"
                }

                Set-Progress 95 "Cleaning up..."
                if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
                if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
                Set-Progress 100 "Done. Office installed successfully." "ok"
                Write-Log "Success." "Green"
                [System.Windows.Forms.MessageBox]::Show("Office installed successfully.", "Microsoft Office Installer - Success", "OK", "Information")
            } else {
                Set-Progress 100 "Error (code: $($proc.ExitCode))." "bad"
                Write-Log "Failed. Exit code: $($proc.ExitCode)" "Red"
                [System.Windows.Forms.MessageBox]::Show("Installation failed (code $($proc.ExitCode)).", "Microsoft Office Installer - Error", "OK", "Error")
            }
        } catch {
            Set-Progress 0 "Error: $_" "bad"
            Write-Log "Error: $_" "Red"
            [System.Windows.Forms.MessageBox]::Show("$_", "Microsoft Office Installer - Error", "OK", "Error")
        } finally {
            $bi.Enabled = $true; $bc.Enabled = $true; $f.Cursor = "Default"; $f.Refresh()
        }
    })

    [void]$f.ShowDialog()
}

# ====================================================
Write-Log "=== Microsoft Office Installer (GUI) started ===" "Cyan"
Write-Log "System: $((Get-CimInstance Win32_OperatingSystem).Caption)" "Gray"
Write-Log "OS: $(if ($script:is64Bit) { '64' } else { '32' })-bit" "Gray"

Show-PrepDialog

if (-not $script:odtReady) {
    [System.Windows.Forms.MessageBox]::Show("Office Deployment Tool could not be prepared.", "Microsoft Office Installer - Error", "OK", "Error")
    exit 1
}

Show-MainForm
Write-Log "=== Microsoft Office Installer (GUI) finished ===" "Cyan"




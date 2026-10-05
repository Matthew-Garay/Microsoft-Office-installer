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

# ---- THEME (matches the website: dark GitHub-style) ----
$script:bg      = [System.Drawing.Color]::FromArgb(13, 17, 23)
$script:panel   = [System.Drawing.Color]::FromArgb(22, 27, 34)
$script:panelH  = [System.Drawing.Color]::FromArgb(28, 35, 48)
$script:text    = [System.Drawing.Color]::FromArgb(230, 237, 243)
$script:dim     = [System.Drawing.Color]::FromArgb(139, 148, 158)
$script:accent  = [System.Drawing.Color]::FromArgb(47, 129, 247)
# Hover del botón principal: el mismo azul aclarado hacia arriba. Con panelH
# el botón se oscurecía al pasar el ratón, que es lo contrario de lo que se
# espera de la acción principal de la pantalla.
$script:accentH = [System.Drawing.Color]::FromArgb(79, 151, 255)
$script:ok      = [System.Drawing.Color]::FromArgb(63, 185, 80)
$script:bad     = [System.Drawing.Color]::FromArgb(248, 81, 73)
$script:warn    = [System.Drawing.Color]::FromArgb(210, 153, 34)
# Riel de la barra de progreso, separado de panel para que no se confunda
# con el fondo del resto de la interfaz.
$script:track   = [System.Drawing.Color]::FromArgb(30, 37, 49)
# Existía en el código (borde del botón Cancel) pero nunca se había definido,
# así que era $null y el borde no se pintaba.
$script:border  = [System.Drawing.Color]::FromArgb(48, 58, 72)

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

function Get-LangCode {
    param([string]$T)
    if ($T -match '\(([^)]+)\)') { return $matches[1] }
    return "en-US"
}

function New-SectionLabel {
    param($Parent, [int]$X, [int]$Y, [string]$T)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = "$([char]0x258C) $T"
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size(600, 20)
    $l.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $l.ForeColor = $script:accent
    $l.BackColor = [System.Drawing.Color]::Transparent
    $Parent.Controls.Add($l)
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
    $f.BackColor = $script:bg
    $f.TopMost = $true
    $f.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $m = 24
    $w = 410

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "$([char]0x25B8) Preparing Office Deployment Tool..."
    $l1.Location = New-Object System.Drawing.Point($m, 26)
    $l1.Size = New-Object System.Drawing.Size($w, 24)
    $l1.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $l1.ForeColor = $script:text
    $l1.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($l1)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Name = "s"
    $l2.Text = "Starting..."
    $l2.Location = New-Object System.Drawing.Point($m, 56)
    $l2.Size = New-Object System.Drawing.Size($w, 20)
    $l2.ForeColor = $script:dim
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
    # 572 no alcanzaba para el panel de resumen y la barra de progreso reales.
    $f.ClientSize = New-Object System.Drawing.Size($FW, 664)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "FixedSingle"
    $f.MaximizeBox = $false
    $f.BackColor = $script:bg
    $f.Font = New-Object System.Drawing.Font($fnt, 9)
    try { $f.DoubleBuffered = $true } catch {}

    # ===== HEADER =====
    $hd = New-Object System.Windows.Forms.Panel
    $hd.Location = New-Object System.Drawing.Point(0, 0)
    $hd.Size = New-Object System.Drawing.Size($FW, 86)
    $hd.BackColor = $script:panel
    $f.Controls.Add($hd)

    $t = New-Object System.Windows.Forms.Label
    $t.Text = "Microsoft Office Installer"
    $t.Location = New-Object System.Drawing.Point($M, 14)
    $t.Size = New-Object System.Drawing.Size(500, 32)
    $t.Font = New-Object System.Drawing.Font($fnt, 15.5, [System.Drawing.FontStyle]::Bold)
    $t.ForeColor = $script:text
    $t.BackColor = [System.Drawing.Color]::Transparent
    $hd.Controls.Add($t)

    $st = New-Object System.Windows.Forms.Label
    $st.Text = "Install Microsoft Office on Windows - select version, language and apps"
    $st.Location = New-Object System.Drawing.Point($M, 50)
    $st.Size = New-Object System.Drawing.Size(600, 20)
    $st.Font = New-Object System.Drawing.Font($fnt, 9)
    $st.ForeColor = $script:dim
    $st.BackColor = [System.Drawing.Color]::Transparent
    $hd.Controls.Add($st)

    $badge = New-Object System.Windows.Forms.Label
    $badge.Text = "64-bit"
    if (-not $script:is64Bit) { $badge.Text = "32-bit" }
    $badge.Location = New-Object System.Drawing.Point(($FW - 110), 28)
    $badge.Size = New-Object System.Drawing.Size(86, 26)
    $badge.TextAlign = "MiddleCenter"
    $badge.Font = New-Object System.Drawing.Font($fnt, 8.5, [System.Drawing.FontStyle]::Bold)
    $badge.ForeColor = $script:ok
    $badge.BackColor = $script:bg
    $hd.Controls.Add($badge)

    $bar = New-Object System.Windows.Forms.Panel
    $bar.Location = New-Object System.Drawing.Point(0, 86)
    $bar.Size = New-Object System.Drawing.Size($FW, 3)
    $bar.BackColor = $script:accent
    $f.Controls.Add($bar)

    $y = 106

    # ===== VERSION =====
    [void](New-SectionLabel $f $M $y "VERSION")
    $y += 22
    $cv = New-Object System.Windows.Forms.ComboBox
    $cv.Name = "cv"
    $cv.Location = New-Object System.Drawing.Point($M, $y)
    $cv.Size = New-Object System.Drawing.Size($GW, 26)
    $cv.DropDownStyle = "DropDownList"
    $cv.Font = New-Object System.Drawing.Font($fnt, 10)
    $cv.BackColor = $script:panel
    $cv.ForeColor = $script:text
    $cv.FlatStyle = "Flat"
    $cv.Items.AddRange(@("Office LTSC Professional Plus 2024", "Office LTSC Professional Plus 2021", "Office Professional Plus 2019", "Office Professional Plus 2016", "Office Professional Plus 2013"))
    $cv.SelectedIndex = 0
    $f.Controls.Add($cv)

    # ===== LANGUAGE =====
    $y += 44
    [void](New-SectionLabel $f $M $y "LANGUAGE")
    $y += 22
    $cl = New-Object System.Windows.Forms.ComboBox
    $cl.Name = "cl"
    $cl.Location = New-Object System.Drawing.Point($M, $y)
    $cl.Size = New-Object System.Drawing.Size($GW, 26)
    $cl.DropDownStyle = "DropDownList"
    $cl.Font = New-Object System.Drawing.Font($fnt, 10)
    $cl.BackColor = $script:panel
    $cl.ForeColor = $script:text
    $cl.FlatStyle = "Flat"
    $cl.Items.AddRange(@("English (en-US)", "Spanish (es-ES)", "French (fr-FR)", "German (de-DE)", "Brazilian Portuguese (pt-BR)", "Italian (it-IT)", "Dutch (nl-NL)", "Polish (pl-PL)", "Russian (ru-RU)", "Japanese (ja-JP)"))
    $cl.SelectedIndex = 0
    $f.Controls.Add($cl)
    # ===== APPLICATIONS =====
    $y += 44
    [void](New-SectionLabel $f $M $y "APPLICATIONS")
    $y += 22
    $ap = New-Object System.Windows.Forms.Panel
    $ap.Location = New-Object System.Drawing.Point($M, $y)
    $ap.Size = New-Object System.Drawing.Size($GW, 122)
    $ap.BackColor = $script:panel
    $f.Controls.Add($ap)

    $pd = 14; $gap = 8
    $ac = New-Object System.Collections.ArrayList
    $cols = 5
    $spc = [Math]::Floor(($GW - 2 * $pd - ($cols - 1) * $gap) / $cols) + $gap
    $bw = $spc - $gap

    $csa = New-Object System.Windows.Forms.CheckBox
    $csa.Name = "csa"; $csa.Text = "Select all"; $csa.Location = New-Object System.Drawing.Point($pd, 10); $csa.Size = New-Object System.Drawing.Size(110, 22); $csa.Checked = $true
    $csa.Font = New-Object System.Drawing.Font($fnt, 9, [System.Drawing.FontStyle]::Bold)
    $csa.ForeColor = $script:text; $csa.BackColor = [System.Drawing.Color]::Transparent
    $ap.Controls.Add($csa)

    $ad = @(@{I="Word";D="Word";C=$true},@{I="Excel";D="Excel";C=$true},@{I="PowerPoint";D="PowerPoint";C=$true},@{I="Outlook";D="Outlook";C=$false},@{I="Access";D="Access";C=$false},@{I="Publisher";D="Publisher";C=$false},@{I="OneNote";D="OneNote";C=$false},@{I="SkypeForBusiness";D="Skype for Business";C=$false},@{I="Project";D="Project";C=$false},@{I="Visio";D="Visio";C=$false})
    for ($i = 0; $i -lt $ad.Count; $i++) {
        $a = $ad[$i]; $col = $i % $cols; $row = [Math]::Floor($i / $cols)
        $c = New-Object System.Windows.Forms.CheckBox
        $c.Name = "c_$($a.I)"; $c.Text = $a.D; $c.Tag = $a.I
        $c.Location = New-Object System.Drawing.Point(($pd + $col * $spc), (42 + $row * 30))
        $c.Size = New-Object System.Drawing.Size($bw, 24)
        $c.Checked = $a.C
        $c.Font = New-Object System.Drawing.Font($fnt, 9)
        $c.ForeColor = $script:text; $c.BackColor = [System.Drawing.Color]::Transparent
        $ap.Controls.Add($c); [void]$ac.Add($c)
        $c.Add_CheckedChanged({ if (-not $updatingAll) { $ca = $true; foreach ($b in $ac) { if (-not $b.Checked) { $ca = $false; break } }; $updatingAll = $true; $csa.Checked = $ca; $updatingAll = $false } })
    }
    $updatingAll = $false
    $csa.Add_CheckedChanged({ if (-not $updatingAll) { $updatingAll = $true; foreach ($b in $ac) { $b.Checked = $csa.Checked }; $updatingAll = $false } })

    # ===== SUMMARY =====
    # El usuario tiene que ver qué va a instalar antes de pulsar. El CLI ya lo
    # hacía; aquí la ventana solo cambiaba un texto al final.
    $y += 140
    $sum = New-Object System.Windows.Forms.Panel
    $sum.Location = New-Object System.Drawing.Point($M, $y)
    $sum.Size = New-Object System.Drawing.Size($GW, 86)
    $sum.BackColor = $script:panel
    $sum.BorderStyle = "FixedSingle"
    $sum.ForeColor = $script:border
    $f.Controls.Add($sum)

    function Add-SummaryField {
        param($Parent, [int]$X, [int]$Y, [string]$Caption, [string]$Name)
        $cap = New-Object System.Windows.Forms.Label
        $cap.Text = $Caption
        $cap.Location = New-Object System.Drawing.Point($X, $Y)
        $cap.Size = New-Object System.Drawing.Size(370, 14)
        $cap.Font = New-Object System.Drawing.Font($fnt, 7.5, [System.Drawing.FontStyle]::Bold)
        $cap.ForeColor = $script:dim
        $cap.BackColor = [System.Drawing.Color]::Transparent
        $Parent.Controls.Add($cap)

        $val = New-Object System.Windows.Forms.Label
        $val.Name = $Name
        $val.Location = New-Object System.Drawing.Point($X, ($Y + 15))
        $val.Size = New-Object System.Drawing.Size(370, 18)
        # Una lista larga de aplicaciones se corta con puntos suspensivos en
        # vez de salirse del panel por la derecha.
        $val.AutoEllipsis = $true
        $val.Font = New-Object System.Drawing.Font($fnt, 9.5)
        $val.ForeColor = $script:text
        $val.BackColor = [System.Drawing.Color]::Transparent
        $Parent.Controls.Add($val)
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
    $sb.ForeColor = $script:dim
    $sb.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($sb)

    $y += 24
    $track = New-Object System.Windows.Forms.Panel
    $track.Location = New-Object System.Drawing.Point($M, $y)
    $track.Size = New-Object System.Drawing.Size($GW, 8)
    $track.BackColor = $script:track
    $f.Controls.Add($track)

    $fill = New-Object System.Windows.Forms.Panel
    $fill.Location = New-Object System.Drawing.Point($M, $y)
    $fill.Size = New-Object System.Drawing.Size(0, 8)
    $fill.BackColor = $script:accent
    $f.Controls.Add($fill)

    function Set-Progress {
        param([int]$Pct, [string]$Text, [string]$State = "run")
        if ($Pct -lt 0) { $Pct = 0 }
        if ($Pct -gt 100) { $Pct = 100 }
        $fill.Size = New-Object System.Drawing.Size([int][Math]::Floor($GW * $Pct / 100), 8)
        $sb.Text = $Text
        switch ($State) {
            "ok"    { $fill.BackColor = $script:ok;    $sb.ForeColor = $script:ok }
            "bad"   { $fill.BackColor = $script:bad;   $sb.ForeColor = $script:bad }
            "warn"  { $fill.BackColor = $script:warn;  $sb.ForeColor = $script:warn }
            "idle"  { $fill.BackColor = $script:accent; $sb.ForeColor = $script:dim }
            default { $fill.BackColor = $script:accent; $sb.ForeColor = $script:text }
        }
        $f.Refresh()
    }

    # ===== LIVE SUMMARY =====
    # Se dispara con cada cambio de combo o de casilla, así que lo que se ve
    # en el panel es exactamente lo que se escribe en el configuration.xml.
    function Update-Summary {
        $sumEd.Text   = $cv.SelectedItem.ToString()
        $sumLang.Text = $cl.SelectedItem.ToString()
        $sumArch.Text = if ($script:is64Bit) { "64-bit" } else { "32-bit" }

        $picked = @()
        foreach ($b in $ac) { if ($b.Checked) { $picked += $b.Text } }

        if ($picked.Count -eq 0) {
            $sumApps.Text = "Nothing selected - Office cannot be installed"
            $sumApps.ForeColor = $script:warn
        } else {
            $sumApps.Text = $picked -join ", "
            $sumApps.ForeColor = $script:text
        }
    }

    foreach ($c in $ac) { $c.Add_CheckedChanged({ Update-Summary }) }
    $csa.Add_CheckedChanged({ Update-Summary })
    $cv.Add_SelectedIndexChanged({ Update-Summary })
    $cl.Add_SelectedIndexChanged({ Update-Summary })
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
    $bi.BackColor = $script:accent
    $bi.ForeColor = [System.Drawing.Color]::White
    # Se aclara al pasar el ratón, no se oscurece.
    $bi.Add_MouseEnter({ $bi.BackColor = $script:accentH })
    $bi.Add_MouseLeave({ $bi.BackColor = $script:accent })
    $f.Controls.Add($bi)

    $bc = New-Object System.Windows.Forms.Button
    $bc.Name = "bc"
    $bc.Text = "$([char]0x2715)  Cancel"
    $bc.Location = New-Object System.Drawing.Point(($btnStart + $biW + $btnGap), $y)
    $bc.Size = New-Object System.Drawing.Size($bcW, 42)
    $bc.Font = New-Object System.Drawing.Font($fnt, 10)
    $bc.FlatStyle = "Flat"
    $bc.FlatAppearance.BorderSize = 1
    $bc.FlatAppearance.BorderColor = $script:border
    $bc.FlatAppearance.MouseOverBackColor = $script:panelH
    $bc.BackColor = $script:panel
    $bc.ForeColor = $script:dim
    $f.Controls.Add($bc)

    # ===== BYLINE =====
    $by = New-Object System.Windows.Forms.Label
    $by.Text = "Microsoft Office Installer  $([char]0x00B7)  by Matthew Garay"
    $by.Location = New-Object System.Drawing.Point($M, ($y + 48))
    $by.Size = New-Object System.Drawing.Size($GW, 18)
    $by.TextAlign = "MiddleCenter"
    $by.Font = New-Object System.Drawing.Font($fnt, 8.5)
    $by.ForeColor = $script:dim
    $by.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($by)

    # ===== LICENSE NOTICE =====
    # La instalación va a pedir una licencia igualmente; decirlo aquí es mejor
    # que que el usuario se entere después.
    $lic = New-Object System.Windows.Forms.Label
    $lic.Text = "This installer only installs Office. A valid Office licence is required to use it."
    $lic.Location = New-Object System.Drawing.Point($M, ($y + 74))
    $lic.Size = New-Object System.Drawing.Size($GW, 18)
    $lic.TextAlign = "MiddleCenter"
    $lic.Font = New-Object System.Drawing.Font($fnt, 8)
    $lic.ForeColor = $script:warn
    $lic.BackColor = [System.Drawing.Color]::Transparent
    $f.Controls.Add($lic)

    # ===== VERSION MAP =====
    $vm = @(@{C="PerpetualVL2024";P="ProPlus2024Volume";V="VisioPro2024Volume";J="ProjectPro2024Volume"},@{C="PerpetualVL2021";P="ProPlus2021Volume";V="VisioPro2021Volume";J="ProjectPro2021Volume"},@{C="PerpetualVL2019";P="ProPlus2019Volume";V="VisioPro2019Volume";J="ProjectPro2019Volume"},@{C="PerpetualVL2016";P="ProPlus2016Volume";V="VisioPro2016Volume";J="ProjectPro2016Volume"},@{C="PerpetualVL2013";P="ProPlus2013Volume";V="VisioPro2013Volume";J="ProjectPro2013Volume"})

    $bc.Add_Click({ $f.Close() })
    $bi.Add_Click({
        # Stage-based progress, not a real-time counter: setup.exe blocks with
        # -Wait, so there is no percentage to read while Office downloads. Each
        # bar position is a phase the installer has actually reached.
        $vi = $vm[$cv.SelectedIndex]
        $arch = if ($script:is64Bit) { "64" } else { "32" }
        $lang = Get-LangCode $cl.SelectedItem.ToString()
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




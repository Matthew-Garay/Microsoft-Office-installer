<#
.SYNOPSIS
    Microsoft Office Installer - CLI Edition
.DESCRIPTION
    Command-line tool to download and install Microsoft Office LTSC.
    Supports multiple editions, architectures, languages and app selection.
.NOTES
    Requirements: Administrator, PowerShell 5.0+, .NET Framework 4.5+
    Usage: irm https://matthew-garay.github.io/Microsoft-Office-installer/instalar-cli.ps1 | iex
#>

# ---- ADMIN CHECK ----
$script:isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $script:isAdmin) {
    Write-Host ""
    Write-Host "  $([char]0x2718) This script requires Administrator privileges." -ForegroundColor Red
    Write-Host ""
    Read-Host "  Press Enter to exit"
    exit 1
}

# ---- CONSTANTS ----
$script:logFile = Join-Path $env:Temp "OfficeInstallerCLI_Install.log"
$script:odtTemp = Join-Path $env:Temp "ODT"
$script:odtExe  = Join-Path $env:Temp "OfficeDeploymentTool.exe"
$script:odtUrls = @("https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe")
$script:is64Bit = [Environment]::Is64BitOperatingSystem
$script:startAt = Get-Date

# ---- COLOR HELPERS ----
function Color {
    param([string]$T, [string]$C = "White", [string]$B = "")
    if ($B) { Write-Host $T -ForegroundColor $C -BackgroundColor $B -NoNewline }
    else    { Write-Host $T -ForegroundColor $C -NoNewline }
}

function Rule {
    param([string]$C = "DarkGray")
    Color "  $([char]0x2500)$([char]0x2500)" $C
    Color ("$([char]0x2500)" * 72) $C
    Write-Host ""
}

function Ok   { param([string]$M) Color "  $([char]0x2714) " "Green";  Color $M "White";  Write-Host "" }
function Fail { param([string]$M) Color "  $([char]0x2718) " "Red";    Color $M "Red";    Write-Host "" }
function Info { param([string]$M) Color "  $([char]0x25B8) " "Cyan";   Color $M "White";  Write-Host "" }
function Dim  { param([string]$M) Color "    $M" "DarkGray"; Write-Host "" }

# ---- LOGGING ----
function Write-Log {
    param([string]$M)
    Add-Content -Path $script:logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $M" -ErrorAction SilentlyContinue
}

# ---- BANNER ----
function Show-Banner {
    Clear-Host
    $W = 66
    $top    = "  $([char]0x2554)" + ("$([char]0x2550)" * $W) + "$([char]0x2557)"
    $bot    = "  $([char]0x255A)" + ("$([char]0x2550)" * $W) + "$([char]0x255D)"
    $side   = "  $([char]0x2551)"
    function Box($txt, $col) {
        $pad = $W - 2 - $txt.Length
        $l = [Math]::Floor($pad / 2); $r = $pad - $l
        Color ($side + (" " * ($l + 1))) "Cyan"
        Color $txt $col
        Color (" " * ($r + 1) + "$([char]0x2551)") "Cyan"
        Write-Host ""
    }
    Write-Host ""
    Color $top "Cyan"; Write-Host ""
    Box "MICROSOFT OFFICE INSTALLER" "White"
    Box "Command Line Edition" "Cyan"
    Box "" "Cyan"
    Color $bot "Cyan"; Write-Host ""
    $by = "by Matthew Garay"
    $pb = [Math]::Floor((($W + 4) - $by.Length) / 2)
    Color (" " * $pb + $by) "DarkGray"; Write-Host ""
    Write-Host ""
    Rule
    Write-Host ""
}

# ---- STEP HEADER ----
function Show-Step {
    param([int]$N, [int]$Total, [string]$Title)
    Write-Host ""
    Color "  $([char]0x25CF) STEP $N/$T" "Cyan"
    Color "  $([char]0x2500) " "DarkGray"
    Color $Title "White"
    $fill = 68 - $Title.Length
    if ($fill -lt 3) { $fill = 3 }
    Color (" $([char]0x2500)" * 1) "DarkGray"
    Color ("$([char]0x2500)" * $fill) "DarkGray"
    Write-Host ""
    Write-Host ""
}

# ---- PROGRESS BAR ----
function Show-ProgressBar {
    param([int]$Pct, [string]$Label = "Progress", [int]$W = 32)
    $f = [Math]::Floor($W * $Pct / 100)
    $e = $W - $f
    $bar = ("$([char]0x2588)" * $f) + ("$([char]0x2591)" * $e)
    Color "  $Label " "DarkGray"
    Color $bar "Cyan"
    Color (" {0,3}%" -f $Pct) "Yellow"
    if ($Pct -eq 100) { Write-Host "" } else { Write-Host "`r" -NoNewline }
}
# __C2__

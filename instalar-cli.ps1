<#
.SYNOPSIS
    Microsoft Office Installer - CLI Edition
.DESCRIPTION
    Command-line tool to download and install Microsoft Office LTSC editions.
    Supports multiple editions, architectures, languages and app selection.

    Guided interactive flow: every choice is picked on screen, and nothing
    is installed without confirmation.
.EXAMPLE
    .\instalar-cli.ps1
.EXAMPLE
    irm https://raw.githubusercontent.com/Matthew-Garay/Microsoft-Office-installer/main/instalar-cli.ps1 | iex
.NOTES
    Requirements: Administrator, PowerShell 5.0+, .NET Framework 4.5+, Internet.

    This script installs Office; it does not activate it. A valid license is
    required and nothing here bypasses activation.
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

# ---- VERSION MAP ----
# C = channel, P = Office product, V = Visio product, J = Project product.
$script:versions = @(
    @{ Label = "Office LTSC Professional Plus 2024"; C = "PerpetualVL2024"; P = "ProPlus2024Volume"; V = "VisioPro2024Volume"; J = "ProjectPro2024Volume" }
    @{ Label = "Office LTSC Professional Plus 2021"; C = "PerpetualVL2021"; P = "ProPlus2021Volume"; V = "VisioPro2021Volume"; J = "ProjectPro2021Volume" }
    @{ Label = "Office Professional Plus 2019";      C = "PerpetualVL2019"; P = "ProPlus2019Volume"; V = "VisioPro2019Volume"; J = "ProjectPro2019Volume" }
    @{ Label = "Office Professional Plus 2016";      C = "PerpetualVL2016"; P = "ProPlus2016Volume"; V = "VisioPro2016Volume"; J = "ProjectPro2016Volume" }
    @{ Label = "Office Professional Plus 2013";      C = "PerpetualVL2013"; P = "ProPlus2013Volume"; V = "VisioPro2013Volume"; J = "ProjectPro2013Volume" }
)

# ---- LANGUAGES ----
$script:languages = @(
    "English (en-US)"
    "Spanish (es-ES)"
    "French (fr-FR)"
    "German (de-DE)"
    "Brazilian Portuguese (pt-BR)"
    "Italian (it-IT)"
    "Dutch (nl-NL)"
    "Polish (pl-PL)"
    "Russian (ru-RU)"
    "Japanese (ja-JP)"
)

# ---- APPLICATIONS ----
# Base = belongs to the Office product. Addon = installed as its own product.
# The defaults match the GUI: the three apps everyone expects, nothing else.
$script:appCatalog = @(
    @{ Tag = "Word";             Label = "Word";               Base = $true;  Addon = $false }
    @{ Tag = "Excel";            Label = "Excel";              Base = $true;  Addon = $false }
    @{ Tag = "PowerPoint";       Label = "PowerPoint";         Base = $true;  Addon = $false }
    @{ Tag = "Outlook";          Label = "Outlook";            Base = $false; Addon = $false }
    @{ Tag = "Access";           Label = "Access";             Base = $false; Addon = $false }
    @{ Tag = "Publisher";        Label = "Publisher";          Base = $false; Addon = $false }
    @{ Tag = "OneNote";          Label = "OneNote";            Base = $false; Addon = $false }
    @{ Tag = "SkypeForBusiness"; Label = "Skype for Business"; Base = $false; Addon = $false }
    @{ Tag = "Project";          Label = "Project";            Base = $false; Addon = $true }
    @{ Tag = "Visio";            Label = "Visio";              Base = $false; Addon = $true }
)

# Apps the Office product ships with. Whatever is not selected is excluded.
$script:baseApps = @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote","SkypeForBusiness")

# Bundled services nobody asked for and that add weight without use.
$script:bundleExcluded = @("Bing","Groove","Lync","OneDrive","Teams")

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
    Color "  $([char]0x25CF) STEP $N/$Total" "Cyan"
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
# ---- SINGLE CHOICE PROMPT ----
# Prints a numbered list and returns the chosen index. Re-prompting on a bad
# number matters here: an empty answer silently picking option 1 would install
# an edition the operator never chose.
function Ask-Choice {
    param([string]$Title, [string[]]$Options, [int]$Default = 1)

    Write-Host ""
    Color "  $Title " "White"
    Color ("$([char]0x2500)" * ($Title.Length + 2)) "DarkGray"
    Write-Host ""
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Color ("    {0,2}. " -f ($i + 1)) "DarkGray"
        Color $Options[$i] "White"
        Write-Host ""
    }
    Write-Host ""
    Dim "Enter a number 1-$($Options.Count), or press Enter for $Default"

    while ($true) {
        Color "  > " "Cyan"
        $raw = Read-Host
        if ($null -eq $raw) { Write-Host ""; return $Default - 1 }
        $raw = $raw.Trim()
        if (-not $raw) { return $Default - 1 }
        $n = 0
        if ([int]::TryParse($raw, [ref]$n) -and $n -ge 1 -and $n -le $Options.Count) { return $n - 1 }
        Fail "Not one of 1-$($Options.Count): $raw"
    }
}

# ---- APPLICATION PROMPT ----
# Multi-select by number. Returns the chosen tags.
function Ask-Apps {
    Write-Host ""
    Color "  Applications " "White"
    Color ("$([char]0x2500)" * 16) "DarkGray"
    Write-Host ""
    for ($i = 0; $i -lt $script:appCatalog.Count; $i++) {
        $a = $script:appCatalog[$i]
        Color ("    {0,2}. " -f ($i + 1)) "DarkGray"
        if ($a.Addon) { Color "+ " "DarkCyan" } else { Color "  " "DarkGray" }
        Color $a.Label "White"
        if ($a.Addon) {
            Color "  (separate product)" "DarkGray"
        } elseif (-not $a.Base) {
            Color "  (not selected by default)" "DarkGray"
        }
        Write-Host ""
    }
    Write-Host ""
    Dim "Numbers separated by commas. 1,2,3 = Word, Excel, PowerPoint."
    Dim "The ones marked + are installed as their own product."

    while ($true) {
        Color "  > " "Cyan"
        $raw = Read-Host
        if ($null -eq $raw) { $raw = "1,2,3" }
        $raw = $raw.Trim()
        if (-not $raw) { $raw = "1,2,3" }

        $picked = @()
        $bad = $false
        foreach ($part in ($raw -split "[,\s]+")) {
            if (-not $part) { continue }
            $n = 0
            if (-not [int]::TryParse($part, [ref]$n) -or $n -lt 1 -or $n -gt $script:appCatalog.Count) {
                Fail "Not a valid option: $part"
                $bad = $true
                continue
            }
            $tag = $script:appCatalog[$n - 1].Tag
            if ($picked -notcontains $tag) { $picked += $tag }
        }
        if ($bad) { continue }
        if ($picked.Count -eq 0) { Fail "Nothing selected."; continue }

        return $picked
    }
}

function Get-LangCode {
    param([string]$T)
    if ($T -match '\(([^)]+)\)') { return $matches[1] }
    return "en-US"
}

# ---- OFFICE DEPLOYMENT TOOL ----
# Downloads the ODT and unpacks setup.exe. Kept as a function so the failure
# path can report which stage broke instead of dying mid-download.
function Initialize-Odt {
    Write-Host ""
    Color "  Cleaning up previous files " "DarkGray"; Write-Host ""
    if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $script:odtTemp -Force | Out-Null

    Show-ProgressBar 5 "Downloading ODT"

    $downloaded = $false
    $attempt = 0
    foreach ($u in $script:odtUrls) {
        $attempt++
        try {
            Write-Log "Downloading ODT from $u"
            (New-Object System.Net.WebClient).DownloadFile($u, $script:odtExe)
            $downloaded = $true
            break
        } catch {
            Write-Log "Source failed: $u ($_)"
            if ($attempt -lt $script:odtUrls.Count) {
                Show-ProgressBar 5 "Retrying another source"
                Fail "Source unavailable, trying the next one."
            }
        }
    }
    if (-not $downloaded) { throw "Could not download the Office Deployment Tool from any source." }

    Show-ProgressBar 45 "Downloaded"
    Show-ProgressBar 50 "Extracting ODT"

    $p = Start-Process -FilePath $script:odtExe -ArgumentList "/quiet /extract:`"$script:odtTemp`"" -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw "Extraction failed (ExitCode: $($p.ExitCode))." }
    if (-not (Test-Path (Join-Path $script:odtTemp "setup.exe"))) { throw "setup.exe not found after extraction." }

    Show-ProgressBar 100 "Ready"
    Write-Log "ODT ready at $script:odtTemp"
}

# ---- CONFIGURATION XML ----
# Mirrors the GUI generator, so both editions produce the same configuration
# for the same choices.
function New-ConfigurationXml {
    param($Ver, [string]$Lang, [string[]]$Selected)

    $arch = if ($script:is64Bit) { "64" } else { "32" }
    $incP = ($Selected -contains "Project")
    $incV = ($Selected -contains "Visio")

    $x = New-Object System.Text.StringBuilder
    [void]$x.AppendLine('<Configuration>')
    [void]$x.AppendLine("    <Add OfficeClientEdition=`"$arch`" Channel=`"$($Ver.C)`">")
    [void]$x.AppendLine("        <Product ID=`"$($Ver.P)`">")
    [void]$x.AppendLine("            <Language ID=`"$Lang`" />")
    foreach ($app in $script:baseApps) {
        if ($Selected -notcontains $app) { [void]$x.AppendLine("            <ExcludeApp ID=`"$app`" />") }
    }
    foreach ($ex in $script:bundleExcluded) { [void]$x.AppendLine("            <ExcludeApp ID=`"$ex`" />") }
    [void]$x.AppendLine("        </Product>")
    if ($incP) { [void]$x.AppendLine("        <Product ID=`"$($Ver.J)`"><Language ID=`"$Lang`" /></Product>") }
    if ($incV) { [void]$x.AppendLine("        <Product ID=`"$($Ver.V)`"><Language ID=`"$Lang`" /></Product>") }
    [void]$x.AppendLine("    </Add>")
    [void]$x.AppendLine('    <Display Level="Full" AcceptEULA="TRUE" />')
    [void]$x.AppendLine('</Configuration>')

    $path = Join-Path $script:odtTemp "configuration.xml"
    Set-Content -Path $path -Value $x.ToString() -Encoding UTF8
    Write-Log "Configuration written to $path"
    return $path
}

# ====================================================
# MAIN
# ====================================================
Write-Log "=== Microsoft Office Installer (CLI) started ==="
Write-Log "System: $((Get-CimInstance Win32_OperatingSystem).Caption)"
Write-Log "OS: $(if ($script:is64Bit) { '64' } else { '32' })-bit"
Write-Log "PowerShell: $($PSVersionTable.PSVersion)"

Show-Banner

# ---- STEP 1: VERSION ----
Show-Step 1 5 "Select the Office version"
$versionLabels = @(); foreach ($v in $script:versions) { $versionLabels += $v.Label }
$vi = $script:versions[(Ask-Choice -Title "VERSION" -Options $versionLabels)]
Ok $vi.Label

# ---- STEP 2: LANGUAGE ----
Show-Step 2 5 "Select the language"
$langCodes = @(); foreach ($l in $script:languages) { $langCodes += (Get-LangCode $l) }
$lang = $langCodes[(Ask-Choice -Title "LANGUAGE" -Options $script:languages)]
Ok $lang

# ---- STEP 3: APPLICATIONS ----
Show-Step 3 5 "Select the applications"
$selected = Ask-Apps
foreach ($t in $selected) {
    $lbl = ($script:appCatalog | Where-Object { $_.Tag -eq $t } | Select-Object -First 1).Label
    Dim "+ $lbl"
}

# ---- STEP 4: OFFICE DEPLOYMENT TOOL ----
Show-Step 4 5 "Prepare the Office Deployment Tool"
try {
    Initialize-Odt
    Ok "Office Deployment Tool ready"
} catch {
    Fail "Could not prepare the Office Deployment Tool: $_"
    Write-Log "Prep error: $_"
    Dim "Download it by hand from https://www.microsoft.com/en-us/download/details.aspx?id=49117"
    exit 1
}

# ---- STEP 5: INSTALL ----
Show-Step 5 5 "Install Office"
$configPath = New-ConfigurationXml -Ver $vi -Lang $lang -Selected $selected
Ok "Configuration written"
Dim $configPath

Write-Host ""
Color "  Summary " "White"
Color ("$([char]0x2500)" * 11) "DarkGray"
Write-Host ""
Dim "Edition  : $($vi.Label)"
Dim "Language : $lang"
Dim "Apps     : $($selected -join ', ')"
Dim "System   : $(if ($script:is64Bit) { '64' } else { '32' })-bit"
Write-Host ""

Dim "An Office installation cannot be undone by this script."
Dim "If Office is already installed, this replaces it."
Write-Host ""
Color "  Press ENTER to start, or type n to abort: " "White"
$ans = Read-Host
$ans = if ($null -eq $ans) { "" } else { $ans.Trim().ToLower() }
if ($ans -eq "n" -or $ans -eq "no") {
    Info "Aborted. Nothing was changed."
    Write-Log "Aborted by the operator"
    if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
    exit 0
}

Show-ProgressBar 0 "Installing"
Write-Log "Running setup.exe /configure $configPath"

$setupExe = Join-Path $script:odtTemp "setup.exe"
$proc = Start-Process -FilePath $setupExe -ArgumentList "/configure `"$configPath`"" -Wait -PassThru

if ($proc.ExitCode -ne 0) {
    Show-ProgressBar 100 "Failed"
    Fail "Installation failed (exit code $($proc.ExitCode))."
    Write-Log "setup.exe failed with exit code $($proc.ExitCode)"
    Dim "The configuration is kept at $configPath"
    Dim "Log: $script:logFile"
    exit 1
}

Show-ProgressBar 100 "Installed"
Ok "Office installed successfully."

Write-Host ""
Color "  Cleaning up " "DarkGray"; Write-Host ""
if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
Ok "Done"

Write-Host ""
Info "A valid Office license is required to use it. This installer does not activate Office."
Dim "Log: $script:logFile"
Write-Host ""
Write-Log "=== Microsoft Office Installer (CLI) finished ==="
exit 0

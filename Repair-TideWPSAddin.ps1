<#
.SYNOPSIS
    Registers the "Tide" (Tide term audit / Tide 术语审校) WPS Spreadsheets JS add-in on Windows.

.DESCRIPTION
    WPS loads JavaScript add-ins from ONE file (confirmed in WPS 12.1.0.26899,
    office6\addons\jsapi\jsapisubserver.dll):

        %APPDATA%\kingsoft\wps\jsaddons\publish.xml

    The shipped installer "Tide术语审校-0.7.3-Windows.exe" creates that folder but,
    on some machines, never writes publish.xml (and its install-win.ps1 default URL
    http://127.0.0.1:8902/ does not match serve-local.bat's port 8901). Result:
    the add-in files exist on disk but WPS shows no "Tide 术语审校" ribbon tab.

    This script copies the add-in payload into jsaddons and writes/merges publish.xml.
    It is idempotent: only the Tide entry is replaced, other entries are preserved.

.PARAMETER CheckOnly
    Read-only diagnosis. Changes nothing. Best first step on a new machine.

.PARAMETER Online
    Register an online (http) entry instead of copying the payload locally.
    WPS then fetches the add-in from -OnlineUrl on every start (needs network).

.PARAMETER Remove
    Roll back: remove the Tide entry from publish.xml and delete the add-on folder.

.PARAMETER SourceDir
    Override the add-in payload folder (the folder that contains manifest.xml,
    index.html, main.js, ribbon.xml). Default: <InstallLocation>\addin.

.PARAMETER OnlineUrl
    URL used by -Online. Default: https://bohuaraynorwei-hash.github.io/tide-wps-addin/

.PARAMETER Json
    Print a machine-readable JSON summary instead of prose.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Repair-TideWPSAddin.ps1 -CheckOnly

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Repair-TideWPSAddin.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Repair-TideWPSAddin.ps1 -Online

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Repair-TideWPSAddin.ps1 -Remove
#>
[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [switch]$Online,
    [switch]$Remove,
    [string]$SourceDir,
    [string]$OnlineUrl = 'https://bohuaraynorwei-hash.github.io/tide-wps-addin/',
    [switch]$Json
)

$ErrorActionPreference = 'Stop'

$PluginName  = 'TideTermAudit'
$AddonType   = 'et'                      # et = WPS Spreadsheets; the tab only shows there
$jsaddonsDir = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$publishPath = Join-Path $jsaddonsDir 'publish.xml'
$backupPath  = "$publishPath.tide-bak"

$summary = [ordered]@{
    plugin        = $PluginName
    action        = ''
    ok            = $false
    machine       = $env:COMPUTERNAME
    user          = $env:USERNAME
    jsaddonsDir   = $jsaddonsDir
    publishXml    = $publishPath
    installDir    = ''
    payloadSource = ''
    payloadTarget = ''
    version       = ''
    wpsRunning    = @()
    problems      = @()
    notes         = @()
}

function Write-Step([string]$text) { if (-not $Json) { Write-Host $text } }
function Add-Problem([string]$text) { $summary.problems += $text; if (-not $Json) { Write-Host "  [PROBLEM] $text" } }
function Add-Note([string]$text)    { $summary.notes    += $text; if (-not $Json) { Write-Host "  [note]    $text" } }

# --------------------------------------------------------------------------- #
# 1. Locate the payload shipped by the installer
# --------------------------------------------------------------------------- #
function Find-TideInstall {
    $roots = @(
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\TideTermAudit',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\TideTermAudit',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\TideTermAudit'
    )
    foreach ($r in $roots) {
        if (Test-Path -LiteralPath $r) {
            $p = Get-ItemProperty -LiteralPath $r
            if ($p.InstallLocation) {
                $loc = $p.InstallLocation.TrimEnd('\')
                if (Test-Path -LiteralPath $loc) {
                    return [pscustomobject]@{ Location = $loc; Version = [string]$p.DisplayVersion }
                }
            }
        }
    }
    $fallback = Join-Path $env:LOCALAPPDATA 'TideTermAudit'
    if (Test-Path -LiteralPath $fallback) {
        return [pscustomobject]@{ Location = $fallback; Version = '' }
    }
    return $null
}

function Test-Payload([string]$dir) {
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return $false }
    foreach ($f in @('manifest.xml', 'index.html', 'main.js', 'ribbon.xml')) {
        if (-not (Test-Path -LiteralPath (Join-Path $dir $f))) { return $false }
    }
    return $true
}

$install = Find-TideInstall
if ($install) { $summary.installDir = $install.Location }

$payload = $null
if ($SourceDir) {
    if (-not (Test-Payload $SourceDir)) { throw "SourceDir does not look like an add-in payload (manifest.xml/index.html/main.js/ribbon.xml missing): $SourceDir" }
    $payload = (Resolve-Path -LiteralPath $SourceDir).Path
} elseif ($install) {
    foreach ($cand in @((Join-Path $install.Location 'addin'), $install.Location)) {
        if (Test-Payload $cand) { $payload = $cand; break }
    }
}
if ($payload) { $summary.payloadSource = $payload }

$version = $install.Version
if ($version -notmatch '^\d+(\.\d+){1,3}$') {
    $version = '0.7.3'
    Add-Note "installer version not readable, assuming $version for the folder name"
}
$summary.version = $version
$folderName = "{0}_{1}" -f $PluginName, $version
$targetDir  = Join-Path $jsaddonsDir $folderName

# --------------------------------------------------------------------------- #
# 2. Detect running WPS
# --------------------------------------------------------------------------- #
$wpsProcs = @(Get-Process -ErrorAction SilentlyContinue |
    Where-Object { $_.ProcessName -match '^(wps|et|wpp|wpscloudsvr|wpscenter|ksolaunch)$' })
$summary.wpsRunning = @($wpsProcs | ForEach-Object { "$($_.ProcessName):$($_.Id)" })
$wpsExe = ($wpsProcs | Where-Object { $_.ProcessName -eq 'wps' } | Select-Object -First 1).Path

# --------------------------------------------------------------------------- #
# 3. Helpers for publish.xml
# --------------------------------------------------------------------------- #
function New-PublishDocument {
    $doc = New-Object System.Xml.XmlDocument
    $pi  = $doc.CreateProcessingInstruction('xml', 'version="1.0" encoding="UTF-8" standalone="yes"')
    [void]$doc.AppendChild($pi)
    [void]$doc.AppendChild($doc.CreateElement('jsplugins'))
    return $doc
}

function Read-PublishDocument([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return (New-PublishDocument) }
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.Load($path)
    } catch {
        Add-Note "existing publish.xml is not valid XML, it will be rebuilt (backup kept): $($_.Exception.Message)"
        return (New-PublishDocument)
    }
    if (-not $doc.DocumentElement -or $doc.DocumentElement.Name -ne 'jsplugins') {
        Add-Note "existing publish.xml has unexpected root, it will be rebuilt (backup kept)"
        return (New-PublishDocument)
    }
    return $doc
}

function Remove-TideEntries([System.Xml.XmlDocument]$doc) {
    if (-not $doc.DocumentElement) { return }
    foreach ($n in @($doc.DocumentElement.ChildNodes)) {
        if ($n.NodeType -ne 'Element') { continue }
        # NOTE: $n.Name returns the "name" ATTRIBUTE for XML nodes in PowerShell,
        # so the tag name must be read through LocalName.
        if ($n.LocalName -ne 'jsplugin' -and $n.LocalName -ne 'jspluginonline') { continue }
        $name = if ($n.HasAttribute('name')) { $n.GetAttribute('name') } else { '' }
        if ($name -ieq $PluginName) { [void]$doc.DocumentElement.RemoveChild($n) }
    }
}

function Save-PublishDocument([System.Xml.XmlDocument]$doc, [string]$path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.IndentChars = '    '
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)   # UTF-8 without BOM
    $writer = [System.Xml.XmlWriter]::Create($path, $settings)
    try { $doc.Save($writer) } finally { $writer.Close() }
}

function Get-PublishEntries([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    try { [xml]$d = Get-Content -LiteralPath $path -Raw -Encoding UTF8 } catch { return @() }
    $out = @()
    foreach ($n in @($d.DocumentElement.ChildNodes)) {
        if ($n.NodeType -ne 'Element') { continue }
        $out += ("{0} name={1} type={2} url={3}" -f $n.LocalName,
                 $(if ($n.HasAttribute('name')) { $n.GetAttribute('name') } else { '' }),
                 $(if ($n.HasAttribute('type')) { $n.GetAttribute('type') } else { '' }),
                 $(if ($n.HasAttribute('url'))  { $n.GetAttribute('url') }  else { '' }))
    }
    return $out
}

# --------------------------------------------------------------------------- #
# 4. -CheckOnly : diagnosis only
# --------------------------------------------------------------------------- #
if ($CheckOnly) {
    $summary.action = 'check'
    Write-Step "=== Tide WPS add-in diagnosis ==="
    Write-Step "user        : $env:USERNAME@$env:COMPUTERNAME"
    Write-Step "install dir : $(if ($install) { $install.Location } else { 'NOT FOUND in registry' })"
    Write-Step "version     : $(if ($install) { $install.Version } else { 'unknown' })"
    Write-Step "payload     : $(if ($payload) { $payload } else { 'MISSING (manifest.xml/index.html/main.js/ribbon.xml)' })"
    Write-Step "jsaddons    : $jsaddonsDir (exists: $(Test-Path -LiteralPath $jsaddonsDir))"
    Write-Step "publish.xml : $publishPath (exists: $(Test-Path -LiteralPath $publishPath))"
    Write-Step "wps exe     : $(if ($wpsExe) { $wpsExe } else { 'not running / not found' })"
    Write-Step "wps procs   : $(if ($summary.wpsRunning.Count) { $summary.wpsRunning -join ', ' } else { 'none' })"
    Write-Step ""
    Write-Step "-- contents of jsaddons --"
    if (Test-Path -LiteralPath $jsaddonsDir) {
        $kids = @(Get-ChildItem -LiteralPath $jsaddonsDir -Force -ErrorAction SilentlyContinue)
        if ($kids.Count -eq 0) { Write-Step "  <empty>   <-- this alone explains the missing ribbon tab" }
        else { $kids | ForEach-Object { Write-Step ("  {0}  {1}" -f $_.Name, $(if ($_.PSIsContainer) { '<dir>' } else { "$($_.Length) B" })) } }
    } else {
        Write-Step "  <jsaddons folder does not exist>"
    }
    Write-Step ""
    Write-Step "-- publish.xml entries --"
    $entries = @(Get-PublishEntries $publishPath)
    if ($entries.Count -eq 0) { Write-Step "  <none>" } else { $entries | ForEach-Object { Write-Step "  $_" } }

    $tideRegistered = $false
    foreach ($e in $entries) { if ($e -match "name=$PluginName(\s|$)") { $tideRegistered = $true } }
    $targetExists = (Test-Path -LiteralPath $targetDir) -and (Test-Path -LiteralPath (Join-Path $targetDir 'manifest.xml'))
    $anyTideFolder = @(Get-ChildItem -LiteralPath $jsaddonsDir -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "$PluginName*" })

    if (-not $payload) { Add-Problem 'add-in payload not found; run with -SourceDir pointing at the folder that holds manifest.xml' }
    if (-not $tideRegistered) { Add-Problem 'publish.xml has no TideTermAudit entry -> WPS cannot see the add-in at all' }
    if ($tideRegistered -and -not $targetExists -and $anyTideFolder.Count -eq 0) {
        Add-Problem "publish.xml points to a local folder that does not exist ($targetDir)"
    }
    foreach ($p in $summary.wpsRunning) { Add-Note "WPS process still running ($p) - it must be fully quit before the tab can appear" }

    $summary.ok = ($summary.problems.Count -eq 0)
    Write-Step ""
    if ($summary.ok) { Write-Step 'VERDICT: registration looks OK.' }
    else             { Write-Step 'VERDICT: NEEDS FIX - run this script without -CheckOnly.' }

    if ($Json) { $summary | ConvertTo-Json -Depth 5 }
    if ($summary.ok) { exit 0 } else { exit 1 }
}

# --------------------------------------------------------------------------- #
# 5. -Remove : rollback
# --------------------------------------------------------------------------- #
if ($Remove) {
    $summary.action = 'remove'
    Write-Step "=== Removing Tide add-in registration ==="

    if (Test-Path -LiteralPath $publishPath) {
        $doc = Read-PublishDocument $publishPath
        Remove-TideEntries $doc
        if (-not (Test-Path -LiteralPath $backupPath)) { Copy-Item -LiteralPath $publishPath -Destination $backupPath -Force }
        Save-PublishDocument $doc $publishPath
        Write-Step "  publish.xml updated: $publishPath"
    } else {
        Write-Step "  publish.xml not present, nothing to unregister"
    }

    foreach ($d in @(Get-ChildItem -LiteralPath $jsaddonsDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "$PluginName*" })) {
        Remove-Item -LiteralPath $d.FullName -Recurse -Force
        Write-Step "  removed folder: $($d.FullName)"
    }

    $summary.ok = $true
    Write-Step "Done. Restart WPS."
    if ($Json) { $summary | ConvertTo-Json -Depth 5 }
    exit 0
}

# --------------------------------------------------------------------------- #
# 6. Repair (default) : local copy, or -Online registration
# --------------------------------------------------------------------------- #
$summary.action = if ($Online) { 'register-online' } else { 'register-local' }
Write-Step "=== Registering Tide add-in for WPS (mode: $($summary.action)) ==="

if (-not $payload) {
    throw 'add-in payload not found. Re-run the Tide installer, or pass -SourceDir <folder containing manifest.xml>.'
}

New-Item -ItemType Directory -Force -Path $jsaddonsDir | Out-Null

if (-not $Online) {
    try {
        # drop previously registered copies (other versions included) so only one Tide entry can load
        foreach ($old in @(Get-ChildItem -LiteralPath $jsaddonsDir -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like "$PluginName*" })) {
            Remove-Item -LiteralPath $old.FullName -Recurse -Force
        }
        if (Test-Path -LiteralPath $targetDir) { Remove-Item -LiteralPath $targetDir -Recurse -Force }
        Copy-Item -LiteralPath $payload -Destination $targetDir -Recurse -Force
    } catch {
        throw "cannot write $targetDir ($($_.Exception.Message)). Quit WPS completely (tray icon too) and re-run."
    }
    $fileCount = @(Get-ChildItem -LiteralPath $targetDir -Recurse -File).Count
    if ($fileCount -lt 10) { throw "payload copy looks incomplete ($fileCount files) in $targetDir" }
    $summary.payloadTarget = $targetDir
    Write-Step "  payload copied -> $targetDir ($fileCount files)"
}

$doc = Read-PublishDocument $publishPath
if ((Test-Path -LiteralPath $publishPath) -and -not (Test-Path -LiteralPath $backupPath)) {
    Copy-Item -LiteralPath $publishPath -Destination $backupPath -Force
    Write-Step "  backup written -> $backupPath"
}
Remove-TideEntries $doc

if ($Online) {
    if (-not $OnlineUrl.EndsWith('/')) { $OnlineUrl = "$OnlineUrl/" }
    $node = $doc.CreateElement('jspluginonline')
    $node.SetAttribute('name', $PluginName)
    $node.SetAttribute('type', $AddonType)
    $node.SetAttribute('url', $OnlineUrl)
    $node.SetAttribute('debug', '')
    $node.SetAttribute('enable', 'enable_dev')
    $node.SetAttribute('install', 'null')
    $node.SetAttribute('customDomain', '')
    Write-Step "  online entry -> $OnlineUrl"
} else {
    $node = $doc.CreateElement('jsplugin')
    $node.SetAttribute('name', $PluginName)
    $node.SetAttribute('type', $AddonType)
    $node.SetAttribute('url', $folderName)
    $node.SetAttribute('version', $version)
    $node.SetAttribute('enable', 'enable_dev')
    $node.SetAttribute('install', 'null')
    $node.SetAttribute('customDomain', '')
    Write-Step "  local entry  -> $folderName"
}
[void]$doc.DocumentElement.AppendChild($node)
Save-PublishDocument $doc $publishPath
Write-Step "  publish.xml written -> $publishPath"

# --------------------------------------------------------------------------- #
# 7. Post-write verification
# --------------------------------------------------------------------------- #
Write-Step ""
Write-Step "-- verify --"
$check = New-Object System.Xml.XmlDocument
$check.Load($publishPath)
$found = $false
foreach ($n in @($check.DocumentElement.ChildNodes)) {
    if ($n.NodeType -eq 'Element' -and $n.HasAttribute('name') -and $n.GetAttribute('name') -ieq $PluginName) { $found = $true }
}
if (-not $found) { Add-Problem 'verification failed: TideTermAudit entry not found after write' }

if (-not $Online) {
    foreach ($f in @('manifest.xml', 'index.html', 'main.js', 'ribbon.xml')) {
        if (-not (Test-Path -LiteralPath (Join-Path $targetDir $f))) { Add-Problem "verification failed: missing $f in $targetDir" }
    }
    Write-Step "  payload files : $(@(Get-ChildItem -LiteralPath $targetDir -Recurse -File).Count) files present"
}
Write-Step "  entries now   : "
foreach ($e in @(Get-PublishEntries $publishPath)) { Write-Step "    $e" }

$summary.ok = ($summary.problems.Count -eq 0)

Write-Step ""
if ($summary.ok) {
    Write-Step 'RESULT: registered successfully.'
} else {
    Write-Step 'RESULT: repair incomplete, see problems above.'
}
if ($summary.wpsRunning.Count) {
    Write-Step ''
    Write-Step 'WPS is still running. It reads publish.xml only at startup:'
    Write-Step '  1) save your work, close every WPS window,'
    Write-Step '  2) right-click the WPS tray icon -> Quit (a stray wps.exe keeps the ribbon cached),'
    Write-Step '  3) reopen WPS *Spreadsheets* (et) - the tab appears only there, not in Writer/Presentation.'
} else {
    Write-Step 'Open WPS Spreadsheets (et) - the tab "Tide 术语审校" should be on the ribbon.'
}
Write-Step 'First use of the side pane asks for the termbase passphrase (defined on the source machine).'

if ($Json) { $summary | ConvertTo-Json -Depth 5 }
if ($summary.ok) { exit 0 } else { exit 2 }

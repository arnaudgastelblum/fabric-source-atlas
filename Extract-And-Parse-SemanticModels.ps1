#Requires -Version 7.2

<#
.SYNOPSIS
One command: downloads semantic models from Fabric (read-only), then documents their sources.
.DESCRIPTION
Run it with /? for the options and examples.
#>
# Convenience launcher with no logic of its own and no network code. It runs, in order:
#   1. Extract-SemanticModels.ps1   the only script that connects to Fabric (read-only, allow-listed requests)
#   2. Parse-SemanticModels.ps1     offline analysis of the exported files -> model_sources.xlsx / .csv + query files
[CmdletBinding()]
param(
    [Parameter(Position = 0)][ValidateNotNullOrEmpty()][string[]]$Workspace,
    [switch]$AllWorkspaces,
    [ValidateNotNullOrEmpty()][string[]]$Exclude,
    [string[]]$Model,
    [switch]$NewOnly,
    [ValidateScript({ if ($_ -cmatch '\A[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\z') { $true } else { throw "-TenantId expects the tenant id, a GUID such as 00000000-0000-0000-0000-000000000000 (az account show --query tenantId), not '$_'." } })]
    [string]$TenantId,
    [ValidateNotNullOrEmpty()][string]$ExportPath = (Join-Path $PSScriptRoot 'data\export'),
    [ValidateNotNullOrEmpty()][string]$OutputPath = (Join-Path $PSScriptRoot 'data\sources'),
    [switch]$AdoptOutputFolder,
    [Alias('h')][switch]$Help
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# ------------------------------------------------------------------ console display (begin)
# Identical block in the three scripts (tests\Test-Extract.ps1 checks it). Colours and symbols on an interactive
# console only: redirected output and NO_COLOR get plain text. The status line (spinner, bar, current item) is drawn
# by a background thread, one frame every 100 ms and asleep in between, so that it keeps moving while a request or an
# analysis blocks the script. -Verbose and -Debug turn it off, so that their messages are never overwritten.
function Initialize-Ui([string]$Origin) {
    $interactive = -not [Console]::IsOutputRedirected -and $Host.Name -eq 'ConsoleHost'
    $color = $interactive -and -not $env:NO_COLOR -and $PSStyle.OutputRendering -ne 'PlainText'
    # Braille spinner, check marks and clickable paths need a recent terminal: Windows Terminal, VS Code, or not Windows.
    $unicode = $interactive -and [bool]($env:WT_SESSION -or $env:TERM_PROGRAM -or -not $IsWindows)
    $e = [char]27
    $script:Ui = @{
        Live = $interactive -and $VerbosePreference -eq 'SilentlyContinue' -and $DebugPreference -eq 'SilentlyContinue'
        Links = $unicode -and $color; Status = $null; Worker = $null; Started = [Diagnostics.Stopwatch]::StartNew()
        # Errors are shown as a short message (exit code 1) when a person typed the command; a calling script gets the exception.
        TopLevel = $Origin -eq 'Runspace'
        # Set by Extract-And-Parse-SemanticModels.ps1: its two steps then print "[1/2]" / "[2/2]" and one summary at the end.
        Run = Get-Variable -Name SemanticModelToolsRun -ValueOnly -ErrorAction Ignore
    }
    foreach ($style in @{ Bold = 1; Dim = 90; Green = 32; Yellow = 33; Red = 31; Cyan = 36; Reset = 0 }.GetEnumerator()) { $script:Ui[$style.Key] = if ($color) { "$e[$($style.Value)m" } else { '' } }
    $marks = if ($unicode) { @{ OkMark = '✓'; WarnMark = '!'; FailMark = '✗'; SkipMark = '↷'; Rule = '─'; Dot = '·'; Ellipsis = '…'; BarFull = '━'; BarEmpty = '━'; Frames = @('⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏') } }
             else { @{ OkMark = '[OK]'; WarnMark = '[!!]'; FailMark = '[XX]'; SkipMark = '[--]'; Rule = '-'; Dot = '-'; Ellipsis = '~'; BarFull = '#'; BarEmpty = '.'; Frames = @('|', '/', '-', '\') } }
    foreach ($mark in $marks.GetEnumerator()) { $script:Ui[$mark.Key] = $mark.Value }
}
function Format-UiText([string]$Style, [string]$Text) { return "$($script:Ui[$Style])$Text$($script:Ui.Reset)" }
function Format-UiName([string]$Name, [int]$Width) {
    if ($Name.Length -gt $Width) { return $Name.Substring(0, $Width - 1) + $script:Ui.Ellipsis }
    return $Name.PadRight($Width)
}
function Format-UiCount([int]$Count, [string]$Noun, [int]$Pad = 0) {
    # "1 table" / "12 tables"; with -Pad, the number is right-aligned and the singular padded, so that columns line up.
    $word = if ($Count -eq 1) { $Noun } else { "${Noun}s" }
    if ($Pad) { $word = $word.PadRight($Noun.Length + 1) }
    return "$("$Count".PadLeft($Pad)) $word"
}
function Format-UiDuration([double]$Seconds) {
    if ($Seconds -lt 10) { return '{0:0.0} s' -f $Seconds }
    if ($Seconds -lt 60) { return '{0:0} s' -f [math]::Floor($Seconds) }
    return '{0} min {1:00} s' -f [math]::Floor($Seconds / 60), [math]::Floor($Seconds % 60)
}
function Format-UiPath([string]$Path) {
    # Relative to the current folder when below it; clickable (OSC 8 hyperlink) in terminals that support it.
    $shown = $Path
    try {
        $relative = [IO.Path]::GetRelativePath($PWD.ProviderPath, $Path)
        if (-not $relative.StartsWith('..') -and -not [IO.Path]::IsPathRooted($relative)) { $shown = $relative }
    } catch { $shown = $Path }
    if (-not $script:Ui.Links) { return $shown }
    $e = [char]27
    return "$e]8;;$(([uri]$Path).AbsoluteUri)$e\$shown$e]8;;$e\"
}
function Write-UiLine([string]$Text = '') {
    # Under the status line's lock: the line being drawn is cleared first, then redrawn below by the worker.
    $status = $script:Ui.Status
    if ($null -eq $status) { Write-Host $Text; return }
    [Threading.Monitor]::Enter($status.Lock)
    try { $Host.UI.Write("`r$([char]27)[2K"); Write-Host $Text } finally { [Threading.Monitor]::Exit($status.Lock) }
}
function Write-UiWarning([string]$Message) {
    $status = $script:Ui.Status
    if ($null -eq $status) { Write-Warning $Message; return }
    [Threading.Monitor]::Enter($status.Lock)
    try { $Host.UI.Write("`r$([char]27)[2K"); Write-Warning $Message } finally { [Threading.Monitor]::Exit($status.Lock) }
}
function Format-UiLink([string]$Url, [string]$Text) {
    if (-not $script:Ui.Links) { return $Text }
    $e = [char]27
    return "$e]8;;$Url$e\$(Format-UiText Cyan $Text)$e]8;;$e\"
}
# Lazy Snail logo, drawn with braille dots (16 x 11 characters). Dots: the characters of each line; Colours: one letter per
# character, '.' for none, a..f = the soft colours below (RGB hex). Shown only where colours and Unicode are.
$script:UiLogo = @{
    Palette = 'dde884 d0dc77 cace71 e1b065 b2ae55 bd9a52'
    Dots    = @(
        '   ⡀ ⢀⣛⣷⡄       '
        '  ⣾⣿⣶⠟⠉⠙⣷⡀      '
        ' ⢰⣟⢀⣀⠰⣶⠆⢸⡇      '
        '  ⠻⣮⣥⢠⣤⠠⣼⠇      '
        ' ⣀⣴⠾⠟⠉⡙⣼⣿⣤⣤⣤⣀   '
        '⢾⡉⠻⢤⣂⣠⠰⡞⠩⢀⠠ ⠩⠻⣦ '
        '⣸⠟ ⠂ ⠄⣚⢡⣼⡶⠖⡾⣄⠂⠹⣧'
        '⣿ ⠠⠈⠐⡈⠇⣾⠇⠄⣵ ⢼⡏⢐⣿'
        '⣿ ⠐⡈⠄⠐⢲⠹⣧⡈⠝⠻⠋⠌⣴⡿'
        '⠘⢧⡠⠐ ⠅⠂⠹⢬⣷⣮⣬⣼⣿⣟⠁'
        '  ⠙⠷⢧⣬⣤⣥⣬⣤⣽⣭⣽⠾⠋ '
    )
    Colours = @(
        '...c.bbcc.......'
        '..bbbbbccc......'
        '.cbbbbbbcc......'
        '..aaabbbcc......'
        '.aabbbbcdddad...'
        'aaabbcceddd.ddd.'
        'aa.b.eefdddddddf'
        'a.bceeeffdd.ddff'
        'a.bceeefffdfffff'
        'aace.eeeffffffff'
        '..cceeeeeeffffe.'
    )
}
function Get-UiLogoLines {
    $e = [char]27
    $colours = @($script:UiLogo.Palette -split ' ' | ForEach-Object { '{0};{1};{2}' -f [Convert]::ToInt32($_.Substring(0, 2), 16), [Convert]::ToInt32($_.Substring(2, 2), 16), [Convert]::ToInt32($_.Substring(4, 2), 16) })
    for ($y = 0; $y -lt $script:UiLogo.Dots.Count; $y++) {
        $dots = $script:UiLogo.Dots[$y]; $letters = $script:UiLogo.Colours[$y]
        $line = [Text.StringBuilder]::new(); $current = '.'
        for ($x = 0; $x -lt $dots.Length; $x++) {
            # One colour code per run of cells of the same colour.
            if ($letters[$x] -ne '.' -and $letters[$x] -ne $current) { [void]$line.Append("$e[38;2;$($colours[[int]$letters[$x] - 97])m"); $current = $letters[$x] }
            [void]$line.Append($dots[$x])
        }
        [void]$line.Append("$e[0m")
        $line.ToString()
    }
}
function Write-UiHeader([string]$Title, [string]$Detail, [string]$Summary) {
    $detailText = if ($Detail) { "   $(Format-UiText Dim $Detail)" } else { '' }
    $run = $script:Ui.Run
    if ($run) { $run.Step++; Write-UiLine "  $(Format-UiText Cyan "[$($run.Step)/$($run.Steps)]") $(Format-UiText Bold $Title)$detailText"; Write-UiLine; return }
    $text = @("$(Format-UiText Bold $Title)$detailText") + @($Summary | Where-Object { $_ }) + @("$(Format-UiText Dim 'by') Arnaud GASTELBLUM $(Format-UiText Dim $script:Ui.Dot) $(Format-UiLink 'https://lazysnail.net' 'lazysnail.net')")
    Write-UiLine
    $width = 0; try { $width = [Console]::WindowWidth } catch { $width = 0 }
    if ($script:Ui.Links -and $width -ge 70) {
        # Logo on the left, title and credit beside its middle.
        $logo = @(Get-UiLogoLines); $at = [int][math]::Floor(($logo.Count - $text.Count) / 2)
        for ($i = 0; $i -lt $logo.Count; $i++) { Write-UiLine "  $($logo[$i])$(if ($i -ge $at -and $i -lt $at + $text.Count) { "    $($text[$i - $at])" })" }
        Write-UiLine
    } else { foreach ($line in $text) { Write-UiLine "  $line" } }
    Write-UiLine "  $(Format-UiText Dim ($script:Ui.Rule * 64))"; Write-UiLine
}
function Write-UiHelp([string]$Script, [string]$Summary, $Examples) {
    # $Examples: ordered arguments -> what the command does. Shown on /?, -h, --help, help (and without argument where a workspace is required).
    Write-UiHeader $Script -Summary $Summary
    Write-UiLine "  $(Format-UiText Bold 'Examples')"
    Write-UiLine
    foreach ($example in $Examples.GetEnumerator()) {
        Write-UiLine "    $(Format-UiText Cyan (".\$Script $($example.Key)".TrimEnd()))"
        Write-UiLine "      $($example.Value)"
        Write-UiLine
    }
    Write-UiLine "  $(Format-UiText Dim 'Names accept the wildcards * and ?. Details and FAQ: README.md')"
    Write-UiLine
}
function Write-UiSummary($Rows) {
    # $Rows: ordered label -> value. Elapsed time since the start of the launcher when there is one.
    $elapsed = if ($script:Ui.Run) { $script:Ui.Run.Started.Elapsed } else { $script:Ui.Started.Elapsed }
    Write-UiLine
    Write-UiLine "  $(Format-UiText Dim ($script:Ui.Rule * 64))"
    Write-UiLine "  $(Format-UiText Green 'Done') in $(Format-UiDuration $elapsed.TotalSeconds)"
    Write-UiLine
    foreach ($row in $Rows.GetEnumerator()) { Write-UiLine "  $(Format-UiText Dim ([string]$row.Key).PadRight(13))$($row.Value)" }
    Write-UiLine "  $(Format-UiText Dim ($script:Ui.Rule * 64))"
    Write-UiLine
}
function Write-UiFailure($ErrorRecord) {
    Stop-UiStatus
    Write-UiLine "  $(Format-UiText Red "$($script:Ui.FailMark)  $($ErrorRecord.Exception.Message)")"
    Write-UiLine
    Write-Verbose "$($ErrorRecord.ScriptStackTrace)"
}
function Set-UiStatus([string]$Label, [int]$Done = 0, [int]$Total = 0, [string]$Item = '') {
    # Total 0: spinner only (nothing to count yet). The first call starts the worker thread.
    if (-not $script:Ui.Live) { return }
    $status = $script:Ui.Status
    if ($null -eq $status) {
        $status = [hashtable]::Synchronized(@{ Lock = [object]::new(); Stop = $false; Label = $Label; Done = $Done; Total = $Total; Item = $Item })
        $shell = [powershell]::Create()
        [void]$shell.AddScript({
                param($Status, $Ui, $HostUi)
                $e = [char]27; $frame = 0
                $HostUi.Write("$e[?25l")   # hide the cursor while the line is animated
                while (-not $Status.Stop) {
                    [Threading.Monitor]::Enter($Status.Lock)
                    try {
                        $width = 100; try { $width = [Console]::WindowWidth - 1 } catch { $width = 100 }
                        $text = "  $($Ui.Cyan)$($Ui.Frames[$frame % $Ui.Frames.Count])$($Ui.Reset) $($Status.Label)"
                        $used = 4 + $Status.Label.Length
                        if ($Status.Total -gt 0) {
                            $filled = [int][math]::Floor(24 * [math]::Min($Status.Done, $Status.Total) / $Status.Total)
                            $counter = "$($Status.Done)/$($Status.Total)"
                            $text += "  $($Ui.Cyan)$($Ui.BarFull * $filled)$($Ui.Dim)$($Ui.BarEmpty * (24 - $filled))$($Ui.Reset)  $counter"
                            $used += 28 + $counter.Length
                        }
                        $item = [string]$Status.Item; $room = $width - $used - 2
                        if ($item.Length -gt $room) { $item = if ($room -gt 3) { $item.Substring(0, $room - 1) + $Ui.Ellipsis } else { '' } }
                        if ($item) { $text += "  $($Ui.Dim)$item$($Ui.Reset)" }
                        $HostUi.Write("`r$e[2K$text")
                    } finally { [Threading.Monitor]::Exit($Status.Lock) }
                    $frame++
                    [Threading.Thread]::Sleep(100)
                }
                [Threading.Monitor]::Enter($Status.Lock)
                try { $HostUi.Write("`r$e[2K$e[?25h") } finally { [Threading.Monitor]::Exit($Status.Lock) }
            }).AddArgument($status).AddArgument($script:Ui).AddArgument($Host.UI)
        $script:Ui.Worker = @{ Shell = $shell; Handle = $shell.BeginInvoke() }
        $script:Ui.Status = $status
        return
    }
    $status.Label = $Label; $status.Done = $Done; $status.Total = $Total; $status.Item = $Item
}
function Stop-UiStatus {
    # Called on every way out (end, error, Ctrl+C through the callers' finally): the worker clears its line and shows the cursor again.
    $status = $script:Ui.Status
    if ($null -eq $status) { return }
    $status.Stop = $true
    [void]$script:Ui.Worker.Handle.AsyncWaitHandle.WaitOne(2000)
    $script:Ui.Worker.Shell.Dispose()
    $script:Ui.Status = $null; $script:Ui.Worker = $null
}
Initialize-Ui $MyInvocation.CommandOrigin
trap { if ($script:Ui.TopLevel) { Write-UiFailure $_; exit 1 }; Stop-UiStatus; break }
# ------------------------------------------------------------------ console display (end)
if (-not $IsWindows) { throw 'This tool runs on Windows only, with PowerShell 7.2 or later (pwsh). Run it from a Windows computer.' }
$usage = [ordered]@{
    "-Workspace 'Finance'"                   = 'Exports the models of Finance, then lists what their tables read. Overwrites earlier exports and results (see -NewOnly).'
    "-Workspace 'Finance' -Model 'Sales*'"   = 'Only the models whose name starts with Sales. The other models already in the workbook are kept.'
    '-AllWorkspaces'                         = 'Every workspace you can export. My workspace and Viewer-only ones are skipped.'
    "-AllWorkspaces -Exclude 'Sandbox*', 'Archive'" = 'Every workspace except those.'
    '-AllWorkspaces -NewOnly'                = 'Only the models not exported and analysed yet: earlier exports and results are kept, not overwritten. Faster.'
    "-Workspace 'Finance' -TenantId '<tenant id>'" = 'Signs in to that tenant (az --tenant), when your account belongs to several.'
    "-Workspace 'Finance' -ExportPath 'D:\Exports' -OutputPath 'D:\Inventory'" = 'Other folders. They must be new, empty, or already used by these scripts.'
    "-Workspace 'Finance' -OutputPath 'D:\Old' -AdoptOutputFolder" = 'Accepts folders that already hold other files. Their model folders get replaced: be sure.'
}
# No argument, or /?, -h, --help or help typed by habit (PowerShell binds them to -Workspace): show the help.
if ($Help -or -not $PSBoundParameters.Count -or @($Workspace | Where-Object { $_ -in '/?', '?', '-h', '--help', 'help' }).Count) {
    Write-UiHelp 'Extract-And-Parse-SemanticModels.ps1' 'Both steps in one go: exports the models from Fabric (read-only), then lists what their tables read in Excel.' $usage
    return
}
if ([bool]$Workspace -eq $AllWorkspaces.IsPresent) { throw 'Name the workspace(s) with -Workspace, or use -AllWorkspaces (not both). Run with /? for help.' }
Write-UiHeader 'Fabric Source Atlas'
# Read by the two scripts below (see the console display block): steps numbered [1/2] / [2/2], one summary at the end.
$SemanticModelToolsRun = @{ Step = 0; Steps = 2; Started = $script:Ui.Started }

$exportArguments = @{ OutputPath = $ExportPath; AdoptOutputFolder = $AdoptOutputFolder.IsPresent; NewOnly = $NewOnly.IsPresent }
if ($Workspace) { $exportArguments.Workspace = $Workspace } else { $exportArguments.AllWorkspaces = $true }
if ($Exclude) { $exportArguments.Exclude = $Exclude }
if ($Model) { $exportArguments.Model = $Model }
if ($TenantId) { $exportArguments.TenantId = $TenantId }
$exported = @(& (Join-Path $PSScriptRoot 'Extract-SemanticModels.ps1') @exportArguments)
if (-not $exported.Count) { throw 'No semantic model matches -Workspace / -Model in Fabric: nothing to analyse.' }
# A row without ModelId is a workspace whose models could not be listed; its Folder is the workspace folder.
$models = @($exported | Where-Object ModelId)
$failed = @($models | Where-Object Status -EQ 'Failed')
$unlisted = @($exported | Where-Object { -not $_.ModelId })
if ($failed.Count) { Write-UiWarning "$($failed.Count) of $($models.Count) model(s) could not be exported (see _export_report.csv). A previous export of those models, if any, is analysed as is." }
if ($unlisted.Count) { Write-UiWarning "$($unlisted.Count) workspace(s) could not be listed. A previous export of their models, if any, is analysed as is." }
if ($failed.Count -or $unlisted.Count) { Write-UiLine }
if (-not $models.Count) { throw 'No workspace could be listed in Fabric (see the warnings above): nothing to analyse.' }

# The analysis selects workspaces by FOLDER name, which differs from the display name when it contains
# characters that are not valid in a file name: take the folder names from the export results.
# Excluded, personal and Viewer-only workspaces have no result, so they are not analysed either.
$workspaceFolders = @($exported | ForEach-Object { if ($_.ModelId) { Split-Path (Split-Path $_.Folder -Parent) -Leaf } else { Split-Path $_.Folder -Leaf } } | Select-Object -Unique)
$sourceArguments = @{ ExportPath = $ExportPath; OutputPath = $OutputPath; Workspace = $workspaceFolders; AdoptOutputFolder = $AdoptOutputFolder.IsPresent; NewOnly = $NewOnly.IsPresent }
if ($Model) { $sourceArguments.Model = $Model }
& (Join-Path $PSScriptRoot 'Parse-SemanticModels.ps1') @sourceArguments

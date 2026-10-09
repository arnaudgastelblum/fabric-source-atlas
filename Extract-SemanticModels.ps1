#Requires -Version 7.2

<#
.SYNOPSIS
Downloads the definition (TMDL) of Fabric / Power BI semantic models to local folders. Read-only towards Fabric.
.DESCRIPTION
Run it with /? for the options and examples.
#>
# READ-ONLY towards Fabric / Power BI: every request goes through Invoke-FabricRequest, whose allow-list only
# permits GET on workspaces, semantic models and operations, plus the body-less POST getDefinition (a read
# implemented as a long-running operation). The previous export of a model is replaced only after its new
# definition has been downloaded and validated. Never point -OutputPath into a OneLake File Explorer folder
# (refused) or into a git working tree connected to Fabric Git integration: local files there can flow back to Fabric.
[CmdletBinding(DefaultParameterSetName = 'Export')]
param(
    [Parameter(ParameterSetName = 'Export', Position = 0)]
    [Parameter(ParameterSetName = 'ListModels', Position = 0)]
    [ValidateNotNullOrEmpty()][string[]]$Workspace,
    [Parameter(ParameterSetName = 'Export')][Parameter(ParameterSetName = 'ListModels')][switch]$AllWorkspaces,
    [Parameter(ParameterSetName = 'Export')][Parameter(ParameterSetName = 'ListModels')][ValidateNotNullOrEmpty()][string[]]$Exclude,
    [Parameter(ParameterSetName = 'Export')][Parameter(ParameterSetName = 'ListModels')][string[]]$Model,
    [Parameter(ParameterSetName = 'Export')][switch]$NewOnly,
    # Default: the data\export folder next to this script, wherever the command is typed from.
    [Parameter(ParameterSetName = 'Export')][ValidateNotNullOrEmpty()][string]$OutputPath = (Join-Path $PSScriptRoot 'data\export'),
    [Parameter(ParameterSetName = 'ListWorkspaces', Mandatory)][switch]$ListWorkspaces,
    [Parameter(ParameterSetName = 'ListModels', Mandatory)][switch]$ListModels,
    [Parameter(ParameterSetName = 'Export')][ValidateRange(1, 2147483647)][int]$PollTimeoutSeconds = 600,
    [Parameter(ParameterSetName = 'Export')][switch]$AdoptOutputFolder,
    # Passed to az as --tenant, for accounts that belong to several tenants. Nothing else is done with it.
    [ValidateScript({ if ($_ -cmatch '\A[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\z') { $true } else { throw "-TenantId expects the tenant id, a GUID such as 00000000-0000-0000-0000-000000000000 (az account show --query tenantId), not '$_'." } })]
    [string]$TenantId,
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
# Windows only: the output-folder guards below (drive-letter paths, reparse points, OneLake / OneDrive folders) are Windows rules.
if (-not $IsWindows) { throw 'This tool runs on Windows only, with PowerShell 7.2 or later (pwsh). Run it from a Windows computer.' }
$usage = [ordered]@{
    '-ListWorkspaces'                         = 'Lists the workspaces your account can see. Writes nothing.'
    "-Workspace 'Finance' -ListModels"        = 'Lists the semantic models of the Finance workspace. Writes nothing.'
    "-Workspace 'Finance'"                    = 'Exports every model of Finance to the data\export folder next to this script. Overwrites the models already exported there (see -NewOnly).'
    "-Workspace 'Finance' -Model 'Sales*'"    = 'Only the models whose name starts with Sales. The report keeps the lines of the other models.'
    "-AllWorkspaces -Exclude 'Sandbox*'"      = 'Every workspace you can export, except Sandbox*. My workspace and Viewer-only ones are skipped.'
    '-AllWorkspaces -NewOnly'                 = 'Only the models not exported yet: the existing exports are kept, not overwritten. Faster.'
    "-Workspace 'Finance' -TenantId '<tenant id>'" = 'Signs in to that tenant (az --tenant), when your account belongs to several.'
    "-Workspace 'Finance' -OutputPath 'D:\Exports'" = 'Writes elsewhere. The folder must be new, empty, or already used by these scripts.'
    "-Workspace 'Finance' -OutputPath 'D:\Old' -AdoptOutputFolder" = 'Accepts a folder that already holds other files. Its model folders get replaced: be sure.'
    "-Workspace 'Finance' -PollTimeoutSeconds 1800" = 'Waits up to 30 minutes per model instead of 10, for very large models.'
}
# No argument, or /?, -h, --help or help typed by habit (PowerShell binds them to -Workspace): show the help, before anything else.
if ($Help -or -not $PSBoundParameters.Count -or @($Workspace | Where-Object { $_ -in '/?', '?', '-h', '--help', 'help' }).Count) {
    Write-UiHelp 'Extract-SemanticModels.ps1' 'Downloads the definitions of your Power BI / Fabric semantic models. Read-only: changes nothing in Fabric.' $usage
    return
}
if ($PSCmdlet.ParameterSetName -ne 'ListWorkspaces' -and [bool]$Workspace -eq $AllWorkspaces.IsPresent) { throw 'Name the workspace(s) with -Workspace, or use -AllWorkspaces (not both). Run with /? for help.' }
$script:FabricToken = $null
$script:FabricTokenRefreshAt = [datetime]::MinValue
$script:SignedInShown = $false
$script:FabricBase = 'https://api.fabric.microsoft.com/v1'
# What the status line shows, so that a wait (throttling, retry, long-running operation) can put it back afterwards.
$script:StatusNow = @{ Label = ''; Done = 0; Total = 0; Item = '' }
# Reserved Windows device names, with or without extension (superscript digits included).
$script:DeviceName = '\A(?i:CON|PRN|AUX|NUL|COM[0-9¹²³]|LPT[0-9¹²³])(\..*)?\z'
# The only file types a semantic model definition is expected to contain. Anything else is refused, so a hostile
# response can never plant a script or an executable (.ps1, .cmd, .exe ...) anywhere, whatever the folder layout.
$script:PartExtensions = [Collections.Generic.HashSet[string]]::new([string[]]('.tmdl', '.json', '.pbism', '.platform', '.dax', '.md', '.txt'), [StringComparer]::OrdinalIgnoreCase)
# Folder names are cut at this length (plus an id suffix), so that no path component comes near the 255-character limit.
$script:MaxFolderName = 80
# Longest total wait (seconds) on HTTP 429 for one request before it is given up.
$script:MaxThrottleWait = 1200

function Get-OptionalProperty($Object, [string]$Name) {
    if ($null -ne $Object) {
        $property = $Object.PSObject.Properties[$Name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}
function Test-NameMatch([string]$Name, [string]$Pattern) {
    # Escape the backtick and the brackets only: * and ? retain their wildcard meaning.
    return $Name -ilike $Pattern.Replace('`', '``').Replace('[', '`[').Replace(']', '`]')
}
function ConvertTo-FolderName([string]$Name, [string]$Id) {
    $safe = [regex]::Replace($Name, '[\\/:*?"<>|\p{Cc}]', '_').TrimEnd('.', ' ')
    if ($safe.Length -gt $script:MaxFolderName) {
        # Long name: cut (never inside a surrogate pair) and tell it apart with the start of the id.
        $cut = $script:MaxFolderName
        if ([char]::IsHighSurrogate($safe[$cut - 1])) { $cut-- }
        $safe = $safe.Substring(0, $cut).TrimEnd('.', ' ')
        if ($safe) { $safe = "$safe ($($Id.Substring(0, 8)))" }
    }
    if ([string]::IsNullOrEmpty($safe)) { return $Id }
    if ($safe -match $script:DeviceName) { $safe = "_$safe" }
    return $safe
}
function Get-FolderNames($Items, [int]$SuffixLength) {
    # Id -> folder name. Deterministic whatever order Fabric lists the items in: ids are taken in ordinal order, the first
    # item of a set of colliding names keeps the plain name, the next ones get " (<id or its first characters>)".
    $byId = @{}
    foreach ($item in @($Items)) { $byId[[string]$item.id] = $item }
    $ids = [string[]]@($byId.Keys)
    [Array]::Sort($ids, [StringComparer]::OrdinalIgnoreCase)
    $taken = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $names = @{}
    foreach ($id in $ids) {
        $name = ConvertTo-FolderName ([string](Get-OptionalProperty $byId[$id] 'displayName')) $id
        $suffix = if ($SuffixLength) { $id.Substring(0, $SuffixLength) } else { $id }
        while (-not $taken.Add($name)) { $name += " ($suffix)" }
        $names[$id] = $name
    }
    return $names
}
function Show-Status([string]$Label, [int]$Done = 0, [int]$Total = 0, [string]$Item = '') {
    $script:StatusNow = @{ Label = $Label; Done = $Done; Total = $Total; Item = $Item }
    Set-UiStatus $Label $Done $Total $Item
}
function Restore-Status { $s = $script:StatusNow; if ($s.Label) { Set-UiStatus $s.Label $s.Done $s.Total $s.Item } }

# ------------------------------------------------------------------ sign-in (Azure CLI)
function Hide-Secret([string]$Text) {
    # Nothing shaped like a token (a JWT starts with eyJ) may reach the console, a file or an exception message.
    $clean = [regex]::Replace([string]$Text, 'eyJ[A-Za-z0-9_\-]*(\.[A-Za-z0-9_\-]*){0,2}', '<hidden>')
    if ($script:FabricToken) { $clean = $clean.Replace($script:FabricToken, '<hidden>') }
    return $clean
}
function Get-TokenClaims([string]$Token) {
    # Reads the payload of the access token locally (base64url JSON, no network, no signature check): only to show who is
    # signed in and to which tenant, and to read the expiry when az does not give it. Returns $null when unreadable.
    try {
        $segments = $Token.Split('.')
        if ($segments.Count -ne 3) { return $null }
        $payload = $segments[1].Replace('-', '+').Replace('_', '/')
        $payload += '=' * ((4 - $payload.Length % 4) % 4)
        return ConvertFrom-Json -InputObject ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)))
    } catch { return $null }
}
function Get-AzLoginCommand([string]$Extra = '') {
    $command = "az login$Extra"
    if ($TenantId) { $command += " --tenant $TenantId" }
    return $command
}
function Get-AzFailureMessage([string[]]$Errors, [int]$ExitCode) {
    # az's own words (stderr only: stdout would hold the token), first AADSTS line when there is one.
    $lines = @($Errors | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    $text = $lines -join "`n"
    $aad = @($lines | Where-Object { $_ -match 'AADSTS\d+' } | Select-Object -First 1)
    $shown = if ($aad.Count) { $aad[0] } elseif ($lines.Count) { $lines[0] } else { "exit code $ExitCode" }
    if ($shown.Length -gt 300) { $shown = $shown.Substring(0, 299) + '…' }
    if ($text -match 'AADSTS(50076|50079|50158|50074|50072|53003|53000|530003|50005)\b' -or $text -match 'conditional access|multi-factor|interaction_required|InteractionRequired') {
        return "Azure CLI could not get a Fabric token: $shown`nYour organisation asks for a fresh sign-in for Fabric (multi-factor or conditional access). Run: $(Get-AzLoginCommand ' --scope https://api.fabric.microsoft.com/.default'), then run this script again."
    }
    if ($aad.Count) { return "Azure CLI could not get a Fabric token: $shown`nRun: $(Get-AzLoginCommand), then run this script again." }
    if ($text -match 'No subscriptions found') {
        return "Azure CLI found no Azure subscription for your account (none is needed for Fabric). Run: $(Get-AzLoginCommand ' --allow-no-subscriptions'), then run this script again."
    }
    if ($text -match "az login|not logged in|Please run|No accounts were found|token.*expired") {
        return "You are not signed in to Azure CLI (or the sign-in expired). Run: $(Get-AzLoginCommand), then run this script again."
    }
    return "Azure CLI could not get a Fabric token: $shown`nRun: $(Get-AzLoginCommand), then run this script again."
}
function Show-SignedIn($Claims) {
    $who = $null
    foreach ($claim in 'upn', 'preferred_username', 'unique_name', 'email') { if (-not $who) { $who = [string](Get-OptionalProperty $Claims $claim) } }
    if (-not $who) { $app = [string](Get-OptionalProperty $Claims 'appid'); if (-not $app) { $app = [string](Get-OptionalProperty $Claims 'azp') }; if ($app) { $who = "application $app" } }
    if (-not $who) { $who = 'an account the token does not name' }
    $tenant = [string](Get-OptionalProperty $Claims 'tid')
    # Claim values are shown as plain text: no control characters, bounded length.
    $who = ($who -replace '\p{C}', ''); if ($who.Length -gt 100) { $who = $who.Substring(0, 100) }
    $tenant = ($tenant -replace '\p{C}', ''); if ($tenant.Length -gt 60) { $tenant = $tenant.Substring(0, 60) }
    $line = "$(Format-UiText Dim 'Signed in as') $who"
    if ($tenant) { $line += "$(Format-UiText Dim ', tenant') $tenant" }
    Write-UiLine "    $line"
    if ($TenantId -and $tenant -and -not [string]::Equals($tenant, $TenantId, [StringComparison]::OrdinalIgnoreCase)) {
        Write-UiWarning "The token is for tenant $tenant, not for -TenantId $TenantId. Run: $(Get-AzLoginCommand), then run this script again."
    }
}
function Get-FabricToken([switch]$Refresh) {
    # Reused until 5 minutes before it expires (expiry from az, else from the token itself, else 20 minutes).
    if (-not $Refresh -and $script:FabricToken -and [datetime]::UtcNow -lt $script:FabricTokenRefreshAt) { return $script:FabricToken }
    $script:FabricToken = $null
    if (-not (Get-Command az -ErrorAction Ignore)) {
        throw "Azure CLI (az) was not found. Install it from https://aka.ms/installazurecliwindows, open a new terminal, run $(Get-AzLoginCommand), then run this script again."
    }
    $arguments = @('account', 'get-access-token', '--resource', 'https://api.fabric.microsoft.com', '--output', 'json')
    if ($TenantId) { $arguments += @('--tenant', $TenantId) }
    # stdout holds the token: it is parsed, never shown. stderr is az's message, shown only through Get-AzFailureMessage.
    $output = [Collections.Generic.List[string]]::new(); $errors = [Collections.Generic.List[string]]::new(); $exitCode = -1
    try {
        $ErrorActionPreference = 'Continue'   # local to this function: a line on stderr is a message to read, not a failure
        $global:LASTEXITCODE = 0
        & az @arguments 2>&1 | ForEach-Object { if ($_ -is [Management.Automation.ErrorRecord]) { $errors.Add((Hide-Secret "$_")) } else { $output.Add([string]$_) } }
        $exitCode = $global:LASTEXITCODE
    } catch { $errors.Add((Hide-Secret $_.Exception.Message)) }
    finally { $ErrorActionPreference = 'Stop' }
    $answer = $null
    if ($exitCode -eq 0 -and $output.Count) { try { $answer = ConvertFrom-Json -InputObject ($output -join "`n") } catch { $answer = $null } }
    $token = [string](Get-OptionalProperty $answer 'accessToken')
    if (-not $token) { throw (Get-AzFailureMessage $errors.ToArray() $exitCode) }
    # It goes into an HTTP header: base64url segments and dots only.
    if ($token -cnotmatch '\A[A-Za-z0-9\-_.]+\z') { throw 'Azure CLI returned an access token in an unexpected format. Update it (az upgrade), then run this script again.' }
    $claims = Get-TokenClaims $token
    $now = [datetime]::UtcNow; $expiry = $null
    # expires_on: POSIX seconds (az 2.54 and later). Then the exp claim of the token. Then expiresOn: LOCAL time, no zone.
    $seconds = 0L
    foreach ($candidate in @([string](Get-OptionalProperty $answer 'expires_on'), [string](Get-OptionalProperty $claims 'exp'))) {
        if ($null -eq $expiry -and [long]::TryParse($candidate, [Globalization.NumberStyles]::None, [cultureinfo]::InvariantCulture, [ref]$seconds)) {
            try { $expiry = [DateTimeOffset]::FromUnixTimeSeconds($seconds).UtcDateTime } catch { $expiry = $null }
        }
    }
    if ($null -eq $expiry) {
        $local = [datetime]::MinValue
        $styles = [Globalization.DateTimeStyles]::AssumeLocal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
        if ([datetime]::TryParseExact([string](Get-OptionalProperty $answer 'expiresOn'), [string[]]@('yyyy-MM-dd HH:mm:ss.ffffff', 'yyyy-MM-dd HH:mm:ss'), [cultureinfo]::InvariantCulture, $styles, [ref]$local)) { $expiry = $local }
    }
    # An expiry in the past or more than a day away is not believed.
    if ($null -ne $expiry -and ($expiry -le $now -or $expiry -gt $now.AddHours(24))) { $expiry = $null }
    if ($null -eq $expiry) { $refreshAt = $now.AddMinutes(20) }
    else {
        $refreshAt = $expiry.AddMinutes(-5)
        # A token close to its end: use it a little, never past 30 s before its expiry.
        if ($refreshAt -lt $now.AddSeconds(60)) { $refreshAt = $now.AddSeconds(60); if ($refreshAt -gt $expiry.AddSeconds(-30)) { $refreshAt = $expiry.AddSeconds(-30) } }
    }
    $script:FabricToken = $token; $script:FabricTokenRefreshAt = $refreshAt
    Write-Verbose "Fabric token acquired, reused until $($refreshAt.ToString('HH:mm:ss')) UTC."
    if (-not $script:SignedInShown) { $script:SignedInShown = $true; Show-SignedIn $claims }
    return $token
}

# ------------------------------------------------------------------ requests to Fabric (the only HTTP call site)
function Get-HeaderValue($Headers, [string]$Name) {
    # Invoke-WebRequest header values are arrays, including Location.
    if ($Headers.ContainsKey($Name)) { return [string](@($Headers[$Name])[0]) }
    return $null
}
function Get-RetryDelay($Headers, [int]$Default, [int]$Maximum, [int]$Minimum = 0) {
    # Retry-After: a number of seconds, or an HTTP date (RFC 1123).
    $value = ([string](Get-HeaderValue $Headers 'Retry-After')).Trim()
    $seconds = [double]$Default; $number = 0L; $date = [DateTimeOffset]::MinValue
    if ([long]::TryParse($value, [Globalization.NumberStyles]::None, [cultureinfo]::InvariantCulture, [ref]$number)) { $seconds = [double]$number }
    elseif ($value -and [DateTimeOffset]::TryParseExact($value, 'r', [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$date)) {
        $seconds = [math]::Ceiling(($date - [DateTimeOffset]::UtcNow).TotalSeconds)
    }
    return [math]::Clamp($seconds, [double]$Minimum, [double]$Maximum)
}
function Get-FabricError($Body, [string]$Raw) {
    $nested = Get-OptionalProperty $Body 'error'
    if ($null -ne $nested) { $Body = $nested }
    $code = Get-OptionalProperty $Body 'errorCode'
    $message = Get-OptionalProperty $Body 'message'
    if ($code -or $message) { return "${code}: $message" }
    return $Raw
}
function Test-TransportError($Exception) {
    # Network-level failures worth retrying: connection reset or refused, DNS, TLS, timeout. Never a refusal of this script.
    for ($e = $Exception; $null -ne $e; $e = $e.InnerException) {
        if ($e -is [Security.SecurityException]) { return $false }
        # An HTTP answer that Invoke-WebRequest turns into an exception (a redirect, refused with MaximumRedirection 0) is a
        # reply of the server, not a network failure: never retried. (It derives from HttpRequestException, hence first.)
        if ($e -is [Microsoft.PowerShell.Commands.HttpResponseException]) { return $false }
        if ($e -is [Net.Http.HttpRequestException] -or $e -is [Threading.Tasks.TaskCanceledException] -or $e -is [TimeoutException] -or
            $e -is [IO.IOException] -or $e -is [Net.Sockets.SocketException]) { return $true }
    }
    return $false
}
function Wait-FabricDelay([double]$Seconds, [datetime]$Deadline, [string]$Reason = '', [Diagnostics.Stopwatch]$Since = $null) {
    # $Reason: a retry, counted down on the status line. $Since: a long-running operation, its elapsed time on the status line.
    $remaining = ($Deadline - [datetime]::UtcNow).TotalSeconds
    if ($remaining -le 0) { throw 'Fabric operation timed out.' }
    $wait = [math]::Min([double]$Seconds, $remaining)
    if ($Reason) {
        Write-Verbose "${Reason}: waiting $([math]::Ceiling($wait)) s before trying again."
        if ($wait -ge 60) { Write-UiWarning "${Reason}: waiting $(Format-UiDuration $wait) before trying again." }
    }
    if ($script:Ui.Live -and ($Reason -or $Since)) {
        # One-second steps, a bounded number of them, so that the countdown or the elapsed time moves.
        $s = $script:StatusNow
        for ($left = $wait; $left -gt 0; $left -= 1) {
            if ($Reason) { Set-UiStatus "$Reason, waiting $([math]::Ceiling($left)) s" $s.Done $s.Total $s.Item }
            else { Set-UiStatus $s.Label $s.Done $s.Total "$($s.Item) $($script:Ui.Dot) Fabric is preparing it, $(Format-UiDuration $Since.Elapsed.TotalSeconds)" }
            Start-Sleep -Seconds ([math]::Min([double]1, $left))
        }
        Restore-Status
    } else { Start-Sleep -Seconds $wait }
    if ([datetime]::UtcNow -ge $Deadline) { throw 'Fabric operation timed out.' }
}
function Invoke-FabricRequest {
    param([string]$Method, [string]$Uri, [string]$Body = '', [datetime]$Deadline = [datetime]::MaxValue)
    # Printable ASCII only, then an ORDINAL host check: PowerShell's -eq/-ne are culture-aware and ignore
    # zero-width/combining characters, which would let a lookalike host receive the bearer token.
    # The refusals are SecurityExceptions: they always stop the whole run, never just one workspace or model.
    if ($Uri -cmatch '[^\x21-\x7E]') { throw [Security.SecurityException]::new('Refusing Fabric request: unexpected characters in the URL.') }
    $endpoint = [uri]$Uri
    if (-not [string]::Equals($endpoint.Scheme, 'https', [StringComparison]::Ordinal) -or
        -not [string]::Equals($endpoint.IdnHost, 'api.fabric.microsoft.com', [StringComparison]::Ordinal) -or
        -not $endpoint.IsDefaultPort -or $endpoint.UserInfo) {
        throw [Security.SecurityException]::new("Refusing Fabric request to unexpected endpoint: $Uri")
    }
    # Read-only allow-list: the only requests this script can ever send. getDefinition is a POST only
    # because it is a long-running operation; it returns the definition and changes nothing.
    $guid = '[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}'
    $signature = "$Method $($endpoint.AbsolutePath)"
    $readOnly = @(
        "\AGET /v1/workspaces\z"
        "\AGET /v1/workspaces/$guid/semanticModels\z"
        "\AGET /v1/operations/$guid(/result)?\z"
        "\APOST /v1/workspaces/$guid/semanticModels/$guid/getDefinition\z"
    )
    if (-not @($readOnly | Where-Object { $signature -cmatch $_ }).Count) { throw [Security.SecurityException]::new("Refusing request outside the read-only allow-list: $signature") }
    if ($Method -ceq 'POST' -and ($endpoint.Query -cne '?format=TMDL' -or $Body)) { throw [Security.SecurityException]::new("Refusing POST with unexpected query or body: $Uri") }
    $throttles = 0; $retries = 0; $refreshed = $false; $throttleWaited = 0.0
    $ProgressPreference = 'SilentlyContinue'   # local: no progress bar from Invoke-WebRequest over the status line
    while ($true) {
        $token = Get-FabricToken
        $remaining = ($Deadline - [datetime]::UtcNow).TotalSeconds
        if ($remaining -le 0) { throw 'Fabric operation timed out.' }
        $arguments = @{
            Method = $Method; Uri = $Uri; Headers = @{ Authorization = "Bearer $token" }
            SkipHttpErrorCheck = $true; MaximumRedirection = 0
            Debug = $false  # in debug mode Invoke-WebRequest prints the request headers, bearer token included
            # Per-request timeout capped at 300 s: HttpClient rejects timeouts above ~24 days, and [double] forces the Min(double, double) overload.
            TimeoutSec = [int][math]::Max([double]1, [math]::Min([double]300, [math]::Ceiling($remaining)))
        }
        if ($Method -eq 'POST') { $arguments.ContentType = 'application/json'; $arguments.Body = $Body }
        $response = $null; $transport = $null
        try { $response = Invoke-WebRequest @arguments }
        catch {
            if (-not (Test-TransportError $_.Exception)) { throw }
            $transport = Hide-Secret $_.Exception.Message
        }
        if ($null -ne $transport) {
            if ($retries -ge 3) { throw "Could not reach Fabric, even after 3 retries: $transport Check the network connection (VPN, proxy), then run the script again." }
            $delay = @(5, 15, 30)[$retries]; $retries++
            Wait-FabricDelay $delay $Deadline 'Connection to Fabric failed'
            continue
        }
        $status = [int]$response.StatusCode
        if ($status -eq 401 -and -not $refreshed) {
            # Token refused (expired early, revoked): one fresh token from az, the same request once more.
            $refreshed = $true; [void](Get-FabricToken -Refresh)
            Write-Verbose 'Fabric answered HTTP 401: new token, same request again.'
            continue
        }
        if ($status -eq 429 -and $throttles -lt 8) {
            # At most 20 minutes of waiting in all for one request: past that, this request fails with a clear message.
            $delay = Get-RetryDelay $response.Headers 30 600
            if ($throttleWaited + $delay -gt $script:MaxThrottleWait) {
                throw "Fabric kept throttling the requests (HTTP 429) for $(Format-UiDuration $throttleWaited); this request was given up. Wait a while (other tools may be using the same capacity), then run the script again (-NewOnly continues where it stopped)."
            }
            $throttles++; $throttleWaited += $delay
            Wait-FabricDelay $delay $Deadline 'Fabric is throttling the requests'; continue
        }
        if ($status -ge 500 -and $status -le 599 -and $retries -lt 3) {
            $delay = @(5, 15, 30)[$retries]; $retries++
            Wait-FabricDelay $delay $Deadline "Fabric answered HTTP $status"; continue
        }
        $parsed = $null
        if (-not [string]::IsNullOrWhiteSpace($response.Content)) {
            try { $parsed = ConvertFrom-Json -InputObject $response.Content }
            catch { if ($status -lt 400) { throw "Invalid JSON from Fabric (HTTP $status)." } }
        }
        if ($status -ge 400) {
            $message = Hide-Secret "HTTP ${status}: $(Get-FabricError $parsed $response.Content)"
            if ($status -eq 401) { $message += " Fabric does not accept the sign-in: run $(Get-AzLoginCommand), then run this script again." }
            throw $message
        }
        return [pscustomobject]@{ StatusCode = $status; Headers = $response.Headers; Body = $parsed }
    }
}
function Get-FabricItems([string]$Uri, [string]$Counting = '') {
    # continuationUri comes from the server: never follow a repeated one, never loop without bound.
    # $Counting: the noun shown on the status line as items arrive ("120 workspaces found").
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal); $pages = 0; $count = 0
    while ($Uri) {
        if (-not $seen.Add($Uri) -or ++$pages -gt 1000) { throw [Security.SecurityException]::new('Fabric pagination does not terminate (repeated or too many continuationUri values).') }
        $page = Invoke-FabricRequest -Method GET -Uri $Uri
        # Anything but 200 (a 202, a 204, a redirect) is an error, never an empty list.
        if ($page.StatusCode -ne 200) { throw "Unexpected answer from Fabric to a listing: HTTP $($page.StatusCode)." }
        foreach ($item in @(Get-OptionalProperty $page.Body 'value')) {
            if ($null -eq $item) { continue }
            # Ids end up in request URLs and in folder names: accept nothing but a canonical GUID.
            $id = [string](Get-OptionalProperty $item 'id')
            if ($id -cnotmatch '\A[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\z') { throw [Security.SecurityException]::new("Unexpected item id in Fabric response: '$id'") }
            $count++
            $item
        }
        if ($Counting) { $s = $script:StatusNow; Set-UiStatus $s.Label $s.Done $s.Total "$(Format-UiCount $count $Counting) found" }
        $Uri = [string](Get-OptionalProperty $page.Body 'continuationUri')
    }
}
function Get-ModelDefinition([string]$WorkspaceId, [string]$ModelId) {
    $response = Invoke-FabricRequest -Method POST -Uri "$script:FabricBase/workspaces/$WorkspaceId/semanticModels/$ModelId/getDefinition?format=TMDL"
    if ($response.StatusCode -eq 202) {
        $deadline = [datetime]::UtcNow.AddSeconds($PollTimeoutSeconds)
        $since = [Diagnostics.Stopwatch]::StartNew()
        $operationId = Get-HeaderValue $response.Headers 'x-ms-operation-id'
        $operationUri = Get-HeaderValue $response.Headers 'Location'
        if ($operationId) { $operationUri = "$script:FabricBase/operations/$([uri]::EscapeDataString($operationId))" }
        if (-not $operationUri) { throw 'Accepted definition request has no operation ID or Location.' }
        $operationUri = $operationUri.TrimEnd('/')
        while ($true) {
            Wait-FabricDelay (Get-RetryDelay $response.Headers 5 30 1) $deadline -Since $since
            $response = Invoke-FabricRequest -Method GET -Uri $operationUri -Deadline $deadline
            $state = Get-OptionalProperty $response.Body 'status'
            if ($state -eq 'Succeeded') {
                $response = Invoke-FabricRequest -Method GET -Uri "$operationUri/result" -Deadline $deadline
                break
            }
            if ($state -eq 'Failed') { throw "Fabric operation failed: $(Get-FabricError $response.Body ($response.Body | ConvertTo-Json -Depth 10 -Compress))" }
            if ($state -notin @('NotStarted', 'Running', 'Undefined')) { throw "Unexpected operation status: $state" }
        }
        Restore-Status
    }
    if ($response.StatusCode -ne 200) { throw "Unexpected definition response: HTTP $($response.StatusCode)" }
    $definition = Get-OptionalProperty $response.Body 'definition'
    if ($null -eq $definition -or $null -eq (Get-OptionalProperty $definition 'parts')) { throw 'Definition response is missing parts.' }
    return $definition
}

# ------------------------------------------------------------------ output folder guards
function Get-OneDriveRoot([string]$Path) {
    # The OneDrive folder that holds $Path ($env:OneDrive, OneDriveCommercial, OneDriveConsumer), or $null.
    # A variable naming a drive root (C:\) or anything but a plain drive-letter folder is ignored.
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/') + '\'
    foreach ($name in 'OneDrive', 'OneDriveCommercial', 'OneDriveConsumer') {
        $value = [Environment]::GetEnvironmentVariable($name)
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        try { $root = [IO.Path]::GetFullPath($value).TrimEnd('\', '/') } catch { continue }
        if ($root -cnotmatch '\A[A-Za-z]:\\[^\\]') { continue }
        if ($full.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { return $root }
    }
    return $null
}
function Assert-SafeDirectory([string]$Path) {
    # Every existing folder from $Path up to the drive root must be a plain folder:
    #  - a symbolic link or a junction is refused (writes and deletes would land elsewhere);
    #  - any other reparse point (mount point, virtualised or cloud-synced folder) is refused, EXCEPT inside OneDrive,
    #    whose files-on-demand folders all carry the ReparsePoint attribute without being links.
    # The OneLake File Explorer check comes first: OneLake is a cloud-files provider too, and stays refused.
    Assert-NotOneLake $Path
    $cursor = $Path
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $entry = Get-Item -LiteralPath $cursor -Force
            if (-not $entry.PSIsContainer) { throw "Unsafe output folder: '$cursor' is a file, not a folder. Choose another folder, for example -OutputPath C:\SemanticModels\export." }
            $linkType = [string]$entry.LinkType; $linkTarget = $entry.LinkTarget
            if ($linkType -or $null -ne $linkTarget) {
                $kind = if ($linkType) { $linkType } else { 'link' }
                throw "Unsafe output folder: '$cursor' is a $kind to '$linkTarget'. Writing through links is refused: choose a plain local folder, for example -OutputPath C:\SemanticModels\export."
            }
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -and -not (Get-OneDriveRoot $cursor)) {
                throw "Unsafe output folder: '$cursor' is a reparse point (a mount point, or a virtual or cloud-synced folder other than OneDrive). Choose a plain local folder, for example -OutputPath C:\SemanticModels\export."
            }
        }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
}
function Assert-DirectChild([string]$Parent, [string]$Child) {
    $expected = [IO.Path]::GetFullPath($Parent).TrimEnd('\', '/')
    $actual = ([string][IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Child))).TrimEnd('\', '/')
    if (-not [string]::Equals($expected, $actual, [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing path outside '$Parent': $Child" }
}

function Assert-NotOneLake([string]$Path) {
    # OneLake File Explorer syncs local writes AND deletes back to Fabric: writing there would modify the service.
    # Checked on -OutputPath and on every workspace / model folder derived from it, because a workspace or model
    # NAMED like the sync root would otherwise land inside it. A bare 'OneLake' segment is refused too (fail-safe).
    # LIMIT: this is a path-name check. A subst drive, a UNC share or an 8.3 short name that aliases the sync
    # folder is NOT detected - never point -OutputPath at such an alias.
    $full = [IO.Path]::GetFullPath($Path)
    if (@(($full -split '[\\/]') | Where-Object { $_ -like 'OneLake - *' -or $_ -eq 'OneLake' }).Count) {
        throw "Refusing to write inside a OneLake File Explorer folder (local changes there are synced to Fabric): $full. Choose a plain local folder, for example -OutputPath C:\SemanticModels\export."
    }
}

# -ListWorkspaces / -ListModels never touch the disk: every output-folder rule below applies to an export only.
$exporting = $PSCmdlet.ParameterSetName -eq 'Export'
if ($exporting) {
    $OutputPath = [IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath))
    # Plain drive-letter paths only. \\?\C:\..., \\.\... and UNC spellings can name the same folder as C:\... while
    # defeating every path comparison below (scripts inside the output folder, OneLake segments, direct-child checks).
    if ($OutputPath -cnotmatch '\A[A-Za-z]:\\') { throw "Refusing -OutputPath '$OutputPath': use a plain drive-letter path such as C:\Exports (no \\?\, \\.\ or UNC form)." }
}
function Assert-ToolsOutsideOutput([string]$Root) {
    # The scripts must never sit inside the folder they write to and clean: an exported part could otherwise
    # replace a script that is executed afterwards (the launcher runs the analysis next), and a re-export could delete them.
    if (-not $PSScriptRoot) { return }
    $rootPrefix = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $tools = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    if ($tools.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing -OutputPath '$Root': the scripts themselves are inside it. Keep the scripts outside the output folder." }
}
function Assert-OwnedOutputRoot([string]$Root, [bool]$Adopt) {
    # Work only inside a folder these scripts created, found empty, or that the user explicitly adopted: a re-export
    # deletes and rewrites sub-folders, so it must never be let loose on a folder that holds anything else
    # (a OneLake sync root reached through an alias, Documents, somebody else's files...).
    $marker = Join-Path $Root '.semantic-model-tools'
    if (-not (Test-Path -LiteralPath $Root)) { [IO.Directory]::CreateDirectory($Root) | Out-Null }
    elseif (-not (Test-Path -LiteralPath $marker -PathType Leaf) -and -not $Adopt -and @(Get-ChildItem -LiteralPath $Root -Force | Select-Object -First 1).Count) {
        throw "Refusing to use '$Root': it already contains files and was not created by these scripts. Use a new or empty folder, or pass -AdoptOutputFolder once if you are certain it is a plain local folder that is not synchronised to OneLake / Fabric."
    }
    if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) { [IO.File]::WriteAllText($marker, "Folder managed by the semantic model export scripts. Its sub-folders are deleted and rewritten on each run.`r`n") }
}
function Get-ViewerMessage([string[]]$Names) {
    $list = @($Names | ForEach-Object { "'$_'" }) -join ', '
    return "Viewer access only to ${list}: Fabric gives model definitions to workspace Admins, Members and Contributors only. Ask a workspace admin for the Contributor role, or choose another workspace (-ListWorkspaces lists them)."
}

if ($exporting) {
    Write-UiHeader 'Export from Fabric' 'read-only'
    Assert-NotOneLake $OutputPath
    Assert-SafeDirectory $OutputPath
    Assert-ToolsOutsideOutput $OutputPath
    Assert-OwnedOutputRoot $OutputPath $AdoptOutputFolder.IsPresent
    $oneDriveRoot = Get-OneDriveRoot $OutputPath
    if ($oneDriveRoot) { Write-UiWarning "$OutputPath is in OneDrive ($oneDriveRoot): the exported definitions (model code, server and database names) will be synced to OneDrive. To keep them on this computer only, choose a folder outside OneDrive, for example -OutputPath C:\SemanticModels\export." }
    $cursor = $OutputPath
    while ($cursor) {
        if (Test-Path -LiteralPath (Join-Path $cursor '.git')) {
            # The scripts' own repository ignores its data\ folder (see .gitignore): nothing to say for an output under
            # that folder. Anywhere else (another repository, or another folder of this one), warn.
            $own = $PSScriptRoot -and [string]::Equals([IO.Path]::GetFullPath($cursor).TrimEnd('\', '/'), [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase) -and
                ($OutputPath.TrimEnd('\', '/') + '\').StartsWith([IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'data')).TrimEnd('\', '/') + '\', [StringComparison]::OrdinalIgnoreCase)
            if (-not $own) { Write-UiWarning "$OutputPath is inside a git repository ($cursor). If that repository is connected to Fabric Git integration, committing these exported definitions and updating the workspace from Git would change the workspace." }
            break
        }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
}
# The status line is stopped on every way out, Ctrl+C included.
try {
    Show-Status 'Signing in with Azure CLI'
    [void](Get-FabricToken)
    Show-Status 'Listing workspaces'
    $visibleWorkspaces = @(Get-FabricItems "$script:FabricBase/workspaces" 'workspace')
    # Dispatch on the parameter SET, never on the switch value: '-ListModels:$false' still selects the ListModels set
    # (so the output-folder guards above were skipped) and must therefore still end here, before any export.
    if ($PSCmdlet.ParameterSetName -eq 'ListWorkspaces') {
        Stop-UiStatus
        Write-UiLine "    $(Format-UiText Dim "$(Format-UiCount $visibleWorkspaces.Count 'workspace') visible with the current login")"
        return @($visibleWorkspaces | Sort-Object displayName | ForEach-Object { [pscustomobject]@{ Workspace = $_.displayName; Id = $_.id; Type = [string](Get-OptionalProperty $_ 'type') } })
    }
    function Test-WorkspaceSelector($Entry, [string]$Selector) {
        $guid = [guid]::Empty
        if ([guid]::TryParse($Selector, [ref]$guid)) { return $Entry.id -eq $guid.ToString() }
        return Test-NameMatch $Entry.displayName $Selector
    }
    $selected = [ordered]@{}
    if ($AllWorkspaces) {
        # My workspace (type Personal) is left out here; Viewer-only workspaces are left out below.
        foreach ($entry in @($visibleWorkspaces | Where-Object { -not [string]::Equals([string](Get-OptionalProperty $_ 'type'), 'Personal', [StringComparison]::OrdinalIgnoreCase) } | Sort-Object displayName)) { $selected[$entry.id] = $entry }
    }
    $named = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)   # workspaces chosen with -Workspace
    foreach ($selector in $Workspace) {
        $found = @($visibleWorkspaces | Where-Object { Test-WorkspaceSelector $_ $selector })
        if ($found.Count -eq 0) { throw "No workspace matches '$selector'. Run with -ListWorkspaces to see the workspaces your account can see." }
        foreach ($entry in $found) { $selected[$entry.id] = $entry; [void]$named.Add([string]$entry.id) }
    }
    $excluded = 0
    if ($Exclude) {
        foreach ($id in @($selected.Keys)) { if (@($Exclude | Where-Object { Test-WorkspaceSelector $selected[$id] $_ }).Count) { $selected.Remove($id); $excluded++ } }
    }
    if (-not $selected.Count) { Write-UiLine "    $(Format-UiText Yellow $script:Ui.WarnMark)  No workspace left to process (see -Exclude)." }
    if ($PSCmdlet.ParameterSetName -eq 'ListModels') {
        $index = 0
        $listed = @(foreach ($entry in @($selected.Values)) {
                Show-Status 'Listing models' ($index++) $selected.Count $entry.displayName
                try { $items = @(Get-FabricItems "$script:FabricBase/workspaces/$($entry.id)/semanticModels") }
                catch {
                    if ($_.Exception -is [Security.SecurityException]) { throw }
                    Write-UiWarning "Could not list the semantic models of '$($entry.displayName)': $($_.Exception.Message)"
                    continue
                }
                foreach ($item in $items) {
                    if ($Model -and -not @($Model | Where-Object { Test-NameMatch $item.displayName $_ }).Count) { continue }
                    [pscustomobject]@{ Workspace = $entry.displayName; Model = $item.displayName; ModelId = $item.id; WorkspaceId = $entry.id }
                }
            })
        Stop-UiStatus
        Write-UiLine "    $(Format-UiText Dim "$(Format-UiCount $listed.Count 'semantic model') in $(Format-UiCount $selected.Count 'workspace')")"
        return $listed
    }
    # Belt and braces: the export below must only ever run with the output folder checked above.
    if (-not $exporting) { throw "Internal error: unexpected parameter set '$($PSCmdlet.ParameterSetName)'. Nothing was exported." }

    # Folder names over ALL visible workspaces, so that a workspace keeps its folder whatever the selection. Checked
    # against OneLake before anything else is asked: a workspace NAMED like the sync root must never get a folder.
    $workspaceFolderNames = Get-FolderNames $visibleWorkspaces 0
    foreach ($id in @($selected.Keys)) { Assert-NotOneLake (Join-Path $OutputPath $workspaceFolderNames[$id]) }

    # Viewer-only workspaces: Fabric gives definitions to Admins, Members and Contributors only. The workspace listing
    # filtered on those roles (same allowed request, with a query string) tells them apart before any definition is
    # asked; if that filter is not available, a Viewer-only workspace is recognised when its FIRST definition is refused.
    $viewerSkipped = [Collections.Generic.List[string]]::new()
    # (Not asked when every selected workspace was named with -Workspace: those are always tried.)
    if (@($selected.Keys | Where-Object { -not $named.Contains([string]$_) }).Count) {
        Show-Status 'Checking your workspace roles'
        $contributor = $null
        try { $contributor = [Collections.Generic.HashSet[string]]::new([string[]]@(Get-FabricItems "$script:FabricBase/workspaces?roles=Admin,Member,Contributor" | ForEach-Object { [string]$_.id }), [StringComparer]::OrdinalIgnoreCase) }
        catch {
            if ($_.Exception -is [Security.SecurityException]) { throw }
            Write-Verbose "Workspace roles not available ($($_.Exception.Message)): a Viewer-only workspace is recognised when Fabric refuses its first definition."
        }
        # An empty answer while workspaces are visible is not believed (a filter Fabric did not apply as expected): the
        # first-definition check below decides instead.
        if ($null -ne $contributor -and -not $contributor.Count -and $visibleWorkspaces.Count) {
            Write-Verbose 'The workspace role filter returned no workspace: not used, a Viewer-only workspace is recognised when Fabric refuses its first definition.'
            $contributor = $null
        }
        if ($null -ne $contributor) {
            # Only workspaces taken by -AllWorkspaces are left out on the filter alone. A workspace named with -Workspace is
            # always tried: if it is Viewer-only, its first refused definition gives the clear message.
            foreach ($id in @($selected.Keys)) {
                if ($named.Contains([string]$id)) { if (-not $contributor.Contains([string]$id)) { Write-Verbose "$($selected[$id].displayName): not in the role filter, tried anyway (named with -Workspace)." }; continue }
                if (-not $contributor.Contains([string]$id)) { $viewerSkipped.Add($selected[$id].displayName); $selected.Remove($id) }
            }
        }
    }
    $selection = @("$(Format-UiCount $selected.Count 'workspace') to export")
    if ($viewerSkipped.Count) { $selection += "$($viewerSkipped.Count) Viewer-only skipped" }
    if ($excluded) { $selection += "$excluded excluded" }
    Write-UiLine "    $(Format-UiText Dim ($selection -join " $($script:Ui.Dot) "))"
    Write-UiLine

    function Format-Counts($Rows) {
        $models = @($Rows | Where-Object { $_.ModelId })
        $done = @($models | Where-Object Status -EQ 'Exported').Count; $kept = @($models | Where-Object Status -EQ 'AlreadyExported').Count
        $failed = $models.Count - $done - $kept
        $parts = @("$done exported")
        if ($failed) { $parts += Format-UiText Red "$failed failed" }
        if ($kept) { $parts += "$kept already exported" }
        return $parts -join " $($script:Ui.Dot) "
    }
    function New-ResultRow($Entry, [string]$Model, [string]$ModelId, [string]$Folder) {
        return [pscustomobject][ordered]@{
            Workspace = $Entry.displayName; WorkspaceId = $Entry.id; Model = $Model; ModelId = $ModelId
            Status = 'Failed'; Files = 0; TmdlFiles = 0; Folder = $Folder; Error = ''; ExportedAtUtc = ''
        }
    }
    $results = [Collections.Generic.List[object]]::new()
    $listingFailures = [Collections.Generic.List[string]]::new()
    $workspaceIndex = 0; $workspaceTotal = $selected.Count
    foreach ($entry in @($selected.Values)) {
        $workspaceIndex++
        $label = if ($workspaceTotal -gt 1) { "($workspaceIndex/$workspaceTotal) $($entry.displayName)" } else { $entry.displayName }
        Show-Status $label 0 0 'listing models'
        $workspaceFolder = Join-Path $OutputPath $workspaceFolderNames[$entry.id]
        Assert-DirectChild $OutputPath $workspaceFolder
        Assert-NotOneLake $workspaceFolder
        Assert-SafeDirectory $workspaceFolder
        $workspaceShown = $false
        try { $models = @(Get-FabricItems "$script:FabricBase/workspaces/$($entry.id)/semanticModels") }
        catch {
            # One workspace that cannot be listed does not stop the others (a refusal of this script's guards does).
            if ($_.Exception -is [Security.SecurityException]) { throw }
            $row = New-ResultRow $entry '' '' $workspaceFolder
            $row.Error = "Could not list the semantic models: $($_.Exception.Message)"
            Write-UiLine "    $(Format-UiText Bold $entry.displayName)"
            Write-UiLine "      $(Format-UiText Red "$($script:Ui.FailMark)  $($row.Error)")"
            $listingFailures.Add($entry.displayName); $results.Add($row)
            continue
        }
        # The models selected by -Model: total of the counter, width of the name column.
        $wanted = @($models | Where-Object { $name = $_.displayName; -not $Model -or @($Model | Where-Object { Test-NameMatch $name $_ }).Count })
        $nameWidth = [math]::Min(40, [math]::Max(12, [int](@($wanted | ForEach-Object { ([string]$_.displayName).Length }) | Measure-Object -Maximum).Maximum))
        # Folder names over ALL models of the workspace (not only those of -Model): a model always gets the same folder.
        $modelFolderNames = Get-FolderNames $models 8
        $currentFolders = [Collections.Generic.HashSet[string]]::new([string[]]@($modelFolderNames.Values | ForEach-Object { "$_.SemanticModel" }), [StringComparer]::OrdinalIgnoreCase)
        $workspaceResults = [Collections.Generic.List[object]]::new()
        $modelIndex = 0; $keptCount = 0
        $reportPath = Join-Path $workspaceFolder '_export_report.csv'
        # Previous report: -NewOnly checks that a model folder belongs to the same model; -Model keeps the other models' lines.
        $previousReport = @()
        if (Test-Path -LiteralPath $reportPath -PathType Leaf) { try { $previousReport = @(Import-Csv -LiteralPath $reportPath) } catch { Write-Verbose "Unreadable $reportPath" } }
        $requested = $false; $viewerOnly = $false
        foreach ($item in $models) {
            if ($Model -and -not @($Model | Where-Object { Test-NameMatch $item.displayName $_ }).Count) { continue }
            $modelIndex++
            $folderName = "$($modelFolderNames[$item.id]).SemanticModel"
            $folder = Join-Path $workspaceFolder $folderName
            $result = New-ResultRow $entry $item.displayName $item.id $folder
            if ($NewOnly -and (Test-Path -LiteralPath (Join-Path $folder 'definition\model.tmdl') -PathType Leaf)) {
                # A model folder only ever holds a complete export (the previous one is replaced after full validation).
                # It counts as this model's export only if the previous report says so: same model id, same folder.
                $previous = @($previousReport | Where-Object {
                        [string]::Equals([string](Get-OptionalProperty $_ 'ModelId'), [string]$item.id, [StringComparison]::OrdinalIgnoreCase) -and
                        [string](Get-OptionalProperty $_ 'Status') -in 'Exported', 'AlreadyExported' -and
                        [string]::Equals([IO.Path]::GetFileName(([string](Get-OptionalProperty $_ 'Folder')).TrimEnd('\', '/')), $folderName, [StringComparison]::OrdinalIgnoreCase)
                    } | Select-Object -First 1)
                if ($previous.Count) {
                    $result.Status = 'AlreadyExported'
                    $result.Files = Get-OptionalProperty $previous[0] 'Files'; $result.TmdlFiles = Get-OptionalProperty $previous[0] 'TmdlFiles'; $result.ExportedAtUtc = Get-OptionalProperty $previous[0] 'ExportedAtUtc'
                    $workspaceResults.Add($result); $keptCount++
                    continue
                }
                Write-Verbose "$($item.displayName): the previous report does not tie '$folderName' to this model id, exported again."
            }
            Show-Status $label ($modelIndex - 1) $wanted.Count $item.displayName
            $watch = [Diagnostics.Stopwatch]::StartNew()
            try {
                Assert-DirectChild $workspaceFolder $folder
                Assert-NotOneLake $folder
                $first = -not $requested; $requested = $true
                try { $definition = Get-ModelDefinition $entry.id $item.id }
                catch {
                    # Viewer-only workspace (fallback when the role filter above was not available): Fabric refuses the
                    # definition to anyone below Contributor. Recognised on the FIRST definition request of the workspace only.
                    if ($first -and $_.Exception.Message -cmatch '\AHTTP 403:|InsufficientPrivileges|InsufficientPermissions') { $viewerOnly = $true }
                    throw
                }
                $files = [Collections.Generic.List[object]]::new()
                $prefix = [IO.Path]::GetFullPath($folder) + [IO.Path]::DirectorySeparatorChar
                $targets = @{}
                foreach ($part in $definition.parts) {
                    if ($part.payloadType -ne 'InlineBase64') { throw "Unsupported payload type: $($part.payloadType)" }
                    $relative = [string]$part.path
                    if (-not $relative -or [IO.Path]::IsPathRooted($relative) -or $relative.Contains(':') -or ($relative -split '[\\/]') -contains '..') { throw "Unsafe part path: $relative" }
                    # Windows silently drops trailing dots and spaces ('NUL ' is NUL, 'x.cmd.' is x.cmd): refuse such segments outright.
                    if (@(($relative -split '[\\/]') | Where-Object { $_ -ne $_.TrimEnd('.', ' ') -or -not $_.Trim() -or $_ -match $script:DeviceName }).Count) { throw "Unsafe part path (reserved device name, or a segment ending with a dot or a space): $relative" }
                    if (-not $script:PartExtensions.Contains([IO.Path]::GetExtension($relative))) { throw "Unexpected file type in the model definition, refused: $relative (allowed: $(@($script:PartExtensions) -join ' '))" }
                    $target = [IO.Path]::GetFullPath((Join-Path $folder $relative))
                    if (-not $target.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Part escapes model folder: $relative" }
                    if ($targets.ContainsKey($target)) { throw "Duplicate part path: $relative" }
                    $targets[$target] = $true
                    $files.Add([pscustomobject]@{ Path = $target; Bytes = [Convert]::FromBase64String($part.payload) })
                }
                Assert-SafeDirectory $folder
                # Download and validate all parts before deleting the previous export.
                if (Test-Path -LiteralPath $folder) { Remove-Item -LiteralPath $folder -Recurse -Force }
                [IO.Directory]::CreateDirectory($folder) | Out-Null
                foreach ($file in $files) {
                    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($file.Path)) | Out-Null
                    [IO.File]::WriteAllBytes($file.Path, $file.Bytes)
                }
                $result.Status = 'Exported'; $result.Files = $files.Count
                $result.TmdlFiles = @($files | Where-Object { [IO.Path]::GetExtension($_.Path) -ieq '.tmdl' }).Count
                $result.ExportedAtUtc = [datetime]::UtcNow.ToString('o')
                # The model had another folder before (a long name now cut, another collision order): that earlier export,
                # tied to the SAME model id by the previous report, would be analysed a second time. Deleted only now that the
                # new export is complete, and only a plain direct child of the workspace folder that no current model uses.
                foreach ($old in @($previousReport | Where-Object { [string]::Equals([string](Get-OptionalProperty $_ 'ModelId'), [string]$item.id, [StringComparison]::OrdinalIgnoreCase) })) {
                    $oldName = [IO.Path]::GetFileName(([string](Get-OptionalProperty $old 'Folder')).TrimEnd('\', '/'))
                    if (-not $oldName -or -not $oldName.EndsWith('.SemanticModel', [StringComparison]::OrdinalIgnoreCase) -or $currentFolders.Contains($oldName)) { continue }
                    $oldFolder = Join-Path $workspaceFolder $oldName
                    if (-not (Test-Path -LiteralPath $oldFolder -PathType Container)) { continue }
                    try {
                        Assert-DirectChild $workspaceFolder $oldFolder
                        Assert-NotOneLake $oldFolder
                        Assert-SafeDirectory $oldFolder
                        Remove-Item -LiteralPath $oldFolder -Recurse -Force
                        Write-Verbose "$($item.displayName): previous export in '$oldName' removed (now in '$folderName')."
                    } catch {
                        if ($_.Exception -is [Security.SecurityException]) { throw }
                        Write-UiWarning "$($item.displayName): its previous export '$oldFolder' was kept ($($_.Exception.Message)). Delete it by hand, or the analysis reads this model twice."
                    }
                }
                $duration = (Format-UiDuration $watch.Elapsed.TotalSeconds).PadLeft(7)
                $line = "$(Format-UiText Green $script:Ui.OkMark)  $(Format-UiName $item.displayName $nameWidth)  $(Format-UiCount $files.Count 'file' 4)  $(Format-UiText Dim $duration)"
            } catch {
                if ($_.Exception -is [Security.SecurityException]) { throw }
                if ($viewerOnly) { break }
                $result.Error = $_.Exception.Message
                $line = "$(Format-UiText Red $script:Ui.FailMark)  $(Format-UiName $item.displayName $nameWidth)  $(Format-UiText Red $result.Error)"
            }
            if (-not $workspaceShown) { Write-UiLine "    $(Format-UiText Bold $entry.displayName)"; $workspaceShown = $true }
            Write-UiLine "      $line"
            $workspaceResults.Add($result)
        }
        if ($viewerOnly) {
            # Nothing of this workspace is counted, written or returned, not even the models already exported (-NewOnly).
            $viewerSkipped.Add($entry.displayName)
            # Named with -Workspace: said now when other workspaces go on, else by the error at the end (said once).
            if ($Workspace -and $workspaceTotal -gt 1) { Write-UiWarning "Skipped: $(Get-ViewerMessage @($entry.displayName))" }
            else { Write-Verbose "$($entry.displayName): skipped, Fabric refuses the definitions (Viewer access only)." }
            continue
        }
        if ($keptCount -or -not $wanted.Count) {
            if (-not $workspaceShown) { Write-UiLine "    $(Format-UiText Bold $entry.displayName)"; $workspaceShown = $true }
            if ($keptCount) { Write-UiLine "      $(Format-UiText Dim "$($script:Ui.SkipMark)  $(Format-UiCount $keptCount 'model') already exported, kept (-NewOnly)")" }
            if (-not $wanted.Count) { Write-UiLine "      $(Format-UiText Dim "$($script:Ui.SkipMark)  no semantic model$(if ($Model) { ' matching -Model' })")" }
        }
        $results.AddRange($workspaceResults)
        [IO.Directory]::CreateDirectory($workspaceFolder) | Out-Null
        # -Model: the lines of the other models (still in the workspace) are kept from the previous report.
        $reportRows = @($workspaceResults)
        if ($Model -and $previousReport.Count) {
            $current = [Collections.Generic.HashSet[string]]::new([string[]]@($workspaceResults | ForEach-Object { [string]$_.ModelId }), [StringComparer]::OrdinalIgnoreCase)
            $existing = [Collections.Generic.HashSet[string]]::new([string[]]@($models | ForEach-Object { [string]$_.id }), [StringComparer]::OrdinalIgnoreCase)
            $others = @($previousReport | Where-Object { $id = [string](Get-OptionalProperty $_ 'ModelId'); $id -and -not $current.Contains($id) -and $existing.Contains($id) } | ForEach-Object {
                    $old = $_
                    $copy = [ordered]@{}
                    foreach ($column in 'Workspace', 'WorkspaceId', 'Model', 'ModelId', 'Status', 'Files', 'TmdlFiles', 'Folder', 'Error', 'ExportedAtUtc') { $copy[$column] = [string](Get-OptionalProperty $old $column) }
                    [pscustomobject]$copy
                })
            $reportRows = @(@($others) + @($workspaceResults))
        }
        # Drop the existing directory entry first, so a planted symbolic or hard link is never written through.
        $stale = Get-Item -LiteralPath $reportPath -Force -ErrorAction Ignore
        if ($stale) { $stale.Delete() }
        if ($reportRows.Count) { $reportRows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8 }
        else { '"Workspace","WorkspaceId","Model","ModelId","Status","Files","TmdlFiles","Folder","Error","ExportedAtUtc"' | Set-Content -LiteralPath $reportPath -Encoding utf8 }
    }
    Stop-UiStatus
    if ($Workspace -and $viewerSkipped.Count -and -not $results.Count) { throw (Get-ViewerMessage $viewerSkipped.ToArray()) }
    $failedModels = @($results | Where-Object { $_.ModelId -and $_.Status -eq 'Failed' }).Count
    $skipped = @()
    if ($viewerSkipped.Count) { $skipped += "$(Format-UiCount $viewerSkipped.Count 'Viewer-only workspace') skipped" }
    if ($listingFailures.Count) { $skipped += Format-UiText Red "$(Format-UiCount $listingFailures.Count 'workspace') not listed" }
    if ($script:Ui.Run) {
        # Inside the launcher: one line closes the step, the analysis prints the summary.
        Write-UiLine "    $(@(@(Format-Counts $results) + $skipped) -join " $($script:Ui.Dot) ")"
        Write-UiLine
    } else {
        $summary = [ordered]@{ Models = Format-Counts $results }
        if ($skipped.Count) { $summary.Workspaces = $skipped -join " $($script:Ui.Dot) " }
        $summary.Output = Format-UiPath $OutputPath
        if ($failedModels -or $listingFailures.Count) {
            $summary['To check'] = 'the errors above. A failed model keeps its previous export; -NewOnly retries only what is missing.'
        }
        if (@($results | Where-Object { $_.Status -in 'Exported', 'AlreadyExported' }).Count) {
            $parse = '.\Parse-SemanticModels.ps1'
            if (-not [string]::Equals($OutputPath.TrimEnd('\'), [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'data\export')).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) { $parse += " -ExportPath '$($OutputPath.Replace("'", "''"))'" }   # a quote in the path is doubled: the line can be pasted
            $summary.Next = "$parse   $(Format-UiText Dim 'lists what these models read')"
        }
        Write-UiSummary $summary
    }
    $results
} finally { Stop-UiStatus }

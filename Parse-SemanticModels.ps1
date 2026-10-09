#Requires -Version 7.2

<#
.SYNOPSIS
Reads exported semantic models (TMDL) and documents what every table reads: an Excel inventory plus one query file per table.
.DESCRIPTION
Run it with /? for the options and examples.
#>
# Works 100 % offline on the TMDL folders written by Extract-SemanticModels.ps1 (or any PBIP project). It never
# connects to Fabric, Power BI or any data source, and sends nothing anywhere. Windows only. Reads data\export and
# writes data\sources next to this script unless -ExportPath / -OutputPath say otherwise. Output, under -OutputPath:
#   model_sources.xlsx                          sheets Sources (one row per model, table, source object), Models, Schemas
#   model_sources.csv                           the Sources rows as text (also used to merge filtered runs)
#   <Workspace>\<Model>\power-query\<Table>.pq  M query of the table (.sql for legacy query partitions)
#   <Workspace>\<Model>\dax\<Table>.dax         calculated tables
#   <Workspace>\<Model>\sql\<Table>.sql         native SQL of the table, with parameters resolved
#   <Workspace>\<Model>\<Model>.sql             every table of the model: commented queries and resolved SQL
#   <Workspace>\<Model>\_shared\<Name>.pq       shared expressions and parameters used by the queries
# A filtered run (-Workspace, -Exclude, -Model, -NewOnly) keeps the rows of the other models already in the CSV.
# ScriptDom (Microsoft's T-SQL parser, from SSMS / DacFx / the SqlServer module, or -ScriptDomPath) is only loaded
# when it carries a valid Microsoft signature; without it native SQL goes to the built-in scanner.
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()][string]$ExportPath = (Join-Path $PSScriptRoot 'data\export'),
    [ValidateNotNullOrEmpty()][string]$OutputPath = (Join-Path $PSScriptRoot 'data\sources'),
    [string[]]$Workspace,
    [string[]]$Model,
    [switch]$IncludeAutoDateTables,
    [switch]$QueryFilesOnly,
    [ValidateNotNullOrEmpty()][string]$Delimiter = (Get-Culture).TextInfo.ListSeparator,
    [switch]$AdoptOutputFolder,
    [string]$ScriptDomPath,
    [switch]$NoTSqlParser,
    [switch]$PassThru,
    [switch]$AllWorkspaces,
    [ValidateNotNullOrEmpty()][string[]]$Exclude,
    [switch]$NewOnly,
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
$usage = [ordered]@{
    ''                                       = 'Analyses every model in data\export (next to this script) into data\sources: workbook, web page, CSV, query files. Overwrites what is there (see -NewOnly).'
    "-Workspace 'Finance' -Model 'Sales*'"   = 'Only these models. The other models already in the workbook are kept.'
    "-Exclude 'Sandbox*'"                    = 'Every exported workspace except Sandbox*.'
    '-NewOnly'                               = 'Only the models not in the workbook yet: the rows and query files already there are kept, not overwritten.'
    '-QueryFilesOnly'                        = 'Rewrites the query files only. The workbook and the CSV stay as they are.'
    "-ExportPath 'D:\Exports' -OutputPath 'D:\Inventory'" = 'Reads and writes other folders. The output folder must be new, empty, or already used by these scripts.'
    "-OutputPath 'D:\Old' -AdoptOutputFolder" = 'Accepts an output folder that already holds other files. Its model folders get replaced: be sure.'
    "-Delimiter ';'"                         = 'Separator of the CSV file. Default: the list separator of your regional settings.'
    '-IncludeAutoDateTables'                 = "Also lists the hidden date tables Power BI creates by itself (hidden by default)."
    '-NoTSqlParser'                          = "Reads SQL with the built-in scanner only, not with Microsoft's parser (ScriptDom, installed with SSMS)."
    "-ScriptDomPath 'C:\Tools\Microsoft.SqlServer.TransactSql.ScriptDom.dll'" = "Uses that copy of Microsoft's SQL parser (it must carry Microsoft's signature)."
    '-PassThru'                              = 'Also returns the rows as PowerShell objects, for your own scripts.'
}
# /?, -h, --help or help typed by habit: PowerShell binds them to the first positional parameter. Show the help instead.
if ($Help -or $ExportPath -in '/?', '?', '-h', '--help', 'help' -or @($Workspace | Where-Object { $_ -in '/?', '?', '-h', '--help', 'help' }).Count) {
    Write-UiHelp 'Parse-SemanticModels.ps1' 'Lists what every table of your exported models reads (SQL tables, views, files...) in an Excel workbook. Offline.' $usage
    return
}
if (-not $IsWindows) { throw 'This script runs on Windows only (it checks Windows paths, links and OneDrive / OneLake folders before writing). Run it in PowerShell 7 on Windows.' }
if ($Workspace -and $AllWorkspaces) { throw 'Use -Workspace or -AllWorkspaces, not both. Run with /? for help.' }

# ------------------------------------------------------------------ generic helpers
function Test-NameMatch([string]$Name, [string]$Pattern) {
    # Escape the backtick and the brackets only: * and ? retain their wildcard meaning (same as Extract-SemanticModels.ps1).
    return $Name -ilike $Pattern.Replace('`', '``').Replace('[', '`[').Replace(']', '`]')
}
function ConvertTo-FileName([string]$Name) {
    $safe = [regex]::Replace($Name, '[\\/:*?"<>|\p{Cc}]', '_').TrimEnd('.', ' ')
    if ([string]::IsNullOrEmpty($safe)) { return '_' }
    # reserved Windows device names, with or without extension (superscript digits included)
    if ($safe -match '\A(?i:CON|PRN|AUX|NUL|COM[0-9¹²³]|LPT[0-9¹²³])(\..*)?\z') { $safe = "_$safe" }
    return $safe
}
function Get-UniqueName([string]$Name, $Used) {
    $candidate = $Name; $n = 2
    while (-not $Used.Add($candidate)) { $candidate = "$Name ($n)"; $n++ }
    return $candidate
}
function Protect-CsvValue([string]$Value) {
    # Excel executes cells that start with = + - @ : neutralise them.
    if ($Value -match '^[=+\-@\t\r]') { return "'" + $Value }
    return $Value
}
function Export-SimpleXlsx([string]$Path, [object[]]$Sheets) {
    # Minimal OpenXML writer (an .xlsx is a zip of XML parts): text-only cells, bold frozen header,
    # auto-filter, column widths, rows to review highlighted. Needs neither Excel nor any external module.
    if (-not $Sheets.Count -or -not @($Sheets[0].Rows).Count) { return }
    # Hot loop: no scriptblock call per cell (column letters precomputed, escaping inlined).
    $invalid = [regex]::new('[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]')
    $escape = { param([string]$Text)
        if ($Text.Length -gt 32000) { $Text = $Text.Substring(0, 32000) + '...' }   # Excel cell limit: 32767 characters
        [Security.SecurityElement]::Escape($invalid.Replace($Text, '')) }
    $letter = { param([int]$Index) $name = ''; $k = $Index + 1; while ($k -gt 0) { $name = "$([char](65 + (($k - 1) % 26)))$name"; $k = [int][math]::Floor(($k - 1) / 26) }; $name }
    # Excel opens a sheet with at most 65,530 hyperlinks: beyond this cap the cells keep their text, without a link.
    $maxLinks = 65000
    $parts = [ordered]@{}; $sheetTags = ''; $filterTags = ''; $overrides = ''; $relationships = ''; $n = 0
    foreach ($definition in $Sheets) {
        $rows = @($definition.Rows); $sheetName = [string]$definition.Name
        $columns = @(if ($rows.Count) { $rows[0].PSObject.Properties.Name } else { $definition['Columns'] })
        if (-not $rows.Count -and -not $definition['Columns']) { continue }
        $highlightColumn = [string]$definition['HighlightColumn']; $linkColumns = @($definition['LinkColumns']); $wrapColumns = @($definition['WrapColumns'])
        $columnCount = $columns.Count
        $letters = [string[]]::new($columnCount); $isLinkColumn = [bool[]]::new($columnCount); $isWrapColumn = [bool[]]::new($columnCount)
        for ($c = 0; $c -lt $columnCount; $c++) { $letters[$c] = & $letter $c; $isLinkColumn[$c] = $columns[$c] -in $linkColumns; $isWrapColumn[$c] = $columns[$c] -in $wrapColumns }
        $n++; $links = [Collections.Generic.List[string]]::new(); $linkRels = [Collections.Generic.List[string]]::new(); $linkId = 0; $unlinked = 0
        $lastColumn = $letters[$columnCount - 1]; $lastRow = $rows.Count + 1
        # One pass over the cells: the column widths are measured while the rows are written, the <cols> element (which
        # comes first in the sheet) is put in front afterwards. The cell loop uses PowerShell operators only, no .NET method
        # call (PowerShell reports each method call to the antivirus, which costs microseconds per cell).
        $widths = [int[]]::new($columnCount); for ($c = 0; $c -lt $columnCount; $c++) { $widths[$c] = $columns[$c].Length }
        $header = for ($c = 0; $c -lt $columnCount; $c++) { "<c r=`"$($letters[$c])1`" t=`"inlineStr`" s=`"1`"><is><t xml:space=`"preserve`">$(& $escape $columns[$c])</t></is></c>" }
        $rowsXml = for ($r = 0; $r -lt $rows.Count; $r++) {
            $row = $rows[$r]; $rowNumber = $r + 2
            $style = if ($highlightColumn -and [string]$row.$highlightColumn) { 2 } else { 0 }
            $cells = for ($c = 0; $c -lt $columnCount; $c++) {
                $text = [string]$row.($columns[$c])
                if (-not $text -and -not $style -and -not $isWrapColumn[$c]) { continue }      # nothing to write
                $cellStyle = $style; $isLink = $false
                if ($text -and $isLinkColumn[$c]) {
                    $segments = $text -split '[\\/]'
                    if ($linkId -lt $maxLinks) {
                        # Encoded the way Excel itself writes a relative file link (checked against a workbook saved by Excel):
                        # backslashes kept, only %, [, ], space and # percent-encoded, everything else (@, (, ), accents...) raw.
                        $uri = ($segments -join '\') -replace '%', '%25' -replace '\[', '%5b' -replace '\]', '%5d' -replace ' ', '%20' -replace '#', '%23'
                        if ($uri -match '[<>&"'']') { $uri = $uri -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&apos;' }
                        $linkId++; $isLink = $true
                        $links.Add("<hyperlink ref=`"$($letters[$c])$rowNumber`" r:id=`"rIdH$linkId`"/>")
                        $linkRels.Add("<Relationship Id=`"rIdH$linkId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink`" Target=`"$uri`" TargetMode=`"External`"/>")
                        $cellStyle = if ($style) { 4 } else { 3 }
                    } else { $unlinked++ }
                    if ($segments.Count -ge 3) { $text = $segments[2..($segments.Count - 1)] -join '\' }
                }
                if ($isWrapColumn[$c]) {
                    $text = $text -replace "`r", ''; $cellStyle = if ($isLink) { if ($style) { 8 } else { 7 } } else { if ($style) { 6 } else { 5 } }
                    foreach ($line in $text -split "`n") { if ($line.Length -gt $widths[$c]) { $widths[$c] = $line.Length } }
                } elseif ($text.Length -gt $widths[$c]) { $widths[$c] = $text.Length }
                if ($text.Length -gt 32000) { $text = $text.Substring(0, 32000) + '...' }   # Excel cell limit: 32767 characters
                if ($text -match '[<>&"''\x00-\x08\x0B\x0C\x0E-\x1F￾￿]') { $text = $text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]', '' -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&apos;' }
                "<c r=`"$($letters[$c])$rowNumber`" t=`"inlineStr`" s=`"$cellStyle`"><is><t xml:space=`"preserve`">$text</t></is></c>"
            }
            "<row r=`"$rowNumber`">$($cells -join '')</row>"
        }
        if ($unlinked) { Write-UiWarning "Sheet $sheetName of the workbook: Excel accepts about 65,000 links per sheet, so $unlinked query file paths are shown as plain text, without a link (the web page has them all)." }
        $cols = for ($c = 0; $c -lt $columnCount; $c++) { "<col min=`"$($c + 1)`" max=`"$($c + 1)`" width=`"$([math]::Min(70, [math]::Max(10, $widths[$c] + 3)))`" customWidth=`"1`"/>" }
        $sheet = [Text.StringBuilder]::new()
        [void]$sheet.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$sheet.Append("<dimension ref=`"A1:$lastColumn$lastRow`"/><sheetViews><sheetView tabSelected=`"$([int]($n -eq 1))`" workbookViewId=`"0`"><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews><cols>")
        [void]$sheet.Append($cols -join '').Append('</cols><sheetData><row r="1">').Append($header -join '').Append('</row>').Append(@($rowsXml) -join '')
        $rowsXml = $null
        [void]$sheet.Append("</sheetData><autoFilter ref=`"A1:$lastColumn$lastRow`"/>")
        if ($linkId) {
            [void]$sheet.Append("<hyperlinks>$($links -join '')</hyperlinks>")
            $parts["xl/worksheets/_rels/sheet$n.xml.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ($linkRels -join '') + '</Relationships>'
        }
        [void]$sheet.Append('</worksheet>'); $parts["xl/worksheets/sheet$n.xml"] = $sheet.ToString()
        $sheetXmlName = & $escape $sheetName
        $filterRange = "'" + $sheetName.Replace("'", "''") + "'!" + '$A$1:$' + $lastColumn + '$' + $lastRow
        $sheetTags += "<sheet name=`"$sheetXmlName`" sheetId=`"$n`" r:id=`"rId$n`"/>"
        $filterTags += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$($n - 1)`" hidden=`"1`">$(& $escape $filterRange)</definedName>"
        $overrides += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $relationships += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
    }
    $baseParts = [ordered]@{
        '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        '_rels/.rels'                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/styles.xml'              = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/><xf numFmtId="49" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    }
    $baseParts['[Content_Types].xml'] = $baseParts['[Content_Types].xml'].Replace('</Types>', $overrides + '</Types>')
    $baseParts['xl/workbook.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $sheetTags + '</sheets><definedNames>' + $filterTags + '</definedNames></workbook>'
    $baseParts['xl/_rels/workbook.xml.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $relationships + '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    # Styles 0..2 stay stable; combinations preserve review fills on links and wrapped cells.
    $styles = $baseParts['xl/styles.xml'].Replace('<fonts count="2">', '<fonts count="3">').Replace('</fonts>', '<font><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><u/></font></fonts>').Replace('<cellXfs count="3">', '<cellXfs count="9">')
    $extraStyles = ''
    foreach ($s in @(@(2,0,0), @(2,3,0), @(0,0,1), @(0,3,1), @(2,0,1), @(2,3,1))) {
        $extraStyles += "<xf numFmtId=`"49`" fontId=`"$($s[0])`" fillId=`"$($s[1])`" borderId=`"0`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`" applyFill=`"1`" applyAlignment=`"1`">"
        if ($s[2]) { $extraStyles += '<alignment wrapText="1" vertical="top"/>' }; $extraStyles += '</xf>'
    }
    $baseParts['xl/styles.xml'] = $styles.Replace('</cellXfs>', $extraStyles + '</cellXfs>')
    foreach ($name in $baseParts.Keys) { $parts[$name] = $baseParts[$name] }
    # Written next to the target then moved over it: a workbook that is open in Excel fails cleanly (IOException).
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        $stream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew)
        try {
            $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($name in $parts.Keys) {
                    $writer = [IO.StreamWriter]::new($zip.CreateEntry($name).Open(), [Text.UTF8Encoding]::new($false))
                    try { $writer.Write([string]$parts[$name]) } finally { $writer.Dispose() }
                }
            } finally { $zip.Dispose() }
        } finally { $stream.Dispose() }
        [IO.File]::Move($temporary, $Path, $true)
    } finally { if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) } }
}

# ------------------------------------------------------------------ TMDL reading
$script:TmdlName = "(?:'(?:[^']|'')*'|[^\s=:']+)"
function ConvertFrom-TmdlName([string]$Name) {
    $Name = $Name.Trim()
    if ($Name.Length -ge 2 -and $Name.StartsWith("'") -and $Name.EndsWith("'")) { return $Name.Substring(1, $Name.Length - 2).Replace("''", "'") }
    return $Name
}
function Get-IndentUnit([string[]]$Lines) {
    # 0 = tab-indented (what Fabric emits); otherwise the number of spaces per level.
    foreach ($line in $Lines) { if ($line.StartsWith("`t")) { return 0 } }
    $unit = 0
    foreach ($line in $Lines) {
        $m = [regex]::Match($line, '^( +)\S')
        if ($m.Success -and ($unit -eq 0 -or $m.Groups[1].Length -lt $unit)) { $unit = $m.Groups[1].Length }
    }
    if ($unit -eq 0) { return 4 }
    return $unit
}
function Get-Level([string]$Line, [int]$Unit) {
    $n = 0
    if ($Unit -eq 0) { while ($n -lt $Line.Length -and $Line[$n] -eq "`t") { $n++ }; return $n }
    while ($n -lt $Line.Length -and $Line[$n] -eq ' ') { $n++ }
    return [int][math]::Floor($n / $Unit)
}
function Read-TmdlExpression([string[]]$Lines, [int]$Index, [string]$Rest, [int]$MinLevel, [int]$Unit) {
    # An expression is either on the declaring line, in a ``` fenced block, or in the following lines
    # indented at least $MinLevel. Returns the de-indented text and the index of the next line to read.
    $Rest = $Rest.Trim(); $next = $Index + 1
    $body = [Collections.Generic.List[string]]::new()
    if ($Rest.StartsWith('```')) {
        $after = $Rest.Substring(3).TrimEnd()
        if ($after.Length -ge 3 -and $after.EndsWith('```')) { $body.Add($after.Substring(0, $after.Length - 3)) }
        else {
            if ($after.Trim()) { $body.Add($after) }
            while ($next -lt $Lines.Count -and $Lines[$next].Trim() -ne '```') { $body.Add($Lines[$next]); $next++ }
            $next++
        }
    } elseif ($Rest) { $body.Add($Rest) }
    else {
        while ($next -lt $Lines.Count -and (-not $Lines[$next].Trim() -or (Get-Level $Lines[$next] $Unit) -ge $MinLevel)) { $body.Add($Lines[$next]); $next++ }
    }
    while ($body.Count -and -not $body[$body.Count - 1].Trim()) { $body.RemoveAt($body.Count - 1) }
    while ($body.Count -and -not $body[0].Trim()) { $body.RemoveAt(0) }
    $prefix = $null
    foreach ($line in $body) {
        if (-not $line.Trim()) { continue }
        $ws = $line.Substring(0, $line.Length - $line.TrimStart().Length)
        if ($null -eq $prefix) { $prefix = $ws; continue }
        $k = 0
        while ($k -lt $prefix.Length -and $k -lt $ws.Length -and $prefix[$k] -eq $ws[$k]) { $k++ }
        $prefix = $prefix.Substring(0, $k)
    }
    $cut = if ($null -eq $prefix) { 0 } else { $prefix.Length }
    $text = @($body | ForEach-Object { if ($_.Length -ge $cut) { $_.Substring($cut).TrimEnd() } else { $_.Trim() } }) -join "`n"
    return [pscustomobject]@{ Text = $text; Next = $next }
}
function Read-TmdlTableFile([string]$Path) {
    $lines = [IO.File]::ReadAllText($Path) -split "\r?\n"
    $unit = Get-IndentUnit $lines
    $table = [pscustomobject]@{ Name = [IO.Path]::GetFileNameWithoutExtension($Path); Partitions = [Collections.Generic.List[object]]::new(); PolicySource = '' }
    $named = $false; $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]; $level = Get-Level $line $unit
        if (-not $named -and $level -eq 0) {
            $m = [regex]::Match($line, "^table\s+($script:TmdlName)\s*$")
            if ($m.Success) { $table.Name = ConvertFrom-TmdlName $m.Groups[1].Value; $named = $true; $i++; continue }
        }
        if ($level -eq 1) {
            $m = [regex]::Match($line, "^\s*partition\s+($script:TmdlName)\s*=\s*(\w+)\s*$")
            if ($m.Success) {
                $p = [ordered]@{ Name = (ConvertFrom-TmdlName $m.Groups[1].Value); Type = $m.Groups[2].Value; Mode = ''; Source = ''
                    DataSource = ''; EntityName = ''; SchemaName = ''; ExpressionSource = '' }
                $i++
                while ($i -lt $lines.Count -and (-not $lines[$i].Trim() -or (Get-Level $lines[$i] $unit) -gt 1)) {
                    $pl = $lines[$i]
                    $pm = [regex]::Match($pl, '^\s*(?:source|query)\s*=(.*)$')
                    if ($pm.Success -and (Get-Level $pl $unit) -eq 2) {
                        $e = Read-TmdlExpression $lines $i $pm.Groups[1].Value 3 $unit
                        $p.Source = $e.Text; $i = $e.Next; continue
                    }
                    $pm = [regex]::Match($pl, '^\s*(mode|dataSource|entityName|schemaName|expressionSource)\s*:\s*(.*?)\s*$')
                    if ($pm.Success) { $p[$pm.Groups[1].Value] = ConvertFrom-TmdlName $pm.Groups[2].Value }
                    $i++
                }
                $table.Partitions.Add([pscustomobject]$p); continue
            }
        }
        if ($level -eq 2) {
            $m = [regex]::Match($line, '^\s*sourceExpression\s*=(.*)$')
            if ($m.Success) { $e = Read-TmdlExpression $lines $i $m.Groups[1].Value 3 $unit; $table.PolicySource = $e.Text; $i = $e.Next; continue }
        }
        # Skip ``` fenced expressions of measures / columns so their content is never read as structure.
        if ([regex]::Matches($line, '```').Count -eq 1 -and $line.TrimEnd().EndsWith('```') -and $line.Trim() -ne '```') {
            $i++
            while ($i -lt $lines.Count -and $lines[$i].Trim() -ne '```') { $i++ }
        }
        $i++
    }
    return $table
}
function Read-TmdlExpressionsFile([string]$Path, $Into) {
    $lines = [IO.File]::ReadAllText($Path) -split "\r?\n"
    $unit = Get-IndentUnit $lines
    $i = 0
    while ($i -lt $lines.Count) {
        $m = [regex]::Match($lines[$i], "^expression\s+($script:TmdlName)\s*=(.*)$")
        if ($m.Success) {
            $e = Read-TmdlExpression $lines $i $m.Groups[2].Value 2 $unit
            $Into[(ConvertFrom-TmdlName $m.Groups[1].Value)] = $e.Text
            $i = $e.Next; continue
        }
        $i++
    }
}
function Read-TmdlDataSources([string]$Path) {
    # Legacy (provider) data sources of a model, used by 'query' partitions: name -> server, database and provider read
    # from the connection string. Nothing else is kept (the connection string may hold a password).
    $result = @{}
    if (-not [IO.File]::Exists($Path)) { return $result }
    $lines = [IO.File]::ReadAllText($Path) -split "\r?\n"
    $unit = Get-IndentUnit $lines; $current = $null
    foreach ($line in $lines) {
        if (-not $line.Trim()) { continue }
        $level = Get-Level $line $unit
        if ($level -eq 0) {
            $m = [regex]::Match($line, "^dataSource\s+($script:TmdlName)\s*$")
            $current = $null
            if ($m.Success) { $current = @{ Server = ''; Database = ''; Provider = '' }; $result[(ConvertFrom-TmdlName $m.Groups[1].Value)] = $current }
            continue
        }
        if ($null -eq $current -or $level -ne 1) { continue }
        $m = [regex]::Match($line, '^\s*(connectionString|provider)\s*:\s*(.*?)\s*$')
        if (-not $m.Success) { continue }
        $value = $m.Groups[2].Value
        if ($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) { $value = $value.Substring(1, $value.Length - 2).Replace('""', '"') }
        if ($m.Groups[1].Value -ceq 'provider') { $current.Provider = $value; continue }
        foreach ($pair in $value -split ';') {
            $kv = $pair -split '=', 2
            if ($kv.Count -ne 2) { continue }
            $key = $kv[0].Trim(); $text = $kv[1].Trim().Trim('"', "'")
            if ($key -match '\A(?i:data source|server|address|addr|network address|host)\z') { if (-not $current.Server) { $current.Server = $text } }
            elseif ($key -match '\A(?i:initial catalog|database)\z') { if (-not $current.Database) { $current.Database = $text } }
            elseif ($key -match '\A(?i:provider)\z') { if (-not $current.Provider) { $current.Provider = $text } }
        }
    }
    return $result
}

# ------------------------------------------------------------------ M analysis
$script:CallRx = [regex]::new('(?<![\w\.#"])([\p{L}_]\w*(?:\.[\p{L}_]\w*)+)\s*\(')
$script:NavRx = [regex]::new('\{\s*\[')
$script:IdentRx = [regex]::new('#"((?:[^"]|"")*)"|(?<![\w\.#"])([\p{L}_][\w\.]*)')

# Connector functions: K = kind, S / D / O / Q = index of the server, database, object and SQL-text arguments.
$script:Connectors = @{}
foreach ($n in 'Sql.Database', 'PostgreSQL.Database', 'MySQL.Database', 'DB2.Database', 'Sybase.Database', 'Informix.Database', 'AmazonRedshift.Database', 'AnalysisServices.Database') { $script:Connectors[$n] = @{ K = 'db'; S = 0; D = 1 } }
foreach ($n in 'Sql.Databases', 'Oracle.Database', 'Teradata.Database', 'SapHana.Database', 'Snowflake.Databases', 'Databricks.Catalogs', 'Databricks.Contents', 'AnalysisServices.Databases', 'CommonDataService.Database', 'AzureStorage.Tables', 'ActiveDirectory.Domains', 'SapBusinessWarehouse.Cubes', 'Impala.Database', 'Vertica.Database') { $script:Connectors[$n] = @{ K = 'db'; S = 0 } }
foreach ($n in 'GoogleBigQuery.Database', 'Lakehouse.Contents', 'Fabric.Warehouse', 'PowerBI.Dataflows', 'PowerPlatform.Dataflows', 'PowerBI.Datamarts', 'Salesforce.Data', 'Salesforce.Reports', 'GoogleAnalytics.Accounts', 'Exchange.Contents', 'Dataverse.Contents') { $script:Connectors[$n] = @{ K = 'db' } }
foreach ($n in 'AzureDataExplorer.Contents', 'Kusto.Contents') { $script:Connectors[$n] = @{ K = 'db'; S = 0; D = 1; O = 2 } }
foreach ($n in 'Odbc.DataSource', 'OleDb.DataSource') { $script:Connectors[$n] = @{ K = 'db'; S = 0; ConnectionString = $true } }
foreach ($n in 'Odbc.Query', 'OleDb.Query') { $script:Connectors[$n] = @{ K = 'db'; S = 0; Q = 1; ConnectionString = $true } }
$script:Connectors['File.Contents'] = @{ K = 'file'; S = 0 }
foreach ($n in 'Folder.Files', 'Folder.Contents', 'SharePoint.Files', 'SharePoint.Contents', 'SharePoint.Tables', 'AzureStorage.Blobs', 'AzureStorage.DataLake', 'Hdfs.Files', 'Hdfs.Contents', 'OData.Feed') { $script:Connectors[$n] = @{ K = 'container'; S = 0 } }
foreach ($n in 'Web.Contents', 'Web.BrowserContents', 'Web.Headers', 'Soda.Feed', 'AzureStorage.BlobContents', 'AzureStorage.DataLakeContents', 'UsageMetricsDataConnector.GetMetricsData') { $script:Connectors[$n] = @{ K = 'url'; S = 0 } }
# Power BI / Adobe navigation-style connectors: no location argument, the object comes from the navigation steps.
foreach ($n in 'AdobeAnalytics.Cubes', 'PowerBIRESTAPI.Navigation', 'pbiAdminAPI.Navigation', 'PowerPlatform.Dataflows', 'PowerBI.Dataflows') { $script:Connectors[$n] = @{ K = 'db' } }
# Connectors whose first navigation level by Name is already the object (Salesforce.Data(){[Name="Account"]}), not a database.
foreach ($n in 'Salesforce.Data', 'Salesforce.Reports', 'Exchange.Contents') { $script:Connectors[$n] = @{ K = 'db'; NoDatabase = $true } }
# Common Data Model folders: Cdm.Contents(AzureStorage.DataLake(...)) - the location is the call inside it.
$script:Connectors['Cdm.Contents'] = @{ K = 'db'; Wrapper = $true }
# M code evaluated at refresh time: what it reads cannot be known from here.
$script:Connectors['Expression.Evaluate'] = @{ K = 'db'; Dynamic = $true }
# The M standard library: calls that read nothing outside the query. Whole namespaces where every function only transforms
# values; single functions where a namespace also holds data access (Web.Contents, Sql.Database, Odbc.*, AzureStorage.*...)
# or dynamic evaluation (Expression.Evaluate). Any other call X.Y( is reported as an unknown connector, to review.
$script:StdLib = [Collections.Generic.HashSet[string]]::new([string[]]('Table List Text Record Date DateTime DateTimeZone Time Duration Number Value Splitter Combiner Json Binary BinaryFormat BinaryEncoding Csv Excel Lines Function Type Logical Character Comparer Replacer Error Diagnostics Cube Xml Html Pdf Parquet DeltaLake Uri Int8 Int16 Int32 Int64 Double Decimal Single Byte Currency Percentage Guid Geography Geometry GeographyPoint GeometryPoint SqlExpression ItemExpression RowExpression Delimiter Culture Occurrence Order JoinKind JoinAlgorithm JoinSide MissingField ExtraValues QuoteStyle RoundingMode TextEncoding Compression GroupKind Precision RelativePosition TraceLevel Day LimitClauseKind ODataOmitValues AccessControlKind PercentileMode BufferMode WebMethod' -split ' '), [StringComparer]::Ordinal)
$script:StdFunctions = [Collections.Generic.HashSet[string]]::new([string[]]('Expression.Constant Expression.Identifier Web.Page Access.Database RData.FromBinary Variable.Value Variable.ValueOrDefault Embedded.Value' -split ' '), [StringComparer]::Ordinal)

# The structural scans below never look at every character: one regular expression lists the brackets and separators of
# a text once (cached), and the scans walk that list. A large query (25 KB) is read in milliseconds instead of seconds.
# Inside these loops: PowerShell operators and array indexing rather than .NET method calls (PowerShell reports each
# method call to the antivirus, which costs microseconds each time).
$script:LetScanRx = [regex]::new('[()\[\]{},]|(?<![\w.#"])in(?![\w.])')
$script:ReferenceScanRx = [regex]::new('[()\[\]{},=]')
# The M keywords (language specification), case-sensitive like M.
$script:MKeywords = [Collections.Generic.HashSet[string]]::new([string[]]('and as each else error false if in is let meta not null or otherwise section shared then true try type' -split ' '), [StringComparer]::Ordinal)
# Every piece of an M text: a string, a comment, a run of other characters, or one lone character (an unterminated " or /).
$script:PieceRx = [regex]::new('"(?:[^"]|"")*"|//[^\n]*|/\*[\s\S]*?\*/|[^"/]+|[\s\S]')
$script:MarkRx = [regex]::new('[()\[\]{},&]')
# Masks and marks of the texts already seen during the current model (the same expressions are read again and again).
$script:MaskCache = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
$script:MarkCache = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
function Get-MMasks([string]$Text) {
    # Clean = comments blanked. Mask = comments and string contents blanked. Same length, so offsets line up.
    $result = $script:MaskCache[$Text]
    if ($null -ne $result) { return $result }
    if ($Text -notmatch '["/]') { $result = [pscustomobject]@{ Clean = $Text; Mask = $Text } }    # neither string nor comment
    else {
        $pieces = foreach ($piece in $script:PieceRx.Matches($Text)) { $piece.Value }
        $cleanParts = foreach ($value in $pieces) {
            if ($value.Length -ge 2 -and $value[0] -ceq '/' -and ($value[1] -ceq '/' -or $value[1] -ceq '*')) { $value -replace '[^\n]', ' ' } else { $value }
        }
        $maskParts = foreach ($value in $pieces) {
            if ($value.Length -ge 2 -and $value[0] -ceq '"') { '"' + (' ' * ($value.Length - 2)) + '"' }
            elseif ($value.Length -ge 2 -and $value[0] -ceq '/' -and ($value[1] -ceq '/' -or $value[1] -ceq '*')) { $value -replace '[^\n]', ' ' }
            else { $value }
        }
        $result = [pscustomobject]@{ Clean = $cleanParts -join ''; Mask = $maskParts -join '' }
    }
    if ($script:MaskCache.Count -ge 50000) { $script:MaskCache.Clear() }
    $script:MaskCache[$Text] = $result
    return $result
}
function Get-MaskMarks([string]$Mask) {
    # Positions of the brackets and separators of a mask, in order, with their kind: 1 = opening, 2 = closing, 0 = separator.
    $marks = $script:MarkCache[$Mask]
    if ($null -ne $marks) { return $marks }
    $values = foreach ($m in $script:MarkRx.Matches($Mask)) { $m.Index; $m.Value }
    $count = [int]($(if ($null -eq $values) { 0 } else { @($values).Count }) / 2)
    $positions = [int[]]::new($count); $chars = [string[]]::new($count); $kinds = [int[]]::new($count)
    for ($k = 0; $k -lt $count; $k++) {
        $positions[$k] = $values[2 * $k]; $ch = $values[2 * $k + 1]; $chars[$k] = $ch
        $kinds[$k] = if ($ch -ceq '(' -or $ch -ceq '[' -or $ch -ceq '{') { 1 } elseif ($ch -ceq ')' -or $ch -ceq ']' -or $ch -ceq '}') { 2 } else { 0 }
    }
    $marks = [pscustomobject]@{ Positions = $positions; Chars = $chars; Kinds = $kinds }
    if ($script:MarkCache.Count -ge 50000) { $script:MarkCache.Clear() }
    $script:MarkCache[$Mask] = $marks
    return $marks
}
function Find-Close([string]$Mask, [int]$Open) {
    # Index of the bracket that closes the one at $Open (any kind: (, [ and { all count), or -1.
    $marks = Get-MaskMarks $Mask; $positions = $marks.Positions; $kinds = $marks.Kinds
    $k = [Array]::BinarySearch($positions, $Open)
    if ($k -lt 0) { $k = -bnot $k }      # (not on a mark: scan from the next one)
    $depth = 0
    for (; $k -lt $positions.Length; $k++) {
        $kind = $kinds[$k]
        if ($kind -eq 1) { $depth++ } elseif ($kind -eq 2) { $depth--; if ($depth -eq 0) { return $positions[$k] } }
    }
    return -1
}
function Split-TopLevel([string]$Clean, [string]$Mask, [char]$Separator) {
    $marks = Get-MaskMarks $Mask; $positions = $marks.Positions; $kinds = $marks.Kinds; $chars = $marks.Chars
    $separatorText = [string]$Separator; $depth = 0; $start = 0
    $parts = @(for ($k = 0; $k -lt $positions.Length; $k++) {
            $kind = $kinds[$k]
            if ($kind -eq 1) { $depth++ } elseif ($kind -eq 2) { $depth-- }
            elseif ($depth -eq 0 -and $chars[$k] -ceq $separatorText) { $Clean.Substring($start, $positions[$k] - $start); $start = $positions[$k] + 1 }
        })
    return [string[]]($parts + $Clean.Substring($start))
}
function ConvertFrom-MString([string]$Inner) {
    $text = $Inner.Replace('""', '"')
    if (-not $text.Contains('#(')) { return $text }      # no escape sequence: the common case, kept fast
    return [regex]::Replace($text, '#\(([^)]*)\)', { param($m)
            (@($m.Groups[1].Value -split ',' | ForEach-Object {
                    switch ($_.Trim()) { 'lf' { "`n" } 'cr' { "`r" } 'tab' { "`t" } '#' { '#' }
                        default { if ($_ -match '^\s*[0-9A-Fa-f]{4,8}\s*$') { [char]::ConvertFromUtf32([Convert]::ToInt32($_.Trim(), 16)) } else { $m.Value } } }
                })) -join '' })
}
function New-MScope($Literals, $Let) {
    # Literals = model parameters (name -> text). Bindings = steps of the query's let (name -> expression).
    return [pscustomobject]@{ Literals = $Literals; Bindings = $Let.Map; Ambiguous = $Let.Ambiguous; Active = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal) }
}
function ConvertTo-Placeholder([string]$Text, [bool]$Scalar) {
    # What could not be evaluated. {…} stands for ONE scalar value (a date, a number: it cannot hide SQL structure);
    # {?…} is opaque text that may contain anything, tables included - results built on it are never reported as certain.
    # Braces and quotes are removed so that a placeholder stays one recognisable token inside SQL text.
    $short = [regex]::Replace(($Text -replace "[{}'?]", ''), '\s+', ' ').Trim()
    if ($short.Length -gt 40) { $short = $short.Substring(0, 37) + '...' }
    return $(if ($Scalar) { "{$short}" } else { "{?$short}" })
}
function Get-LetBindings([string]$Clean, [string]$Mask) {
    # Steps of the outermost let, name -> expression text, so that a SQL text assembled in a variable
    # (sqlQuery = "select ..." & x & "...") can be followed from [Query = sqlQuery].
    # Returns Map (name -> expression), Steps (name + position of each step, to know which step an event belongs to)
    # and Ambiguous (names defined more than once, e.g. in a nested let: their value depends on a scope that is
    # not modelled here, so they are never resolved).
    $bindings = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $result = [pscustomobject]@{ Map = $bindings; Steps = [Collections.Generic.List[object]]::new(); Ambiguous = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal) }
    $let = [regex]::Match($Mask, '(?<![\w\.#"])let(?![\w\.])')
    if (-not $let.Success) { return $result }
    $depth = 0
    $marks = Get-MaskMarks $Mask
    for ($k = 0; $k -lt $marks.Positions.Length -and $marks.Positions[$k] -lt $let.Index; $k++) { $kind = $marks.Kinds[$k]; if ($kind -eq 1) { $depth++ } elseif ($kind -eq 2) { $depth-- } }
    if ($depth -ne 0) { return $result }
    $pieces = [Collections.Generic.List[int[]]]::new(); $start = $let.Index + $let.Length; $lastIn = -1; $n = $Mask.Length
    # Brackets, commas and the keyword "in" (not part of a longer name), in order.
    foreach ($m in $script:LetScanRx.Matches($Mask, $start)) {
        $ch = $m.Value; $i = $m.Index
        if ($ch -ceq '(' -or $ch -ceq '[' -or $ch -ceq '{') { $depth++; continue }
        if ($ch -ceq ')' -or $ch -ceq ']' -or $ch -ceq '}') { $depth--; continue }
        if ($depth -ne 0) { continue }
        if ($ch -ceq ',') { $pieces.Add([int[]]($start, $i)); $start = $i + 1; $lastIn = -1; continue }
        $lastIn = $i
    }
    $pieces.Add([int[]]($start, $(if ($lastIn -gt $start) { $lastIn } else { $n })))
    foreach ($piece in $pieces) {
        $text = $Clean.Substring($piece[0], $piece[1] - $piece[0])
        # POS = let sqlQuery = "..." : the steps of a let nested directly in a step are read as steps too.
        while ($true) { $nested = [regex]::Match($text, '\A\s*(?:#"(?:[^"]|"")*"|[\p{L}_][\w\.]*)\s*=\s*let(?![\w\.])([\s\S]*)\z'); if (-not $nested.Success) { break }; $text = $nested.Groups[1].Value }
        $b = [regex]::Match($text, '\A\s*(?:#"((?:[^"]|"")*)"|([\p{L}_][\w\.]*))\s*=(?!>)([\s\S]*)\z')
        if (-not $b.Success) { continue }
        $name = if ($b.Groups[1].Success) { $b.Groups[1].Value.Replace('""', '"') } else { $b.Groups[2].Value }
        if ($bindings.ContainsKey($name)) { [void]$result.Ambiguous.Add($name) } else { $bindings[$name] = $b.Groups[3].Value.Trim() }
        $result.Steps.Add([pscustomobject]@{ Name = $name; Start = $piece[0]; End = $piece[1] })
    }
    return $result
}
$script:ModelParameters = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)   # names of the current model's parameters, of any type
function Resolve-MText([string]$Expression, $Scope, [int]$Depth = 0) {
    # Evaluates a text expression made of string literals, let steps, model parameters and & . What cannot be evaluated
    # becomes a placeholder: scalar {…} for a single value (a date / number conversion, a non-text parameter such as
    # RangeStart), opaque {?…} for anything else (Text.Combine, if … then … else, an unknown name, an ambiguous step).
    # Resolved = no placeholder at all. Opaque = at least one opaque placeholder: the text may hide SQL structure.
    # Only model parameters are reported in Params (a let step holding the SQL text is not one).
    # Placeholders = the placeholder texts inserted, so that later steps tell them apart from braces of the source
    # text itself ({d '2024-01-01'} ODBC escapes, driver={...}, {[Measures]} MDX sets).
    $result = [pscustomobject]@{ Text = ''; Resolved = $true; Opaque = $false; Params = [Collections.Generic.List[string]]::new(); Placeholders = [Collections.Generic.List[string]]::new() }
    $absorb = { param($inner) $result.Text += $inner.Text; $result.Params.AddRange($inner.Params); $result.Placeholders.AddRange($inner.Placeholders); if (-not $inner.Resolved) { $result.Resolved = $false }; if ($inner.Opaque) { $result.Opaque = $true } }
    $unknown = { param([string]$text, [bool]$scalar) $placeholder = ConvertTo-Placeholder $text $scalar; $result.Text += $placeholder; $result.Placeholders.Add($placeholder); $result.Resolved = $false; if (-not $scalar) { $result.Opaque = $true } }
    $masks = Get-MMasks $Expression.Trim()
    foreach ($term in @(Split-TopLevel $masks.Clean $masks.Mask '&')) {
        $t = $term.Trim()
        $m = [regex]::Match($t, '\A"((?:[^"]|"")*)"\z', 'Singleline')
        if ($m.Success) { $result.Text += ConvertFrom-MString $m.Groups[1].Value; continue }
        $m = [regex]::Match($t, '\A(?:#"((?:[^"]|"")*)"|([\p{L}_][\w\.]*))\z')
        if ($m.Success) {
            $name = if ($m.Groups[1].Success) { $m.Groups[1].Value.Replace('""', '"') } else { $m.Groups[2].Value }
            if ($name -ceq 'null') { continue }
            if ($Scope.Ambiguous.Contains($name)) { & $unknown $name $false; continue }
            if ($Scope.Bindings.ContainsKey($name) -and $Depth -lt 15 -and $Scope.Active.Add($name)) {   # a local step shadows a parameter
                $inner = Resolve-MText $Scope.Bindings[$name] $Scope ($Depth + 1)
                [void]$Scope.Active.Remove($name)
                & $absorb $inner; continue
            }
            if ($Scope.Literals.ContainsKey($name)) {
                $value = $Scope.Literals[$name]; $result.Text += $value
                $result.Params.Add("$name=$(if ($value.Length -gt 100) { $value.Substring(0, 97) + '...' } else { $value })")
            } else { & $unknown $name $script:ModelParameters.Contains($name) }
            continue
        }
        if ($t -match '\A-?\d+(\.\d+)?\z') { $result.Text += $t; continue }
        # ( … ) or Function( … )
        $m = [regex]::Match($t, '\A((?:[\p{L}_]\w*\.)*[\p{L}_]\w*)?\s*\(')
        if ($m.Success -and (Find-Close (Get-MMasks $t).Mask ($m.Length - 1)) -eq $t.Length - 1) {
            $function = $m.Groups[1].Value
            $inside = $t.Substring($m.Length, $t.Length - $m.Length - 1); $insideMasks = Get-MMasks $inside
            $arguments = @(Split-TopLevel $insideMasks.Clean $insideMasks.Mask ',')
            if ($function -in '', 'Text.From', 'Number.ToText', 'Text.Trim' -and $arguments.Count -eq 1) {
                $inner = Resolve-MText $arguments[0] $Scope ($Depth + 1)
                if ($function -ceq 'Text.Trim') { $inner.Text = $inner.Text.Trim() }
                & $absorb $inner; continue
            }
            # conversions of dates and numbers give ONE value; any other function may return arbitrary text
            & $unknown $t ($function -cmatch '\A(?:DateTime|Date|DateTimeZone|Time|Duration|Number|Int64|Int32|Decimal|Double|Currency|Percentage|Logical)\.')
            continue
        }
        & $unknown $t $false
    }
    return $result
}
function Get-MRecordFields([string]$Clean, [string]$Mask) {
    # $Clean / $Mask = the text between [ and ]. Returns name -> raw value expression (names are case-insensitive).
    $fields = [ordered]@{}
    $pieces = @(Split-TopLevel $Clean $Mask ',')
    foreach ($piece in $pieces) {
        $m = [regex]::Match($piece, '\A\s*(?:#"((?:[^"]|"")*)"|([\p{L}_][\w\.]*))\s*=(?!>)([\s\S]*)\z')
        if (-not $m.Success) { continue }
        $name = if ($m.Groups[1].Success) { $m.Groups[1].Value.Replace('""', '"') } else { $m.Groups[2].Value }
        $fields[$name] = $m.Groups[3].Value.Trim()
    }
    return $fields
}
# ------------------------------------------------------------------ SQL analysis
# Two analysers. 1) Microsoft's own T-SQL grammar (ScriptDom, the parser behind SSMS / DacFx) when its DLL is on the
# machine: a real parser, so CTEs, sub-queries, comments, strings, hints, APPLY, comma joins... are handled by
# construction. 2) A built-in lexer + scanner for other SQL dialects, for text that is not valid T-SQL, and for
# machines without ScriptDom. Both name their rows the same way, so the inventory does not say which one ran.
$script:TSql = $null
# Object types of SQL objects: a database (or server) prefix in front of such a name goes to its own column.
$script:SqlObjectTypes = @('Table/View', 'Stored procedure', 'Function')
function Format-SqlName([string[]]$Parts) {
    # dbo.Sales, [Other Db]..Country, stg.[Weird.Name] : brackets only where they are needed, so that
    # [dbo].[Sales] and dbo.Sales are the same object.
    return (@(foreach ($part in $Parts) { if ($part -eq '' -or $part -match '\A[\w@#$]+\z' -or $part -match '\A\{[^{}]*\}\z') { $part } else { '[' + $part.Replace(']', ']]') + ']' } }) -join '.')
}
function Split-SqlName([string]$Name) {
    # The inverse of Format-SqlName: 'db.[Weird.Name]' -> 'db', 'Weird.Name'. Dots inside [brackets] and inside
    # {placeholders} do not separate parts; an empty part ('db..T') is kept so that the position of each part holds.
    if ($Name.IndexOfAny([char[]]@('[', '{')) -lt 0) { return $Name.Split('.') }      # plain name: the common case
    $parts = [Collections.Generic.List[string]]::new(); $current = [Text.StringBuilder]::new(); $i = 0
    while ($i -lt $Name.Length) {
        $ch = $Name[$i]
        if ($ch -eq '[') {
            $end = $i + 1
            while ($end -lt $Name.Length) { if ($Name[$end] -eq ']') { if ($end + 1 -lt $Name.Length -and $Name[$end + 1] -eq ']') { $end += 2; continue }; break }; $end++ }
            if ($end -ge $Name.Length) { [void]$current.Append($Name.Substring($i)); break }
            [void]$current.Append($Name.Substring($i + 1, $end - $i - 1).Replace(']]', ']')); $i = $end + 1; continue
        }
        if ($ch -eq '{') { $end = $Name.IndexOf('}', $i); if ($end -lt 0) { $end = $Name.Length - 1 }; [void]$current.Append($Name.Substring($i, $end - $i + 1)); $i = $end + 1; continue }
        if ($ch -eq '.') { $parts.Add($current.ToString()); [void]$current.Clear(); $i++; continue }
        [void]$current.Append($ch); $i++
    }
    $parts.Add($current.ToString())
    return $parts.ToArray()
}
function Get-ObjectSchema([string]$Object, [string]$ObjectType) {
    if (-not $Object -or ($ObjectType -notin $script:SqlObjectTypes -and $ObjectType -ne 'Table')) { return '' }
    $parts = @(Split-SqlName $Object)
    if ($parts.Count -ge 2) { return $parts[0] }
    return '(default)'
}
function Get-InventorySheets([object[]]$Rows) {
    # Grouped with plain string keys (each part prefixed with its length, so that no name can collide with another),
    # case-insensitively like the rest of the inventory. This loop runs once per row: it uses PowerShell operators and
    # dictionary indexing only, no .NET method call (each one costs microseconds more, as PowerShell reports every
    # method call to the antivirus).
    $models = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $schemas = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $schemaOf = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)     # Get-ObjectSchema, once per object
    foreach ($row in $Rows) {
        $workspaceText = [string]$row.Workspace; $modelText = [string]$row.SemanticModel; $serverText = [string]$row.Server; $databaseText = [string]$row.Database
        $objectText = [string]$row.Object; $tableText = [string]$row.Table; $review = $row.NeedsReview -eq 'yes'
        $key = "$($workspaceText.Length):$workspaceText$($modelText.Length):$modelText"
        $group = $models[$key]
        if ($null -eq $group) {
            $group = @{ First = $row; Count = 0; Review = 0; Tables = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal)
                Sources = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal); Schemas = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal) }
            $models[$key] = $group
        }
        $group.Count++; $group.Tables[$tableText] = 1
        if ($review) { $group.Review++ }
        $source = if ($serverText -and $databaseText) { "$serverText / $databaseText" } elseif ($serverText -or $databaseText) { "$serverText$databaseText" } elseif ($objectText) { [string]$row.SourceType } else { '' }
        if ($source) { $group.Sources[$source] = 1 }      # insertion order is kept: first seen first
        if (-not ($serverText -or $databaseText -or $objectText)) { continue }
        $objectKey = "$($row.ObjectType)|$objectText"
        $schema = $schemaOf[$objectKey]
        if ($null -eq $schema) { $schema = [string](Get-ObjectSchema $objectText $row.ObjectType); $schemaOf[$objectKey] = $schema }
        if ($schema) { $group.Schemas[$(if ($serverText -and $databaseText) { "$serverText / $databaseText / $schema" } elseif ($serverText -or $databaseText) { "$serverText$databaseText / $schema" } else { $schema })] = 1 }
        $sourceTypeText = [string]$row.SourceType
        $key = "$($workspaceText.Length):$workspaceText$($modelText.Length):$modelText$($sourceTypeText.Length):$sourceTypeText$($serverText.Length):$serverText$($databaseText.Length):$databaseText$($schema.Length):$schema"
        $entry = $schemas[$key]
        if ($null -eq $entry) {
            $entry = @{ Schema = $schema; First = $row; Review = $false; Objects = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal); Tables = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal) }
            $schemas[$key] = $entry
        }
        if ($objectText) { $entry.Objects[$objectText] = 1 }
        $entry.Tables[$tableText] = 1
        if ($review) { $entry.Review = $true }
    }
    $modelRows = @(foreach ($group in $models.Values) {
        [pscustomobject][ordered]@{ Workspace = $group.First.Workspace; SemanticModel = $group.First.SemanticModel
            Tables = $group.Tables.Count; SourceRows = $group.Count; ToReview = $group.Review
            Sources = @($group.Sources.Keys) -join "`n"; Schemas = @($group.Schemas.Keys) -join "`n" }
    })
    $schemaRows = @(foreach ($entry in $schemas.Values) {
        $row = $entry.First
        [pscustomobject][ordered]@{ Workspace = $row.Workspace; SemanticModel = $row.SemanticModel
            SourceType = $row.SourceType; Server = $row.Server; Database = $row.Database; Schema = $entry.Schema
            Objects = $entry.Objects.Count; Tables = $entry.Tables.Count; NeedsReview = if ($entry.Review) { 'yes' } else { '' } }
    })
    # Sorted by Workspace, SemanticModel, Server, Database, Schema, SourceType (ordinal, ignoring case): one key per row,
    # the columns separated by a character that sorts before any other.
    $sortKeys = [string[]]@(foreach ($row in $schemaRows) { @([string]$row.Workspace, [string]$row.SemanticModel, [string]$row.Server, [string]$row.Database, [string]$row.Schema, [string]$row.SourceType) -join [char]0 })
    # (The [Array] / [IComparer] casts select the non-generic overload: with the generic one, PowerShell sorts a COPY of the items.)
    if ($schemaRows.Count -gt 1) { [Array]::Sort([Array]$sortKeys, [Array]$schemaRows, [Collections.IComparer][StringComparer]::OrdinalIgnoreCase) }
    @{ Name = 'Sources'; Rows = $Rows; HighlightColumn = 'NeedsReview'; LinkColumns = @('QueryFile', 'SqlFile') }
    @{ Name = 'Models'; Rows = $modelRows; WrapColumns = @('Sources', 'Schemas') }
    @{ Name = 'Schemas'; Rows = $schemaRows; HighlightColumn = 'NeedsReview'
        Columns = @('Workspace', 'SemanticModel', 'SourceType', 'Server', 'Database', 'Schema', 'Objects', 'Tables', 'NeedsReview') }
}
function Initialize-TSqlParser([string]$ExplicitPath) {
    $file = 'Microsoft.SqlServer.TransactSql.ScriptDom.dll'
    $candidates = if ($ExplicitPath) { @(Get-Item -LiteralPath $ExplicitPath) } else {
        # The SqlServer module may sit in any folder of PSModulePath, including a Documents folder redirected to OneDrive.
        $moduleFolders = @(@($env:PSModulePath -split ';') + @(
                (Join-Path $env:ProgramFiles 'PowerShell\Modules')
                (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules')
                (Join-Path $HOME 'Documents\PowerShell\Modules')) | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\', '/') })
        $patterns = @(
            "$env:ProgramFiles\Microsoft SQL Server Management Studio *\Release\Common7\IDE\Extensions\Application\$file"
            "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio *\Common7\IDE\Extensions\Application\$file"
            "$env:ProgramFiles\Microsoft SQL Server\*\DAC\bin\$file", "${env:ProgramFiles(x86)}\Microsoft SQL Server\*\DAC\bin\$file"
            "$env:ProgramFiles\Microsoft SQL Server\*\SDK\Assemblies\$file", "${env:ProgramFiles(x86)}\Microsoft SQL Server\*\SDK\Assemblies\$file"
            foreach ($folder in $moduleFolders) { "$folder\SqlServer\*\coreclr\$file" }
            "$HOME\.nuget\packages\microsoft.sqlserver.transactsql.scriptdom\*\lib\net*\$file"
        ) | Where-Object { $_ -match '\A[A-Za-z]:\\' } | Select-Object -Unique
        # Program Files and Windows (admin-writable only) before the user-profile locations, pattern order kept otherwise
        # (stable sort); newest first within one pattern. Every candidate must still carry a valid Microsoft signature.
        $machineWide = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:windir) | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') + '\' }
        $patterns = @($patterns | Sort-Object -Stable { $pattern = $_; -not @($machineWide | Where-Object { $pattern.StartsWith($_, [StringComparison]::OrdinalIgnoreCase) }).Count })
        $seenFiles = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        @(foreach ($pattern in $patterns) { Get-Item -Path $pattern -ErrorAction Ignore | Where-Object { $seenFiles.Add($_.FullName) } | Sort-Object { [version]($_.VersionInfo.FileVersion -replace '[^\d\.].*$') } -Descending })
    }
    foreach ($candidate in $candidates) {
        try {
            # Loading a DLL runs its code: accept nothing but a valid Microsoft signature.
            # The file is held open WITHOUT write / delete sharing while its signature is checked and its bytes are read
            # and loaded: it cannot be replaced between the check and the load. (Get-AuthenticodeSignature -Content only
            # understands scripts, not DLLs, hence the check by path under that lock.)
            $stream = [IO.File]::Open($candidate.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            try {
                $signature = Get-AuthenticodeSignature -LiteralPath $candidate.FullName
                if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch '(\A|, )O=Microsoft Corporation(,|\z)') { throw "not signed by Microsoft ($($signature.Status))" }
                $bytes = [byte[]]::new($stream.Length); $read = 0
                while ($read -lt $bytes.Length) { $got = $stream.Read($bytes, $read, $bytes.Length - $read); if ($got -le 0) { throw 'file could not be read completely' }; $read += $got }
                $assembly = [Reflection.Assembly]::Load($bytes)
            } finally { $stream.Dispose() }
            try { $types = $assembly.GetTypes() } catch [Reflection.ReflectionTypeLoadException] { $types = @($_.Exception.Types | Where-Object { $_ }) }
            $parserType = $types | Where-Object { $_.Name -match '\ATSql(\d+)Parser\z' } | Sort-Object { [int]($_.Name -replace '\D') } -Descending | Select-Object -First 1
            $fragmentType = $assembly.GetType('Microsoft.SqlServer.TransactSql.ScriptDom.TSqlFragment')
            if ($null -eq $parserType -or $null -eq $fragmentType) { throw 'no TSql parser type found' }
            $tableReferenceType = $assembly.GetType('Microsoft.SqlServer.TransactSql.ScriptDom.TableReference')
            if ($null -eq $tableReferenceType) { throw 'TableReference type not found' }
            # Node kinds that never contain a table reference (their sub-trees are not walked).
            $leafTypes = @(foreach ($name in 'Identifier', 'MultiPartIdentifier', 'Literal', 'ColumnReferenceExpression', 'DataTypeReference', 'VariableReference', 'GlobalVariableExpression') { $leafType = $assembly.GetType("Microsoft.SqlServer.TransactSql.ScriptDom.$name"); if ($null -ne $leafType) { $leafType } })
            $script:TSql = @{ ParserType = $parserType; FragmentType = $fragmentType; TableReferenceType = $tableReferenceType; Properties = @{}; Path = $candidate.FullName; Version = $assembly.GetName().Version; LeafTypes = $leafTypes }
            return
        } catch {
            if ($ExplicitPath) { throw "Cannot use -ScriptDomPath '$ExplicitPath': $($_.Exception.Message)" }
            # Never silent: falling back to the scanner changes how much the results can be trusted.
            Write-UiWarning "Microsoft ScriptDom found but not usable ($($candidate.FullName)): $($_.Exception.Message)"
        }
    }
}
function Get-TSqlObjects([string]$Sql, [int]$Depth = 0) {
    # Ok = $false when the text is not valid T-SQL (the caller then uses the built-in scanner).
    # Incomplete = $true when the statement reads something this analysis cannot see into (OPENROWSET, ad hoc sources,
    # EXEC of a non-constant string, a table-source node type unknown to this script): never reported as certain.
    $failed = [pscustomobject]@{ Ok = $false; Objects = @(); Incomplete = $false }
    try {
        $parser = [Activator]::CreateInstance($script:TSql.ParserType, @($true)); $errors = $null
        $root = $parser.Parse([IO.StringReader]::new($Sql), [ref]$errors)
    } catch { return $failed }
    if ($null -eq $root -or ($null -ne $errors -and $errors.Count -gt 0)) { return $failed }
    $fragmentType = $script:TSql.FragmentType; $tableReferenceType = $script:TSql.TableReferenceType; $cache = $script:TSql.Properties
    $objects = [Collections.Generic.List[object]]::new(); $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $state = @{ Incomplete = $false }
    $emit = { param($schemaObject, [string]$objectType, [int]$at)
        if ($null -eq $schemaObject) { return }
        $name = Format-SqlName ([string[]]@($schemaObject.Identifiers | ForEach-Object { [string]$_.Value }))
        if ($seen.Add($name)) { $objects.Add([pscustomobject]@{ Name = $name; Type = $objectType; At = $at }) } }
    $embedded = { param([string]$text, [string]$label, [int]$at)     # constant SQL inside OPENQUERY / OPENROWSET / EXEC('…')
        if ($Depth -ge 3) { $state.Incomplete = $true; return }
        $inner = Get-SqlObjects $text ($Depth + 1)
        if ($inner.Incomplete -or -not $inner.Objects.Count) { $state.Incomplete = $true }
        foreach ($o in $inner.Objects) { $name = "$($o.Name)"; if ($seen.Add("$label|$name")) { $objects.Add([pscustomobject]@{ Name = $name; Type = "$($o.Type) ($label)"; At = $at }) } } }
    # Structural table references (their children are walked) and rowset functions that read no object.
    $harmless = 'QualifiedJoin', 'UnqualifiedJoin', 'JoinParenthesisTableReference', 'OdbcQualifiedJoinTableReference', 'QueryDerivedTable', 'InlineDerivedTable', 'PivotedTableReference', 'UnpivotedTableReference', 'DataModificationTableReference', 'OpenJsonTableReference', 'OpenXmlTableReference', 'BuiltInFunctionTableReference', 'GlobalFunctionTableReference', 'VariableTableReference', 'VariableMethodCallTableReference', 'PredictTableReference', 'VectorSearchTableReference', 'AIGenerateChunksTableReference', 'AIGenerateFixedChunksTableReference'
    $statements = @(foreach ($batch in $root.Batches) { foreach ($statement in $batch.Statements) { $statement } })
    foreach ($statement in $statements) {
        # T-SQL accepts a bare word at the start of a batch as a procedure call without EXEC: "EVALUATE 'Sales'" (DAX) or
        # any other non-SQL text then reads as "procedure EVALUATE". Without the EXEC keyword, nothing is certain.
        if ($statement.GetType().Name -eq 'ExecuteStatement' -and $statement.FirstTokenIndex -ge 0) {
            $first = $root.ScriptTokenStream[$statement.FirstTokenIndex]
            if ($null -eq $first -or $first.Text -notmatch '\A(?i:EXEC|EXECUTE)\z') { $state.Incomplete = $true }
        }
        # CTE names and aliases are scoped to their statement: "WITH x AS (…) SELECT … FROM x; SELECT … FROM x" reads a real table x.
        $nodes = [Collections.Generic.List[object]]::new(); $cte = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); $aliases = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $stack = [Collections.Generic.Stack[object]]::new(); $stack.Push($statement)
        while ($stack.Count) {      # generic walk through every fragment-typed property (no compiled visitor needed)
            $node = $stack.Pop(); $type = $node.GetType()
            # What is known of a node type is worked out once: its role, and its fragment-typed properties (one fragment, or a list).
            $shape = $cache[$type]
            if ($null -eq $shape) {
                $shape = @{ Cte = $type.Name -eq 'CommonTableExpression'; Reference = $tableReferenceType.IsAssignableFrom($type) -or $type.Name -in 'ExecutableProcedureReference', 'ExecutableStringList'
                    Single = [Collections.Generic.List[Reflection.PropertyInfo]]::new(); Many = [Collections.Generic.List[Reflection.PropertyInfo]]::new() }
                # Names, literals, column references and data types cannot hold a table reference: not walked into (most
                # nodes of a query are of these kinds, and each property read costs).
                $leaf = $false; foreach ($leafType in $script:TSql.LeafTypes) { if ($leafType.IsAssignableFrom($type)) { $leaf = $true; break } }
                foreach ($property in $(if ($leaf) { @() } else { $type.GetProperties('Public,Instance') })) {
                    if ($property.GetIndexParameters().Count) { continue }
                    if ($fragmentType.IsAssignableFrom($property.PropertyType)) { $shape.Single.Add($property) }
                    elseif ($property.PropertyType.IsGenericType -and $fragmentType.IsAssignableFrom($property.PropertyType.GetGenericArguments()[0])) { $shape.Many.Add($property) }
                }
                $cache[$type] = $shape
            }
            if ($shape.Cte) { if ($null -ne $node.ExpressionName) { [void]$cte.Add($node.ExpressionName.Value) } }
            elseif ($shape.Reference) { $nodes.Add($node) }
            foreach ($property in $shape.Single) {
                $value = $property.GetValue($node)
                if ($null -eq $value) { continue }
                if ($property.Name -ceq 'Alias') { [void]$aliases.Add([string]$value.Value) }
                $stack.Push($value)
            }
            foreach ($property in $shape.Many) {
                $value = $property.GetValue($node)
                if ($null -ne $value) { foreach ($item in $value) { if ($null -ne $item) { $stack.Push($item) } } }
            }
        }
        # Stable sort by position (equal positions keep the order of discovery).
        $sorted = $nodes.ToArray(); $offsets = [long[]]::new($sorted.Count); for ($k = 0; $k -lt $sorted.Count; $k++) { $offsets[$k] = [long]$sorted[$k].StartOffset * 1000000 + $k }
        [Array]::Sort([Array]$offsets, [Array]$sorted)
        foreach ($node in $sorted) {
            $kind = $node.GetType().Name; $at = $node.StartOffset
            switch ($kind) {
                'NamedTableReference' {
                    $parts = @($node.SchemaObject.Identifiers)
                    if ($parts.Count -eq 1 -and $cte.Contains($parts[0].Value)) { break }
                    if ($parts.Count -eq 1 -and $parts[0].Value -ieq 'dual') { break }      # Oracle's one-row pseudo table
                    if ($parts[-1].Value.StartsWith('#')) { break }        # #temp table: created by the query itself, not a source
                    # UPDATE t SET … FROM dbo.Real AS t : the target "t" is the alias, not an object
                    if ($parts.Count -eq 1 -and $null -eq $node.Alias -and $aliases.Contains($parts[0].Value)) { break }
                    & $emit $node.SchemaObject 'Table/View' $at
                }
                'SchemaObjectFunctionTableReference' { & $emit $node.SchemaObject 'Function' $at }
                { $_ -in 'FullTextTableReference', 'SemanticTableReference' } { & $emit $node.TableName 'Table/View' $at }        # CONTAINSTABLE, FREETEXTTABLE, SEMANTIC…TABLE
                { $_ -in 'ChangeTableChangesTableReference', 'ChangeTableVersionTableReference' } { & $emit $node.Target 'Table/View' $at }   # CHANGETABLE
                'ExecutableProcedureReference' { if ($null -ne $node.ProcedureReference -and $null -ne $node.ProcedureReference.ProcedureReference) { & $emit $node.ProcedureReference.ProcedureReference.Name 'Stored procedure' $at } else { $state.Incomplete = $true } }
                'ExecutableStringList' {
                    $literals = @($node.Strings | Where-Object { $_.GetType().Name -eq 'StringLiteral' })
                    if ($literals.Count -and $literals.Count -eq @($node.Strings).Count) { & $embedded ((@($literals | ForEach-Object { $_.Value }) -join '')) 'EXEC string' $at } else { $state.Incomplete = $true }
                }
                'OpenQueryTableReference' { if ($null -ne $node.Query) { & $embedded $node.Query.Value "OPENQUERY $($node.LinkedServer.Value)" $at } else { $state.Incomplete = $true } }
                'OpenRowsetTableReference' {
                    if ($null -ne $node.Object) { & $emit $node.Object 'Table/View (OPENROWSET)' $at }
                    elseif ($null -ne $node.Query) { & $embedded $node.Query.Value 'OPENROWSET' $at }
                    else { $state.Incomplete = $true }
                }
                'BulkOpenRowset' { foreach ($file in @($node.DataFiles)) { if ($null -ne $file -and $seen.Add([string]$file.Value)) { $objects.Add([pscustomobject]@{ Name = [string]$file.Value; Type = 'File (OPENROWSET BULK)'; At = $at }) } } }
                default { if ($kind -notin $harmless) { $state.Incomplete = $true } }        # AdHocTableReference, future node types…
            }
        }
    }
    $found = $objects.ToArray(); $offsets = [long[]]::new($found.Count); for ($k = 0; $k -lt $found.Count; $k++) { $offsets[$k] = [long]$found[$k].At * 1000000 + $k }
    if ($found.Count -gt 1) { [Array]::Sort([Array]$offsets, [Array]$found) }
    return [pscustomobject]@{ Ok = $true; Objects = @(foreach ($o in $found) { [pscustomobject]@{ Name = $o.Name; Type = $o.Type } }); Incomplete = $state.Incomplete }
}
$script:SqlTokenRx = [regex]::new('\G(?:(?<ws>\s+)|(?<lc>--[^\n]*)|(?<bc>/\*)|(?<str>[Nn]?''(?:[^'']|'''')*'')|(?<br>\[(?:[^\]]|\]\])*\])|(?<dq>"(?:[^"]|"")*")|(?<bt>`(?:[^`]|``)*`)|(?<ph>\uE000\d+\uE001)|(?<id>[\w@#$]+)|(?<p>[\s\S]))')
# Row-generating functions (Snowflake FLATTEN / SPLIT_TO_TABLE / GENERATOR included): they read no stored object.
$script:SqlRowsetWords = [Collections.Generic.HashSet[string]]::new([string[]]('SELECT VALUES TABLE UNNEST OPENROWSET OPENQUERY OPENJSON OPENXML STRING_SPLIT GENERATE_SERIES JSON_TABLE XMLTABLE FLATTEN SPLIT_TO_TABLE GENERATOR' -split ' '), [StringComparer]::OrdinalIgnoreCase)
function Get-TokenSqlObjects([string]$Sql, [string[]]$Placeholders = @()) {
    # Dialect-neutral analysis: a real lexer (nested comments, strings, [bracket] "quoted" `backtick` identifiers),
    # then the table factors that follow FROM / JOIN / APPLY (comma joins, derived tables, aliases and hints
    # included), EXEC targets, CTE names excluded, and FROM inside EXTRACT / TRIM / SUBSTRING / IS DISTINCT FROM ignored.
    # $Placeholders: the unresolved M fragments ({x} / {?x}) inserted in this text. Each one is ONE token (a name where a
    # name is expected); any other brace is plain SQL ({d '2024-01-01'} ODBC escapes, {[Measures]} MDX sets...).
    for ($p = 0; $p -lt $Placeholders.Count; $p++) { $Sql = $Sql.Replace($Placeholders[$p], "$([char]0xE000)$p$([char]0xE001)") }
    $tokens = [Collections.Generic.List[object]]::new(); $position = 0; $length = $Sql.Length
    while ($position -lt $length) {
        $m = $script:SqlTokenRx.Match($Sql, $position)
        if (-not $m.Success -or $m.Length -eq 0) { $position++; continue }
        if ($m.Groups['bc'].Success) {
            $depth = 1; $k = $position + 2
            while ($k -lt $length -and $depth -gt 0) {
                if ($k + 1 -lt $length -and $Sql[$k] -ceq '/' -and $Sql[$k + 1] -ceq '*') { $depth++; $k += 2 }
                elseif ($k + 1 -lt $length -and $Sql[$k] -ceq '*' -and $Sql[$k + 1] -ceq '/') { $depth--; $k += 2 }
                else { $k++ }
            }
            $position = $k; continue
        }
        $position += $m.Length
        if ($m.Groups['ws'].Success -or $m.Groups['lc'].Success) { continue }
        foreach ($kind in 'str', 'br', 'dq', 'bt', 'ph', 'id', 'p') { if ($m.Groups[$kind].Success) { $tokens.Add(@{ T = $kind; V = $m.Value }); break } }
    }
    $count = $tokens.Count
    $isPunct = { param([int]$k, [string]$text) $k -ge 0 -and $k -lt $count -and $tokens[$k].T -eq 'p' -and $tokens[$k].V -ceq $text }
    $isWord = { param([int]$k, [string]$word) $k -ge 0 -and $k -lt $count -and $tokens[$k].T -eq 'id' -and $tokens[$k].V -ieq $word }
    $isName = { param([int]$k) $k -ge 0 -and $k -lt $count -and ($tokens[$k].T -in 'br', 'dq', 'bt', 'ph' -or ($tokens[$k].T -eq 'id' -and $tokens[$k].V -notmatch '\A[\d@]')) }
    $unquote = { param([int]$k)
        $value = $tokens[$k].V
        switch ($tokens[$k].T) { 'br' { $value.Substring(1, $value.Length - 2).Replace(']]', ']') } 'dq' { $value.Substring(1, $value.Length - 2).Replace('""', '"') } 'bt' { $value.Substring(1, $value.Length - 2).Replace('``', '`') } 'ph' { $Placeholders[[int]$value.Substring(1, $value.Length - 2)] } default { $value } } }
    $skipParens = { param([int]$k)   # $k is on '(' ; returns the index after the matching ')'
        $depth = 0
        while ($k -lt $count) { if (& $isPunct $k '(') { $depth++ } elseif (& $isPunct $k ')') { $depth--; if ($depth -eq 0) { return $k + 1 } }; $k++ }
        return $count }
    $readName = { param([int]$k)
        if (-not (& $isName $k)) { return $null }
        $parts = [Collections.Generic.List[string]]::new()
        $push = { param([int]$at) $text = & $unquote $at; if ($tokens[$at].T -eq 'bt' -and $text.Contains('.')) { $parts.AddRange([string[]]($text -split '\.')) } else { $parts.Add($text) } }   # `project.dataset.table` (BigQuery) is a multi-part name
        & $push $k; $k++
        while (& $isPunct $k '.') {
            $k++
            if (& $isPunct $k '.') { $parts.Add(''); continue }       # db..table
            if (& $isName $k) { & $push $k; $k++ } else { break }
        }
        return @{ Parts = [string[]]$parts.ToArray(); Next = $k } }
    $cte = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $testCte = { param([int]$k)       # name [(columns)] AS [NOT] [MATERIALIZED] (
        if (& $isWord $k 'RECURSIVE') { $k++ }
        if (-not (& $isName $k)) { return }
        $name = & $unquote $k; $k++
        if (& $isPunct $k '(') { $k = & $skipParens $k }
        if (-not (& $isWord $k 'AS')) { return }
        $k++
        while ((& $isWord $k 'NOT') -or (& $isWord $k 'MATERIALIZED')) { $k++ }
        if (& $isPunct $k '(') { [void]$cte.Add($name) } }
    $objects = [Collections.Generic.List[object]]::new(); $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $addObject = { param([string[]]$parts, [string]$type, [int]$at) $name = Format-SqlName $parts; if ($seen.Add($name)) { $objects.Add([pscustomobject]@{ Name = $name; Type = $type; At = $at }) } }
    $readTable = { param([int]$k)      # one table factor: a name, a function call, or a derived table
        if ((& $isWord $k 'LATERAL') -or (& $isWord $k 'ONLY')) { $k++ }
        if (& $isPunct $k '(') {
            # ( SELECT … ) is a derived table: its own FROM is met by the main scan. ( a JOIN b ON … ) is a joined
            # table: its first factor is NOT introduced by a FROM, so it is read here.
            if (-not ((& $isWord ($k + 1) 'SELECT') -or (& $isWord ($k + 1) 'WITH') -or (& $isWord ($k + 1) 'VALUES'))) { & $readTable ($k + 1) }
            return
        }
        $nameAt = $k; $name = & $readName $k
        if ($null -eq $name) { return }
        $isFunction = & $isPunct $name.Next '('
        $single = $name.Parts.Count -eq 1
        if ($name.Parts[-1].StartsWith('#')) { return }                   # #temp table: created by the query itself, not a source
        if ($isFunction -and $single) {
            $word = $name.Parts[0]; $inside = $name.Next + 1
            if ($word -match '\A(?i:TABLE|CONTAINSTABLE|FREETEXTTABLE|SEMANTICKEYPHRASETABLE|SEMANTICSIMILARITYTABLE|SEMANTICSIMILARITYDETAILSTABLE)\z') { & $readTable $inside; return }   # the object is the first argument
            if ($word -ieq 'CHANGETABLE') { if ((& $isWord $inside 'CHANGES') -or (& $isWord $inside 'VERSION')) { $inside++ }; & $readTable $inside; return }
            if ($word -match '\A(?i:OPENQUERY|OPENROWSET|OPENDATASOURCE)\z') { $flags.PassThrough = $true; return }     # SQL sent elsewhere as text: not visible here
            # FROM t (NOLOCK) : old-style table hint, not a function call
            if ($isFunction -and (& $isPunct ($inside + 1) ')') -and $inside -lt $count -and $tokens[$inside].V -match '\A(?i:NOLOCK|READUNCOMMITTED|READCOMMITTED|REPEATABLEREAD|SERIALIZABLE|HOLDLOCK|UPDLOCK|ROWLOCK|PAGLOCK|TABLOCK|TABLOCKX|XLOCK|READPAST|NOWAIT|NOEXPAND|SNAPSHOT)\z') { $isFunction = $false }
        }
        if ($single -and $script:SqlRowsetWords.Contains($name.Parts[0])) { return }
        if ($single -and -not $isFunction -and $cte.Contains($name.Parts[0])) { return }
        if ($single -and -not $isFunction -and $name.Parts[0] -ieq 'dual') { return }      # Oracle's one-row pseudo table
        & $addObject $name.Parts $(if ($isFunction) { 'Function' } else { 'Table/View' }) $nameAt }
    # One scope per parenthesis level. InFrom = "between FROM and the end of its clause": there, a comma starts another
    # table factor (FROM a, b  /  FROM a CROSS APPLY f(x) s, b  /  FROM a JOIN b ON ..., c); elsewhere it may start a CTE.
    $clauseEnd = [Collections.Generic.HashSet[string]]::new([string[]]('WHERE GROUP HAVING ORDER UNION EXCEPT INTERSECT MINUS OPTION WINDOW QUALIFY LIMIT FETCH OFFSET SELECT SET VALUES RETURNING FOR WHEN' -split ' '), [StringComparer]::OrdinalIgnoreCase)
    $hasFrom = $false; $flags = @{ PassThrough = $false }; $scopes = [Collections.Generic.Stack[object]]::new(); $scopes.Push(@{ Function = $false; InFrom = $false })
    for ($i = 0; $i -lt $count; $i++) {
        $token = $tokens[$i]
        if ($token.T -eq 'p') {
            if ($token.V -ceq '(') { $scopes.Push(@{ InFrom = $false; Function = ($i -gt 0 -and $tokens[$i - 1].T -eq 'id' -and $tokens[$i - 1].V -match '\A(?i:EXTRACT|TRIM|SUBSTRING|OVERLAY|POSITION)\z') }) }
            elseif ($token.V -ceq ')') { if ($scopes.Count -gt 1) { [void]$scopes.Pop() } }
            elseif ($token.V -ceq ',') { if ($scopes.Peek().InFrom) { & $readTable ($i + 1) } else { & $testCte ($i + 1) } }
            elseif ($token.V -ceq ';') { $scopes.Peek().InFrom = $false; if ($scopes.Count -eq 1) { $cte.Clear() } }     # a CTE lives in its statement only
            continue
        }
        if ($token.T -ne 'id') { continue }
        # FOR XML / FOR JSON end the FROM clause; FOR SYSTEM_TIME belongs to a table factor inside it
        if ($clauseEnd.Contains($token.V) -and -not ($token.V -ieq 'FOR' -and (& $isWord ($i + 1) 'SYSTEM_TIME'))) { $scopes.Peek().InFrom = $false }
        switch ($token.V.ToUpperInvariant()) {
            'WITH' { & $testCte ($i + 1) }
            'FROM' {
                if ($scopes.Peek().Function) { break }                               # EXTRACT(YEAR FROM d), TRIM(x FROM y) ...
                if (& $isWord ($i - 1) 'DISTINCT') { if ((& $isWord ($i - 2) 'IS') -or (& $isWord ($i - 2) 'NOT')) { break } }
                $hasFrom = $true; $scopes.Peek().InFrom = $true; & $readTable ($i + 1)
            }
            { $_ -in 'JOIN', 'APPLY' } { & $readTable ($i + 1) }
            'MERGE' { $k = $i + 1; if (& $isWord $k 'INTO') { $k++ }; & $readTable $k }
            'USING' { if (-not (& $isPunct ($i + 1) '(')) { & $readTable ($i + 1) } }   # MERGE … USING src ; not JOIN … USING (col)
            { $_ -in 'EXEC', 'EXECUTE', 'CALL' } {
                $k = $i + 1
                if (& $isPunct $k '(') { $flags.PassThrough = $true; break }           # EXEC ('dynamic sql')
                if ($k -lt $count -and $tokens[$k].T -eq 'id' -and $tokens[$k].V.StartsWith('@') -and (& $isPunct ($k + 1) '=')) { $k += 2 }   # EXEC @rc = proc
                $name = & $readName $k
                if ($null -ne $name) { & $addObject $name.Parts 'Stored procedure' $k }
            }
        }
    }
    return [pscustomobject]@{ Objects = @($objects | Sort-Object At | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Type = $_.Type } }); HasFrom = $hasFrom; PassThrough = $flags.PassThrough }
}
function Test-HasPlaceholder([string]$Text, [string[]]$Placeholders) {
    foreach ($placeholder in $Placeholders) { if ($Text.Contains($placeholder)) { return $true } }
    return $false
}
function Get-SqlObjects([string]$Sql, [int]$Depth = 0, [string[]]$Placeholders = @()) {
    # Incomplete = the analysis could not see into part of the statement; the caller must not report certainty.
    # $Placeholders = the unresolved M fragments ({x} / {?x}) that Resolve-MText inserted in $Sql. Other braces are SQL.
    $scanned = $null
    if ($Placeholders.Count) {
        # Unresolved M fragment. Where a TABLE NAME is expected it must be reported as such (dbo.{x}), never replaced.
        $scanned = Get-TokenSqlObjects $Sql $Placeholders
        foreach ($o in $scanned.Objects) { if (Test-HasPlaceholder $o.Name $Placeholders) { return [pscustomobject]@{ Objects = @($scanned.Objects); Parser = 'built-in scanner'; TSqlFailed = $false; HasFrom = $scanned.HasFrom; Incomplete = $true } } }
    }
    $tsqlFailed = $false
    if ($null -ne $script:TSql) {
        # Elsewhere a placeholder stands for a value: the T-SQL grammar sees 0. Whether that is safe (a date) or not
        # (opaque text that may hide a FROM) is decided by the caller from the way the M text was resolved.
        $forParser = $Sql
        foreach ($placeholder in @($Placeholders | Sort-Object Length -Descending)) { $forParser = $forParser.Replace($placeholder, '0') }
        $parsed = Get-TSqlObjects $forParser $Depth
        if ($parsed.Ok) { return [pscustomobject]@{ Objects = @($parsed.Objects); Parser = 'T-SQL parser'; TSqlFailed = $false; HasFrom = $true; Incomplete = $parsed.Incomplete } }
        $tsqlFailed = $true
    }
    if ($null -eq $scanned) { $scanned = Get-TokenSqlObjects $Sql $Placeholders }
    return [pscustomobject]@{ Objects = @($scanned.Objects); Parser = 'built-in scanner'; TSqlFailed = $tsqlFailed; HasFrom = $scanned.HasFrom; Incomplete = $scanned.PassThrough }
}
function Get-SafeConnectionString([string]$Text) {
    # Keep only the keys that identify the source; never copy credentials.
    $kept = foreach ($pair in ($Text -split ';')) {
        $kv = $pair -split '=', 2
        if ($kv.Count -eq 2 -and $kv[0].Trim() -match '\A(?i:dsn|driver|server|host|hostname|data source|database|initial catalog|dbq|port|provider|httppath|catalog|schema|warehouse)\z') { "$($kv[0].Trim())=$($kv[1].Trim())" }
    }
    return (@($kept) -join ';')
}
# Secrets never reach a written file, the CSV, the workbook or the web page: the value is replaced with ***. Masked only in
# the places where a secret is written, so that the SQL and M the inventory describes stay intact:
#  - a token shaped like a JWT (eyJ...), and what follows "Bearer ";
#  - URL query parameters (?sig= &key= &api_key= &access_token= &client_secret= &token=, ?code= when 8 characters or more);
#  - a record field or header whose NAME says secret ([Password = "..."], [Authorization = "..."], [#"x-api-key" = "..."]);
#  - a connection-string key (pwd, password, accountkey, client_secret...): written key=value with nothing around "=", or
#    after a ";". The value is a quoted text ("...", ""..."" inside an M string, '...', {...}) or runs up to ; " ' & or a space;
#    a value starting with @ or ( or a function call (password = HASHBYTES(...)) is SQL, never masked.
$script:SecretKeys = 'pwd|password|passwd|accountkey|account key|sharedaccesskey|sharedaccesssignature|client_secret|clientsecret|api_key|apikey'
$script:SecretValue = '(?<v>""(?:""""|(?!"")[^\r\n])*""|"(?!")(?:""|[^"\r\n])*"|''(?:''''|[^''\r\n])*''|\{(?:\}\}|[^}\r\n])*\}|(?![@(])(?![A-Za-z_][\w.]*\s*\()[^;"''\s&]+)'
$script:SecretRules = @(
    @{ Rx = [regex]::new('(?<![A-Za-z0-9_\-])eyJ[A-Za-z0-9_\-]*(?:\.[A-Za-z0-9_\-]*){0,2}', 'CultureInvariant'); Fixed = '***'; Value = $false }
    @{ Rx = [regex]::new('(?<=\bBearer\s+)[A-Za-z0-9\-._~+/]+=*', 'IgnoreCase, CultureInvariant'); Fixed = '***'; Value = $false }
    @{ Rx = [regex]::new('(?<=[?&](?:sig|key|api_key|apikey|api-key|access_token|client_secret|token|subscription-key)=)[^&#"''\s]+', 'IgnoreCase, CultureInvariant'); Fixed = '***'; Value = $false }
    @{ Rx = [regex]::new('(?<=[?&]code=)[^&#"''\s]{8,}', 'IgnoreCase, CultureInvariant'); Fixed = '***'; Value = $false }
    @{ Rx = [regex]::new('(?<prefix>(?<![\w.])(?:#")?[\w-]*(?:secret|password|passwd|pwd|api[-_]?key|access[-_]?token|token|authorization|subscription-key)[\w-]*"?\s*=\s*)"(?!")(?:""|[^"\r\n])*"', 'IgnoreCase, CultureInvariant'); Fixed = '${prefix}"***"'; Value = $false }
    @{ Rx = [regex]::new("(?<prefix>(?<![^;`"'{\s])(?:$script:SecretKeys)=|(?<=;)\s*(?:$script:SecretKeys)\s*=\s*)$script:SecretValue", 'IgnoreCase, CultureInvariant'); Value = $true }
)
$script:SecretHint = [regex]::new('eyJ|bearer|[?&](?:sig|key|api|access|client|token|subscription|code)|secret|passw|pwd|accountkey|sharedaccess|token|authorization|api[-_]?key', 'IgnoreCase, CultureInvariant')
function Hide-Secret([string]$Text) {
    if (-not $Text -or -not $script:SecretHint.IsMatch($Text)) { return $Text }
    foreach ($rule in $script:SecretRules) {
        if (-not $rule.Rx.IsMatch($Text)) { continue }
        if ($rule.Value) {
            # The quotes around the value are kept, so that the text still reads (and an M string still closes).
            $Text = $rule.Rx.Replace($Text, {
                    param($m)
                    $value = $m.Groups['v'].Value
                    $masked = if ($value.StartsWith('""') -and $value.Length -ge 4) { '""***""' } elseif ($value.StartsWith('"')) { '"***"' } elseif ($value.StartsWith("'")) { "'***'" } elseif ($value.StartsWith('{')) { '{***}' } else { '***' }
                    $m.Groups['prefix'].Value + $masked })
        } else { $Text = $rule.Rx.Replace($Text, $rule.Fixed) }
    }
    return $Text
}
function Get-SourceLabel([string]$Function, [string]$Server) {
    if ($Function -cne 'Sql.Database' -and $Function -cne 'Sql.Databases') { return $Function }
    if ($Server -match '(?i)\.datawarehouse\.fabric\.microsoft\.com') { return "$Function (Fabric SQL endpoint)" }
    if ($Server -match '(?i)\.sql\.azuresynapse\.net') { return "$Function (Synapse)" }
    if ($Server -match '(?i)\.database\.windows\.net') { return "$Function (Azure SQL)" }
    return $Function
}
function New-SourceContext([string]$Function, [string]$Kind) {
    return @{ Function = $Function; Kind = $Kind; SourceType = $Function; Server = ''; Database = ''; Schema = ''; BasePath = ''
        Containers = [Collections.Generic.List[string]]::new(); Params = [Collections.Generic.List[string]]::new()
        BaseRow = $null; Emitted = $false; Review = $false; Inherited = $false; Wrapper = $false; NoDatabase = $false }
}
function Add-SourceRow($Rows, $Context, [string]$Object, [string]$ObjectType, [string]$DetectedBy, [bool]$Review, $Params) {
    $row = @{ Context = $Context; Object = $Object; ObjectType = $ObjectType; DetectedBy = $DetectedBy; Review = $Review
        Params = [Collections.Generic.List[string]]::new(); Database = $Context.Database; Bare = $false; Via = '' }
    if ($Params) { $row.Params.AddRange([string[]]@($Params)) }
    $Rows.Add($row); $Context.Emitted = $true
    return $row
}
function Complete-SourceContext($Rows, $Context) {
    # A wrapper (Cdm.Contents) around another connector says nothing more than that connector's rows.
    if ($null -ne $Context -and -not $Context.Emitted -and -not $Context.Inherited -and -not ($Context.Wrapper -and $Rows.Count)) {
        $row = Add-SourceRow $Rows $Context '' '' 'Source call only' $true $null
        $row.Bare = $true
    }
}
function Add-NativeSqlRows($Rows, $Context, [string]$Expression, $Scope, [string]$Origin) {
    if ($Context.Function -clike 'AnalysisServices.*') {
        # [Query = "EVALUATE ..."] on Analysis Services / a Power BI model is DAX or MDX, not SQL: never sent to the SQL
        # analysers (T-SQL would read "EVALUATE 'Sales'" as a call of a procedure EVALUATE). The query is in the .pq file.
        [void](Add-SourceRow $Rows $Context '' 'DAX/MDX query' 'DAX/MDX query (not analysed)' $true $null)
        return
    }
    Add-SqlTextRows $Rows $Context (Resolve-MText $Expression $Scope) $Origin
}
function Add-SqlTextRows($Rows, $Context, $Sql, [string]$Origin) {
    # $Sql: Text, Resolved, Opaque, Params and Placeholders, as returned by Resolve-MText (a legacy query partition passes
    # its text as is: resolved, no placeholder).
    $firstRow = $Rows.Count; $parser = 'none (text not resolvable)'
    $placeholders = [string[]]@($Sql.Placeholders)
    try {
        $rest = $Sql.Text; foreach ($placeholder in $placeholders) { $rest = $rest.Replace($placeholder, '') }
        if ($placeholders.Count -and -not $rest.Trim()) { [void](Add-SourceRow $Rows $Context '{dynamic SQL}' '' 'Native SQL (text not resolvable)' $true $Sql.Params); return }
        $parsed = Get-SqlObjects $Sql.Text 0 $placeholders; $parser = $parsed.Parser
        # Certainty is claimed only when nothing is hidden: no opaque M fragment (Text.Combine, if … then … else, an
        # unknown name: it may contain a FROM), nothing the SQL analysis could not see into (OPENROWSET, EXEC of a
        # variable…), and either the real T-SQL grammar accepted the text, or the scanner read a complete text that
        # was not expected to be T-SQL.
        $expectsTSql = $Context.Function -cmatch '\ASql\.Databases?\z' -or ($Context.ContainsKey('ExpectsTSql') -and $Context.ExpectsTSql)
        $doubt = $Sql.Opaque -or $parsed.Incomplete -or ($parsed.Parser -ne 'T-SQL parser' -and (-not $Sql.Resolved -or ($parsed.TSqlFailed -and $expectsTSql)))
        if (-not $parsed.Objects.Count) {
            # A complete, valid text without any table (SELECT NULL AS x WHERE 1 = 2, typical of measure tables) is a certainty.
            $certain = -not $doubt -and ($parsed.Parser -eq 'T-SQL parser' -or (-not $parsed.HasFrom -and $Sql.Text -match '\A\s*(?:--[^\n]*\s*|/\*[\s\S]*?\*/\s*)*SELECT\b'))
            [void](Add-SourceRow $Rows $Context '' '' $(if ($certain) { 'Native SQL (no table referenced)' } else { 'Native SQL (no object found)' }) (-not $certain) $Sql.Params)
            return
        }
        foreach ($o in $parsed.Objects) { [void](Add-SourceRow $Rows $Context $o.Name $o.Type "Native SQL ($($parsed.Parser))" ($doubt -or (Test-HasPlaceholder $o.Name $placeholders)) $Sql.Params) }
    } finally {
        # Keep the exact resolved text (and which analyser read it) before object rows are merged.
        for ($index = $firstRow; $index -lt $Rows.Count; $index++) {
            $Rows[$index].Sql = $Sql.Text; $Rows[$index].SqlOrigin = $Origin; $Rows[$index].Parser = $parser; $Rows[$index].SqlResolved = [bool]$Sql.Resolved
        }
    }
}
function Get-MSources([string]$Query, $Map, $Initial) {
    # Walks source-function calls and navigation steps in document order. Returns the rows and the last context.
    $rows = [Collections.Generic.List[object]]::new()
    $masks = Get-MMasks $Query; $clean = $masks.Clean; $mask = $masks.Mask
    # Model parameters + the steps of this query's outer let (Schema = "dbo", sqlQuery = "select ..." & x ...).
    $let = Get-LetBindings $clean $mask
    $local = New-MScope $Map $let
    # A query can open several connections (a = Sql.Database("s1",…), b = Sql.Database("s2",…)). Each let step remembers
    # the connection it holds, so that x = a{[…]}[Data] is attributed to a's server and not to the last one opened.
    $stepContext = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    $contexts = [Collections.Generic.List[object]]::new()
    $stepOf = { param([int]$position) foreach ($step in $let.Steps) { if ($position -ge $step.Start -and $position -lt $step.End) { return $step.Name } } }
    $contextOf = { param([string]$expression)      # the connection held by an identifier, if known
        $m = [regex]::Match($expression, '(?:#"((?:[^"]|"")*)"|(?<![\w\.#"])([\p{L}_][\w\.]*))\s*\z')
        if (-not $m.Success) { return $null }
        $name = if ($m.Groups[1].Success) { $m.Groups[1].Value.Replace('""', '"') } else { $m.Groups[2].Value }
        if ($stepContext.ContainsKey($name)) { return $stepContext[$name] }
        return $null }
    $events = [Collections.Generic.List[object]]::new()
    foreach ($m in $script:CallRx.Matches($mask)) { $events.Add(@{ Pos = $m.Index; Type = 'call'; Name = $m.Groups[1].Value; Open = $m.Index + $m.Length - 1 }) }
    foreach ($m in $script:NavRx.Matches($mask)) { $events.Add(@{ Pos = $m.Index; Type = 'nav'; Open = $m.Index + $m.Length - 1 }) }
    $state = @{ Context = $null }
    $pending = [Collections.Generic.List[object]]::new()   # native queries wait until their own arguments are walked
    $ensure = {
        if ($null -eq $state.Context) {
            if ($null -ne $Initial) { $c = New-SourceContext $Initial.Function $Initial.Kind; foreach ($k in 'SourceType', 'Server', 'Database', 'Schema', 'BasePath') { $c[$k] = $Initial[$k] }; $c.Containers.AddRange($Initial.Containers); $c.Params.AddRange($Initial.Params); $c.Inherited = $true; $state.Context = $c }
            else { $c = New-SourceContext '(unknown source)' 'db'; $c.Review = $true; $state.Context = $c }
        }
        $state.Context
    }
    $flush = { param([int]$Before)
        if (-not $pending.Count) { return }
        foreach ($q in $pending.ToArray()) {
            if ($q.End -ge $Before) { continue }
            $target = & $contextOf $q.Target
            if ($null -ne $target) { $state.Context = $target }
            Add-NativeSqlRows $rows (& $ensure) $q.Sql $local 'Value.NativeQuery'; [void]$pending.Remove($q)
        }
    }
    $ordered = $events.ToArray(); $positions = [int[]]@(foreach ($e in $ordered) { $e.Pos })
    [Array]::Sort([Array]$positions, [Array]$ordered)
    foreach ($e in $ordered) {
        & $flush $e.Pos
        $close = Find-Close $mask $e.Open
        if ($close -lt 0) { continue }
        $innerClean = $clean.Substring($e.Open + 1, $close - $e.Open - 1); $innerMask = $mask.Substring($e.Open + 1, $close - $e.Open - 1)
        if ($e.Type -eq 'call') {
            $arguments = @(Split-TopLevel $innerClean $innerMask ',')
            if ($e.Name -ceq 'Value.NativeQuery') { if ($arguments.Count -ge 2) { $pending.Add(@{ End = $close; Sql = $arguments[1]; Target = $arguments[0] }) }; continue }
            $def = $script:Connectors[$e.Name]
            if ($null -eq $def) {
                if ($script:StdLib.Contains(($e.Name -split '\.')[0]) -or $script:StdFunctions.Contains($e.Name)) { continue }
                $def = @{ K = 'db'; S = 0; Unknown = $true }
            }
            $c = New-SourceContext $e.Name $def.K; $state.Context = $c; $contexts.Add($c)
            $step = & $stepOf $e.Pos
            if ($step) { $stepContext[$step] = $c }
            if ($def.ContainsKey('Unknown')) { $c.Review = $true }
            if ($def.ContainsKey('Wrapper')) { $c.Wrapper = $true }
            if ($def.ContainsKey('NoDatabase')) { $c.NoDatabase = $true }
            if ($def.ContainsKey('Dynamic')) { $c.Review = $true; [void](Add-SourceRow $rows $c '{dynamic M code}' '' 'M code evaluated at refresh time' $true $null) }
            if ($def.ContainsKey('S') -and $arguments.Count -gt $def.S -and $arguments[$def.S].Trim()) {
                $v = Resolve-MText $arguments[$def.S] $local
                $c.Server = if ($def.ContainsKey('ConnectionString')) { Get-SafeConnectionString $v.Text } else { $v.Text }
                $c.Params.AddRange($v.Params); if (-not $v.Resolved) { $c.Review = $true }
            }
            if ($def.ContainsKey('D') -and $arguments.Count -gt $def.D) {
                $v = Resolve-MText $arguments[$def.D] $local
                $c.Database = $v.Text; $c.Params.AddRange($v.Params); if (-not $v.Resolved) { $c.Review = $true }
            }
            $c.SourceType = Get-SourceLabel $e.Name $c.Server
            if ($def.K -in 'file', 'container', 'url') {
                $c.BasePath = $c.Server; $c.Server = ''
                $path = $c.BasePath
                if ($e.Name -ceq 'Web.Contents' -and $arguments.Count -ge 2) {
                    $am = Get-MMasks $arguments[1]; $open = $am.Mask.IndexOf('[')
                    if ($open -ge 0) {
                        $end = Find-Close $am.Mask $open
                        if ($end -gt $open) {
                            $options = Get-MRecordFields $am.Clean.Substring($open + 1, $end - $open - 1) $am.Mask.Substring($open + 1, $end - $open - 1)
                            if ($options.Contains('RelativePath')) { $rel = Resolve-MText $options['RelativePath'] $local; $path = $path.TrimEnd('/') + '/' + $rel.Text.TrimStart('/'); if (-not $rel.Resolved) { $c.Review = $true } }
                        }
                    }
                }
                $kind = switch ($def.K) { 'file' { 'File' } 'url' { 'URL' } default { 'Folder / site' } }
                $c.BaseRow = Add-SourceRow $rows $c $path $kind 'Function argument' $c.Review $null
                continue
            }
            if ($def.ContainsKey('O') -and $arguments.Count -gt $def.O) {
                $v = Resolve-MText $arguments[$def.O] $local
                [void](Add-SourceRow $rows $c $v.Text 'Table / query' 'Function argument' (-not $v.Resolved) $v.Params)
            }
            if ($def.ContainsKey('Q') -and $arguments.Count -gt $def.Q) { Add-NativeSqlRows $rows $c $arguments[$def.Q] $local $e.Name }
            foreach ($argument in $arguments) {   # options record with a native query: [Query = "select ..."]
                $am = Get-MMasks $argument
                if (-not $am.Mask.TrimStart().StartsWith('[')) { continue }
                $open = $am.Mask.IndexOf('['); $end = Find-Close $am.Mask $open
                if ($end -le $open) { continue }
                $options = Get-MRecordFields $am.Clean.Substring($open + 1, $end - $open - 1) $am.Mask.Substring($open + 1, $end - $open - 1)
                if ($options.Contains('Query')) { Add-NativeSqlRows $rows $c $options['Query'] $local '[Query=…]' }
            }
            continue
        }
        # Navigation step: only {[...]}[Data] and {[...]}[Content] navigate a source; other lookups read rows.
        if (-not [regex]::IsMatch($mask.Substring($close + 1), '\A\s*\}\s*\[\s*(?:Data|Content)\s*\]')) { continue }
        $fields = Get-MRecordFields $innerClean $innerMask
        if (-not $fields.Count) { continue }
        $value = @{}; $parameters = [Collections.Generic.List[string]]::new(); $unresolved = $false
        foreach ($key in @($fields.Keys)) { $v = Resolve-MText $fields[$key] $local; $value[$key] = $v.Text; $parameters.AddRange($v.Params); if (-not $v.Resolved) { $unresolved = $true } }
        # The connection that is navigated: the identifier right before "{", when it is a step holding one.
        $target = & $contextOf $clean.Substring([Math]::Max(0, $e.Pos - 200), [Math]::Min(200, $e.Pos))
        if ($null -ne $target) { $state.Context = $target }
        $c = & $ensure
        $step = & $stepOf $e.Pos
        if ($step) { $stepContext[$step] = $c }      # cat = Source{…}[Data] : "cat" now holds that connection too
        if ($fields.Contains('Schema') -and $fields.Contains('Item')) { [void](Add-SourceRow $rows $c (Format-SqlName @($value['Schema'], $value['Item'])) 'Table/View' 'Navigation' $unresolved $parameters); continue }
        # {[Schema="HR"]} alone (Oracle / Teradata hierarchical navigation): the next step by Name is an object of that schema.
        if ($fields.Count -eq 1 -and $fields.Contains('Schema')) { $c.Schema = $value['Schema']; $c.Params.AddRange($parameters); continue }
        if ($fields.Contains('Name') -and $fields.Contains('Kind')) {
            switch -Regex ($value['Kind']) {
                '\A(?i:Database|Catalog)\z' { $c.Database = $value['Name']; $c.Params.AddRange($parameters); break }
                '\A(?i:Schema)\z' { $c.Schema = $value['Name']; $c.Params.AddRange($parameters); break }
                default {
                    $object = if ($c.Schema) { Format-SqlName @($c.Schema, $value['Name']) } else { Format-SqlName @($value['Name']) }
                    [void](Add-SourceRow $rows $c $object $value['Kind'] 'Navigation' $unresolved $parameters)
                }
            }
            continue
        }
        if ($fields.Contains('Item') -and $fields.Contains('Kind')) {   # Excel sheet / table / named range
            $detail = "$($value['Kind']): $($value['Item'])"
            if ($null -ne $c.BaseRow -and $c.BaseRow.Object -ceq $c.BasePath) { $c.BaseRow.Object = "$($c.BasePath) | $detail"; if ($unresolved) { $c.BaseRow.Review = $true }; $c.BaseRow.Params.AddRange($parameters) }
            else { [void](Add-SourceRow $rows $c $(if ($c.BasePath) { "$($c.BasePath) | $detail" } else { $detail }) 'File' 'Navigation' $unresolved $parameters) }
            continue
        }
        if ($fields.Contains('Id') -and $fields.Contains('ItemKind')) { [void](Add-SourceRow $rows $c $(if ($fields.Contains('Schema')) { Format-SqlName @($value['Schema'], $value['Id']) } else { Format-SqlName @($value['Id']) }) $value['ItemKind'] 'Navigation' $unresolved $parameters); continue }
        if ($fields.Count -eq 1 -and ($fields.Contains('Id') -or $fields.Contains('Key'))) {
            # {[Id="report suite"]} / {[Key="AppWorkspace"]}{[Key="Reports"]} : each step goes one level deeper; the last one is the object,
            # the previous ones its container (kept in Database, e.g. the Adobe company).
            $step = $value[@($fields.Keys)[0]]
            if ($null -ne $c.BaseRow -and $rows.Contains($c.BaseRow)) { [void]$rows.Remove($c.BaseRow); $c.Database = $(if ($c.Database) { "$($c.Database) / $($c.BaseRow.Object)" } else { $c.BaseRow.Object }) }
            $c.BaseRow = Add-SourceRow $rows $c $step 'Item' 'Navigation' $unresolved $parameters
            $c.BaseRow.IdStep = $true
            continue
        }
        # Dataflows: {[Id="Workspaces"]} then {[workspaceId=…]}{[dataflowId=…]}{[entity=…]}. The Id step was a navigation root,
        # not the object: it is dropped once container keys or the entity follow.
        $dropIdStep = {
            if ($null -ne $c.BaseRow -and $c.BaseRow.ContainsKey('IdStep') -and $rows.Contains($c.BaseRow)) {
                [void]$rows.Remove($c.BaseRow); $c.BaseRow = $null
                $c.Emitted = $false; foreach ($other in $rows) { if ([object]::ReferenceEquals($other.Context, $c)) { $c.Emitted = $true; break } }
            } }
        if ($fields.Contains('entity')) { & $dropIdStep; [void](Add-SourceRow $rows $c $value['entity'] 'Dataflow entity' 'Navigation' $unresolved $parameters); continue }
        $containerKeys = @($fields.Keys | Where-Object { $_ -match '(?i)\A(workspaceId|lakehouseId|warehouseId|dataflowId|datamartId|databaseId|artifactId)\z' })
        if ($containerKeys.Count) { & $dropIdStep; foreach ($key in $containerKeys) { $c.Containers.Add("$key=$($value[$key])") }; $c.Params.AddRange($parameters); continue }
        if ($fields.Contains('Name')) {
            if ($c.Kind -eq 'db' -and -not $c.Database -and -not $c.Schema -and -not $c.NoDatabase -and -not $fields.Contains('Signature')) { $c.Database = $value['Name']; $c.Params.AddRange($parameters); continue }
            $objectType = if ($c.Kind -ne 'db') { 'File / item' } elseif ($c.NoDatabase) { 'Object' } else { 'Table/View' }
            $object = if ($c.Kind -ne 'db' -or $c.NoDatabase) { $value['Name'] } elseif ($c.Schema) { Format-SqlName @($c.Schema, $value['Name']) } else { Format-SqlName @($value['Name']) }
            if ($fields.Contains('Folder Path')) { $object = $value['Folder Path'].TrimEnd('/', '\') + $(if ($value['Folder Path'] -match '\\') { '\' } else { '/' }) + $value['Name'] }
            if ($null -ne $c.BaseRow -and $rows.Contains($c.BaseRow) -and $c.BaseRow.Object -ceq $c.BasePath) { [void]$rows.Remove($c.BaseRow) }
            $row = Add-SourceRow $rows $c $object $objectType 'Navigation' $unresolved $parameters
            continue
        }
    }
    & $flush ([int]::MaxValue)
    foreach ($context in $contexts) { Complete-SourceContext $rows $context }      # a connection from which nothing was read
    return [pscustomobject]@{ Rows = $rows; Context = $state.Context }
}
function Get-ReferencedExpressions([string]$Query, $Queries, $Literals) {
    # Names of the other queries of the model (shared expressions AND other tables) used as identifiers in this
    # query, in order of appearance. A step of the query's own let hides a model query of the same name.
    $masks = Get-MMasks $Query; $names = [Collections.Generic.List[string]]::new()
    # M is case-sensitive (date is not the table Date); the ordered dictionary of queries is not, hence this exact-case
    # set, built once per set of queries.
    if (-not [object]::ReferenceEquals($script:ExactNamesOf, $Queries) -or $script:ExactNames.Count -ne $Queries.Count) {
        $script:ExactNames = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::Ordinal)
        foreach ($key in @($Queries.Keys)) { $script:ExactNames[[string]$key] = 1 }
        $script:ExactNamesOf = $Queries
    }
    $exactNames = $script:ExactNames
    # First the identifiers that name another query at all (most queries reference none): the finer checks below run
    # only when there is at least one.
    $candidates = @(foreach ($m in $script:IdentRx.Matches($masks.Clean)) {
            $name = if ($m.Groups[1].Success) { $m.Groups[1].Value.Replace('""', '"') } else { $m.Groups[2].Value }
            if ($null -ne $exactNames[$name] -and $null -eq $Literals[$name]) { $m }
        })
    if (-not $candidates.Count) { return [string[]]@() }
    $ownSteps = (Get-LetBindings $masks.Clean $masks.Mask).Map
    # (Sales) => Sales : a function parameter hides a model query of the same name
    $functionParameters = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($fm in [regex]::Matches($masks.Mask, '\(([^()]*)\)\s*(?:as\s+[\w\. ]+?\s*)?=>')) { foreach ($pm in [regex]::Matches($fm.Groups[1].Value, '(?:\boptional\s+)?([\p{L}_]\w*)(?:\s+as\s+[\w\. ]+)?\s*(?:,|\z)')) { [void]$functionParameters.Add($pm.Groups[1].Value) } }
    $keywords = $script:MKeywords
    # Field names are not references: [Sales Region] or x[Sales] (field access) and the NAME side of a record
    # [Schema = ..., Item = ...]. Only the value side of a record can reference a query.
    $mask = $masks.Mask; $skip = [bool[]]::new($mask.Length); $open = [Collections.Generic.Stack[object]]::new()
    foreach ($m in $script:ReferenceScanRx.Matches($mask)) {     # only the characters that matter: ( ) [ ] { } , =
        $ch = $m.Value; $i = $m.Index
        if ($ch -ceq '[') { $open.Push(@{ At = $i; Depth = 0; Segment = $i + 1; Names = [Collections.Generic.List[int[]]]::new(); IsRecord = $false }); continue }
        if (-not $open.Count) { continue }
        $frame = $open.Peek()
        if ($ch -ceq ']') {
            [void]$open.Pop()
            if ($frame.IsRecord) { foreach ($range in $frame.Names) { if ($range[1] -gt $range[0]) { [Array]::Fill($skip, $true, $range[0], $range[1] - $range[0]) } } }
            else { [Array]::Fill($skip, $true, $frame.At, $i - $frame.At + 1) }
            continue
        }
        if ($ch -ceq '(' -or $ch -ceq '{') { $frame.Depth++ } elseif ($ch -ceq ')' -or $ch -ceq '}') { $frame.Depth-- }
        elseif ($frame.Depth -eq 0 -and $ch -ceq '=' -and $frame.Segment -ge 0 -and $mask[[Math]::Min($i + 1, $mask.Length - 1)] -cne '>') { $frame.IsRecord = $true; $frame.Names.Add([int[]]($frame.Segment, $i)); $frame.Segment = -1 }
        elseif ($frame.Depth -eq 0 -and $ch -ceq ',') { $frame.Segment = $i + 1 }
    }
    foreach ($m in $candidates) {
        if ($masks.Mask[$m.Index] -ne $masks.Clean[$m.Index]) { continue }   # inside a string literal
        if ($skip[$m.Index]) { continue }
        if ($m.Groups[2].Success) {
            if ($keywords.Contains($m.Groups[2].Value) -or $functionParameters.Contains($m.Groups[2].Value)) { continue }   # only #"each" can name a query called each
            # type date, as text, is number : the name of a type, not a query that happens to be called date / text / number
            if ($masks.Mask.Substring([Math]::Max(0, $m.Index - 12), [Math]::Min(12, $m.Index)) -cmatch '(?<![\w\.])(?:type|as|is|nullable)\s+\z') { continue }
        }
        $name = if ($m.Groups[1].Success) { $m.Groups[1].Value.Replace('""', '"') } else { $m.Groups[2].Value }
        if (-not $ownSteps.ContainsKey($name) -and -not $names.Contains($name)) { $names.Add($name) }
    }
    return $names.ToArray()
}
$script:ExactNames = $null; $script:ExactNamesOf = $null
function Get-QuerySources([string]$Query, $ModelInfo, [string[]]$Stack) {
    # Sources of an M query, including those of the shared expressions and tables it references (transitively).
    $all = [Collections.Generic.List[object]]::new(); $initial = $null
    foreach ($name in @(Get-ReferencedExpressions $Query $ModelInfo.Queries $ModelInfo.Literals)) {
        if (-not $name -or $Stack -contains $name) { continue }
        if ($Stack.Count -gt 50) { throw "Queries nested deeper than 50 levels near '$name'." }
        if (-not $ModelInfo.Cache.ContainsKey($name)) { $ModelInfo.Cache[$name] = Get-QuerySources $ModelInfo.Queries[$name] $ModelInfo ($Stack + $name) }
        $sub = $ModelInfo.Cache[$name]
        foreach ($row in $sub.Rows) { $copy = $row.Clone(); if (-not $copy.Via) { $copy.Via = $name }; $all.Add($copy) }
        if ($null -eq $initial -and $null -ne $sub.Context) { $initial = $sub.Context }
    }
    $own = Get-MSources $Query $ModelInfo.Literals $initial
    $usedInherited = $false; foreach ($row in $own.Rows) { if ($row.Context.Inherited) { $usedInherited = $true; break } }
    foreach ($row in $all.ToArray()) { if ($row.Bare -and $usedInherited) { [void]$all.Remove($row) } }
    foreach ($row in $own.Rows) { if ($row.Context.Inherited -and -not $row.Via -and $null -ne $initial) { $row.Via = '(connection from shared expression)' }; $all.Add($row) }
    $context = if ($null -ne $own.Context) { $own.Context } else { $initial }
    return [pscustomobject]@{ Rows = $all; Context = $context }
}

function Write-ModelQueryFile([string]$RelativePath, [string]$Text) {
    $path = [IO.Path]::GetFullPath([IO.Path]::Combine($modelOut, $RelativePath))
    if (-not $path.StartsWith($modelOut.TrimEnd('\', '/') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing query file outside model folder: $path" }
    # Each folder is checked once per model (Assert-SafeDirectory remembers the folders it checked): the model folder was
    # checked, deleted and created again just before, and its sub-folders are created here.
    $parent = [IO.Path]::GetDirectoryName($path)
    if (-not $script:SafeFolders.Contains($parent)) {
        Assert-NotOneLake $parent
        Assert-SafeDirectory $parent
        [void][IO.Directory]::CreateDirectory($parent)
        [void]$script:SafeFolders.Add($parent)
    }
    # Passwords and keys in connection strings or URLs are masked in everything written (and so in the web page).
    $content = (Hide-Secret $Text).Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText($path, $content, $utf8)
    # Kept for the web page: reading back thousands of files that were just written is slow (antivirus scans).
    $script:WrittenFiles[$path] = $content
}
# Query files written during this run: full path -> text.
$script:WrittenFiles = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)

# ------------------------------------------------------------------ web page
# model_sources.html: the page below with the inventory embedded. The rows and the text of their query files are
# serialised to JSON, gzip-compressed and base64-encoded into one <script type="application/octet-stream"> element that
# the page decodes in the browser (DecompressionStream): base64 holds no '<', so no data can close the element or be
# read as markup, and the page itself only ever inserts data as text. It loads nothing from the network.
function Export-HtmlReport([string]$Path, [object[]]$Rows, [string]$Root) {
    $columns = @('Workspace', 'SemanticModel', 'Table', 'SourceType', 'Server', 'Database', 'Object', 'ObjectType', 'Mode', 'Via', 'Parameters', 'NeedsReview', 'QueryFile', 'SqlFile')
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/'); $rootPrefix = $rootFull + '\'
    $list = [Collections.Generic.List[object]]::new()
    # Keys: the exact spellings used by the rows (the page looks them up case-sensitively). Reads are cached per file.
    $files = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $read = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $plainFolders = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)   # folders already checked: no link
    foreach ($row in $Rows) {
        $record = [Collections.Generic.Dictionary[string, string]]::new()
        foreach ($column in $columns) { $record[$column] = [string]$row.$column }
        $list.Add($record)
        # The query files of the row, read back from the output folder. Paths come from the CSV of earlier runs too:
        # only regular files inside -OutputPath are read, with no symbolic link or junction on the way (a link could
        # point anywhere and put that file's content in the page). Cloud files of OneDrive are reparse points without
        # being links: accepted inside the OneDrive folder only (the rule of Assert-NotLinked).
        foreach ($relative in $record.QueryFile, $record.SqlFile) {
            if (-not $relative -or $null -ne $files[$relative]) { continue }
            $full = [IO.Path]::GetFullPath([IO.Path]::Combine($Root, $relative))   # a rooted $relative stays itself: refused below
            # A file written by this run: its text is known, nothing to read back (and so no link to fear).
            $text = $script:WrittenFiles[$full]
            if ($null -ne $text) { $files[$relative] = $text; continue }
            if (-not $read.ContainsKey($full)) {
                if ($full.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -and [IO.File]::Exists($full)) {
                    $linked = $false; $cursor = $full
                    while ($cursor.Length -gt $rootFull.Length -and -not $plainFolders.Contains($cursor)) {
                        $entry = if ($cursor -eq $full) { [IO.FileInfo]::new($cursor) } else { [IO.DirectoryInfo]::new($cursor) }
                        if ($null -ne $entry.LinkTarget -or (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -and -not (Test-UnderOneDrive $cursor))) { $linked = $true; break }
                        if ($cursor -ne $full) { [void]$plainFolders.Add($cursor) }
                        $cursor = [IO.Path]::GetDirectoryName($cursor)
                    }
                    if (-not $linked) { $text = [IO.File]::ReadAllText($full) }
                }
                $read[$full] = $text
            }
            if ($null -ne $read[$full]) { $files[$relative] = [string]$read[$full] }
        }
    }
    $payload = [Collections.Generic.Dictionary[string, object]]::new()
    # ISO 8601 with milliseconds, whatever the culture (JavaScript's Date parses exactly this form everywhere).
    $payload.generatedAtUtc = [datetime]::UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", [cultureinfo]::InvariantCulture); $payload.generator = 'Parse-SemanticModels.ps1'
    $payload.rows = $list; $payload.files = $files
    $page = $script:HtmlTemplate.Replace("`r`n", "`n")
    $marker = [regex]::Matches($page, '(?m)^<!--MODEL_SOURCES_DATA-->$')
    if ($marker.Count -ne 1) { throw 'Internal error: the web page template has no single data marker.' }
    $utf8 = [Text.UTF8Encoding]::new($false)
    $head = $utf8.GetBytes($page.Substring(0, $marker[0].Index) + '<script id="model-sources-data" type="application/octet-stream" data-encoding="gzip-base64">')
    $tail = $utf8.GetBytes('</script>' + $page.Substring($marker[0].Index + $marker[0].Length))
    # Written to a new temporary file (CreateNew: never through an existing link), then moved over the target: the move
    # replaces the directory entry itself, so a link planted at the target is replaced, not followed. The data is streamed
    # (JSON -> gzip -> base64 -> file) between the two halves of the page: no copy of the whole inventory as one string.
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        $stream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew)
        try {
            $stream.Write($head, 0, $head.Length)
            $base64 = [Security.Cryptography.CryptoStream]::new($stream, [Security.Cryptography.ToBase64Transform]::new(), [Security.Cryptography.CryptoStreamMode]::Write, $true)
            try {
                $gzip = [IO.Compression.GZipStream]::new($base64, [IO.Compression.CompressionLevel]::Optimal, $true)
                try { [Text.Json.JsonSerializer]::Serialize([IO.Stream]$gzip, [object]$payload, $payload.GetType(), [Text.Json.JsonSerializerOptions]::new()) } finally { $gzip.Dispose() }
            } finally { $base64.Dispose() }   # writes the last base64 characters and the padding
            $stream.Write($tail, 0, $tail.Length)
        } finally { $stream.Dispose() }
        [IO.File]::Move($temporary, $Path, $true)
    } finally { if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) } }
}
# The page (HTML, CSS and JavaScript in one file). Kept inline so that the tool stays three scripts; the line
# <!--MODEL_SOURCES_DATA--> is replaced with the data.
$script:HtmlTemplate = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light dark">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'self' 'unsafe-inline'; style-src 'unsafe-inline'; img-src 'none'; connect-src 'none'; font-src 'none'; base-uri 'none'; form-action 'none'">
<title>Fabric Source Atlas · Semantic model inventory</title>
<style>
:root{--bg:#f5f7f9;--panel:#fff;--ink:#202e40;--muted:#5f6b7a;--line:#e4e9ef;--accent:#06717c;--tint:#e7f5f4;--nav:#edf1f5;--warn:#b66d18;--warn-bg:#fff4e2;--shadow:0 3px 16px #23364d05;--code:#f8fafc;--tok-key:#6a50c8;--tok-string:#0f7562;color-scheme:light}
@media(prefers-color-scheme:dark){:root:not([data-theme=light]){--bg:#111922;--panel:#19232f;--ink:#e2eaf4;--muted:#97a8bc;--line:#2e3b4c;--accent:#66d4ce;--tint:#1d3b42;--nav:#253241;--warn:#e8b366;--warn-bg:#443524;--shadow:none;--code:#141e28;--tok-key:#b7a6f7;--tok-string:#5ccfb4;color-scheme:dark}}
:root[data-theme=dark]{--bg:#111922;--panel:#19232f;--ink:#e2eaf4;--muted:#97a8bc;--line:#2e3b4c;--accent:#66d4ce;--tint:#1d3b42;--nav:#253241;--warn:#e8b366;--warn-bg:#443524;--shadow:none;--code:#141e28;--tok-key:#b7a6f7;--tok-string:#5ccfb4;color-scheme:dark}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:13px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}button,input,select{font:inherit}button,select{cursor:pointer}button{color:inherit;background:var(--panel);border:1px solid var(--line);border-radius:7px;padding:7px 12px;transition:background .15s,border-color .15s}button:hover{background:var(--nav);border-color:var(--muted)}button:focus-visible,a:focus-visible,input:focus-visible,select:focus-visible,summary:focus-visible{outline:3px solid var(--accent);outline-offset:3px}a{color:var(--accent);text-decoration:none}button:disabled{opacity:.45;cursor:default}h1,h2,h3,p{margin:0}h1{font-size:27px;letter-spacing:-.8px;font-weight:650}h2{font-size:15px;font-weight:650}h3{font-size:13px}small,.muted{color:var(--muted)}.eyebrow{font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--muted)}.flex{display:flex;align-items:center;gap:10px}.between{justify-content:space-between}.wrap{flex-wrap:wrap}.stack{display:grid;gap:18px}.pill{display:inline-block;vertical-align:middle;max-width:100%;overflow:hidden;text-overflow:ellipsis;background:var(--nav);border-radius:5px;padding:3px 7px;font-size:11px;white-space:nowrap}.pill.teal{background:var(--tint);color:var(--accent)}.pill.warn{background:var(--warn-bg);color:var(--warn)}.dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--accent);flex:none}.link{padding:0;border:0;background:transparent;color:var(--accent);text-align:left;overflow-wrap:anywhere}.link:hover{background:transparent;text-decoration:underline}.primary{background:var(--accent);color:var(--panel);border-color:var(--accent)}.primary:hover{filter:brightness(1.08);background:var(--accent)}.icon-btn{min-width:34px;padding:6px 9px}.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip-path:inset(50%)}
header{height:72px;display:flex;align-items:center;gap:32px;padding:0 26px;border-bottom:1px solid var(--line);background:var(--panel);position:sticky;top:0;z-index:20}.brand{min-width:188px;display:flex;align-items:center;gap:10px;font-size:18px;font-weight:700;letter-spacing:-.6px}.brand svg{width:32px;height:32px;color:var(--accent)}.brand small{display:block;font-size:10px;font-weight:500;letter-spacing:1px}.search-wrap{position:relative;max-width:690px;flex:1}.search-icon{position:absolute;left:13px;top:9px;color:var(--muted);font-size:19px}#search{width:100%;padding:11px 48px 11px 40px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink)}kbd{font:11px inherit;border:1px solid var(--line);border-radius:4px;padding:0 5px;color:var(--muted)}.search-wrap>kbd{position:absolute;right:12px;top:11px}#suggestions{position:absolute;left:0;right:0;top:48px;background:var(--panel);border:1px solid var(--line);box-shadow:0 10px 40px #0002;border-radius:10px;padding:6px;max-height:440px;overflow:auto}#suggestions[hidden]{display:none}.suggestion{display:flex;width:100%;border:0;text-align:left;gap:12px;align-items:center;padding:9px}.suggestion span:last-child{min-width:0;overflow-wrap:anywhere}.suggestion.active{background:var(--tint)}.header-end{margin-left:auto;display:flex;gap:16px;align-items:center}.offline{font-size:11px;color:var(--muted);white-space:nowrap}
.shell{display:grid;grid-template-columns:222px minmax(0,1fr);min-height:calc(100vh - 72px)}aside{border-right:1px solid var(--line);padding:28px 16px;background:var(--panel);display:flex;flex-direction:column;gap:22px}nav{display:grid;gap:5px}nav button{display:flex;gap:11px;align-items:center;text-align:left;border:0;background:transparent;padding:10px 13px;font-size:12px}nav button.active{background:var(--tint);color:var(--accent);font-weight:650}.nav-icon{font-size:17px;width:19px;text-align:center}nav .count{margin-left:auto;font-size:10px;background:var(--nav);border-radius:4px;padding:1px 5px}.aside-note{border:1px solid var(--line);border-radius:8px;padding:13px;font-size:11px;color:var(--muted);margin-top:auto}.aside-note strong{display:block;color:var(--ink);margin-bottom:6px}main{min-width:0;padding:27px 30px 20px;max-width:1900px;width:100%;margin:auto}.breadcrumb{font-size:11px;color:var(--muted);margin-bottom:17px}.page-heading{display:flex;align-items:flex-start;justify-content:space-between;gap:15px;margin-bottom:22px}.page-heading p{color:var(--muted);margin-top:5px}.page-heading h1{overflow-wrap:anywhere}.heading-actions{display:flex;gap:8px;flex-shrink:0;padding-top:7px}.stats{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:14px;margin-bottom:22px}.stat{background:var(--panel);border:1px solid var(--line);box-shadow:var(--shadow);border-radius:9px;padding:16px 18px;text-align:left}.stat .value{display:block;font-size:28px;font-weight:650;letter-spacing:-1px;margin:5px 0}.stat .label{color:var(--muted);font-size:11px}.stat .hint{font-size:10px;color:var(--accent)}.panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:var(--shadow);overflow:hidden}.panel-head{padding:17px 20px;display:flex;justify-content:space-between;align-items:center;gap:12px;border-bottom:1px solid var(--line)}.panel-head p{font-size:11px;color:var(--muted);margin-top:3px}.panel-body{padding:18px 20px}.grid-two{display:grid;grid-template-columns:1fr 1fr;gap:18px}.grid-impact{display:grid;grid-template-columns:minmax(0,1fr) 310px;gap:18px}.rank-row{display:flex;align-items:center;gap:12px;width:100%;text-align:left;padding:11px 0;border:0;border-radius:0;border-bottom:1px solid var(--line);background:none}.rank-row:last-child{border-bottom:0}.rank-row:hover{background:var(--bg)}.rank-num{width:21px;font:11px ui-monospace,monospace;color:var(--muted)}.rank-name{flex:1;min-width:0}.rank-name strong{display:block;font-size:12px;font-weight:550;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.rank-name small{font-size:10px}.bar-track{height:4px;background:var(--nav);border-radius:4px;margin-top:7px;max-width:92%}.bar-fill{height:100%;border-radius:4px;background:var(--accent);opacity:.65}.rank-value{font:12px ui-monospace,monospace;color:var(--muted);white-space:nowrap}.filters{display:flex;align-items:center;gap:10px;padding:12px 16px;margin:0 0 18px;background:var(--panel);border:1px solid var(--line);border-radius:9px;flex-wrap:wrap}.filters select{max-width:260px;background:var(--bg);color:var(--ink);border:1px solid var(--line);border-radius:6px;padding:7px}.filters label{display:flex;gap:7px;align-items:center;font-size:11px}.filters .explain{margin-left:auto;color:var(--muted);font-size:10px}.model-card{padding:17px 20px;border-bottom:1px solid var(--line)}.model-title{font-size:14px;font-weight:650;color:var(--ink)}.model-card summary{cursor:pointer;color:var(--accent);font-size:11px;margin-top:11px}.chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:12px}.chips button{font-size:10px;padding:3px 7px;max-width:100%;overflow-wrap:anywhere}.model-sources{margin-top:12px;border-top:1px solid var(--line)}.source-row{display:grid;grid-template-columns:90px minmax(0,1fr) 110px;gap:10px;padding:9px 0;border-bottom:1px solid var(--line);font-size:11px;align-items:center}.source-row small{display:block;overflow-wrap:anywhere}.empty{padding:35px;text-align:center;color:var(--muted)}.load{display:block;margin:14px auto}.table-wrap{overflow:auto}table{border-collapse:collapse;width:100%;font-size:12px}th{text-align:left;color:var(--muted);font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;background:var(--bg)}th,td{padding:12px 17px;border-bottom:1px solid var(--line);vertical-align:top}td{max-width:390px;overflow-wrap:anywhere}td small{display:block;margin-top:3px}footer{display:flex;justify-content:space-between;gap:12px;margin-top:27px;padding-top:16px;border-top:1px solid var(--line);color:var(--muted);font-size:10px;flex-wrap:wrap}
.lineage-panel{margin-bottom:22px}.lineage-panel>.panel-head{flex-wrap:wrap}.diagram-toolbar{display:flex;align-items:center;gap:5px;flex-wrap:wrap}.diagram-toolbar select{background:var(--panel);border:1px solid var(--line);color:var(--muted);padding:6px;border-radius:6px;font-size:11px}.diagram-shell{position:relative;background-color:var(--panel);background-image:radial-gradient(var(--line) .8px,transparent .8px);background-size:18px 18px}.diagram{display:block;width:100%;height:455px;touch-action:pan-y pinch-zoom;cursor:grab;user-select:none}.diagram:active{cursor:grabbing}.diagram .node{cursor:pointer;transition:opacity .16s}.diagram .node rect.main{fill:var(--panel);stroke:var(--line);stroke-width:1.2;filter:drop-shadow(0 3px 3px #18263909)}.diagram .node:hover rect.main,.diagram .node:focus rect.main{stroke:var(--accent);stroke-width:2}.diagram text{fill:var(--ink);font-family:inherit;pointer-events:none}.diagram .node-sub{fill:var(--muted);font-size:9px}.diagram .lane-label{fill:var(--muted);font-size:10px;font-weight:650;letter-spacing:1px}.diagram .edge{fill:none;stroke-opacity:.27;transition:opacity .16s,stroke-opacity .16s}.diagram .edge.hot{stroke-opacity:.85}.diagram .dim{opacity:.075}.diagram .aggregate rect.main{stroke-dasharray:4 3;fill:var(--bg)}.diagram-footer{padding:11px 18px;display:flex;gap:15px;align-items:center;justify-content:space-between;border-top:1px solid var(--line);font-size:10px;color:var(--muted)}.legend{display:flex;gap:12px;flex-wrap:wrap}.diagram-note{padding:9px 18px;background:var(--bg);border-top:1px solid var(--line);font-size:10px;color:var(--muted)}.diagram-animate{animation:arrive .3s ease-out}@keyframes arrive{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
dialog{background:var(--panel);color:var(--ink);border:1px solid var(--line);border-radius:12px;padding:0;max-width:1100px;width:88vw;max-height:86vh;box-shadow:0 24px 90px #0004}dialog::backdrop{background:#111d3066;backdrop-filter:blur(3px)}dialog .dialog-head{position:sticky;top:0;background:var(--panel);z-index:2;padding:18px 22px;border-bottom:1px solid var(--line)}dialog .dialog-body{padding:18px 22px}.code-tabs{display:flex;gap:7px;flex-wrap:wrap;margin:12px 0}.code-tabs button{font-size:11px;overflow-wrap:anywhere}.code-tabs button.active{color:var(--accent);background:var(--tint);border-color:var(--accent)}.code-block{font:12px/1.8 ui-monospace,SFMono-Regular,Consolas,monospace;background:var(--code);border:1px solid var(--line);border-radius:7px;overflow:auto;max-height:53vh;padding:10px 0;tab-size:4}.code-line{display:flex;min-height:21px}.line-no{color:var(--muted);width:54px;flex-shrink:0;text-align:right;padding-right:16px;user-select:none;font-size:10px}.line-code{white-space:pre;padding-right:20px}.tok-key{color:var(--tok-key)}.tok-string{color:var(--tok-string)}.tok-comment{color:var(--muted);font-style:italic}.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);padding:10px 20px;background:var(--ink);color:var(--panel);border-radius:8px;z-index:100;box-shadow:0 4px 25px #0003;font-size:12px}.toast[hidden]{display:none}.context-banner{display:flex;gap:12px;align-items:center;background:var(--tint);border:1px solid color-mix(in srgb,var(--accent) 20%,transparent);padding:13px 17px;border-radius:8px;margin-bottom:18px;font-size:12px;color:var(--accent)}.context-banner strong{font-weight:650}.context-banner span{overflow-wrap:anywhere}.section-gap{margin-top:20px}
@media(min-width:1600px){main{padding:32px 42px}.diagram{height:500px}}
@media(max-width:1150px){.shell{grid-template-columns:180px minmax(0,1fr)}header{gap:16px}.brand{min-width:155px}main{padding:24px 20px}.grid-impact{grid-template-columns:minmax(0,1fr) 260px}.offline{display:none}.stat{padding:13px}.heading-actions{flex-wrap:wrap}.diagram{height:410px}}
@media(max-width:850px){header{padding:0 15px;gap:12px}.brand{min-width:0;font-size:15px}.brand small{display:none}.shell{grid-template-columns:1fr}aside{padding:10px 15px;border-right:0;border-bottom:1px solid var(--line)}aside>.eyebrow,.aside-note,aside>div{display:none}nav{display:flex;overflow:auto}nav button{white-space:nowrap}.grid-impact,.grid-two{grid-template-columns:1fr}.page-heading{flex-wrap:wrap}.stats{gap:8px}.stat .value{font-size:23px}.header-end{gap:0}.diagram-toolbar{flex-wrap:wrap}.panel-head{align-items:flex-start}.diagram{height:380px}}
@media(max-width:560px){.brand>span{display:none}.stats{grid-template-columns:repeat(2,1fr)}.stat:last-child{grid-column:span 2}main{padding:20px 12px}.source-row{grid-template-columns:60px 1fr}.source-row>*:last-child{grid-column:2}.diagram-footer{flex-wrap:wrap}.heading-actions{width:100%}dialog{width:96vw}.filters select{max-width:180px}}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation:none!important;transition:none!important}}
.style-switch{display:flex;border:1px solid var(--line);border-radius:7px;padding:2px;background:var(--bg)}.style-switch button{border:0;background:transparent;padding:4px 9px;font-size:11px}.style-switch button[aria-pressed=true]{background:var(--panel);color:var(--accent);box-shadow:0 1px 5px #0001}.source-chip{border-color:color-mix(in srgb,var(--group-color) 35%,var(--line));background:color-mix(in srgb,var(--group-color) 8%,var(--panel))}.source-chip.selected{border-color:var(--group-color);background:color-mix(in srgb,var(--group-color) 22%,var(--panel));font-weight:650;box-shadow:inset 0 0 0 1px var(--group-color)}.source-chip .dot{background:var(--group-color);margin-right:5px}.diagram .flow-band{fill-opacity:.28;transition:opacity .16s,fill-opacity .16s}.diagram .flow-band.hot{fill-opacity:.8}.diagram .flow-node text{paint-order:stroke;stroke:var(--panel);stroke-width:4px;stroke-linejoin:round}.diagram .flow-node:focus .flow-hit,.diagram .flow-node:hover .flow-hit{stroke:var(--accent);stroke-width:1;stroke-dasharray:3 3}.diagram .flow-reveal{animation:flow-reveal .55s ease-out both}@keyframes flow-reveal{from{clip-path:inset(0 100% 0 0)}to{clip-path:inset(0 0 0 0)}}.loading-state{padding:45px 20px;color:var(--muted)}.loading-state h2{color:var(--ink);margin-bottom:8px}.header-credit{font-size:10px;color:var(--muted);white-space:nowrap}.header-credit a{color:var(--accent)}@media(max-width:1250px){.header-credit{display:none}}h1[tabindex="-1"]:focus{outline:none}
</style>
</head>
<body>
<header>
  <a class="brand" href="#home" aria-label="Source Atlas home"><svg viewBox="0 0 40 40" fill="none" aria-hidden="true"><path d="M6 27c-4-6-1-18 9-19 12-2 17 14 8 18-6 3-12-3-8-7 3-3 7 1 4 3" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"/><path d="M5 29h21c8 0 10-6 8-12m-4 1-1-5m5 4 3-5" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"/><circle cx="29" cy="12" r="1.6" fill="currentColor"/><circle cx="37" cy="11" r="1.6" fill="currentColor"/></svg><span>Source Atlas<small>SEMANTIC MODEL EXPLORER</small></span></a>
  <div class="search-wrap"><span class="search-icon" aria-hidden="true">⌕</span><input id="search" placeholder="Search schemas, objects, models…" aria-label="Search inventory" role="combobox" aria-autocomplete="list" aria-controls="suggestions" aria-expanded="false" autocomplete="off"><kbd>/</kbd><div id="suggestions" role="listbox" hidden></div></div>
  <div class="header-end"><span class="header-credit"><a href="https://lazysnail.net" target="_blank" rel="noopener">lazysnail.net</a></span><button id="theme" class="icon-btn" aria-label="Switch color theme" title="Color theme">◐</button></div>
</header>
<div class="shell">
<aside>
 <span class="eyebrow" style="padding-left:13px">Explore inventory</span>
 <nav aria-label="Main navigation">
  <button data-view="home"><span class="nav-icon">▦</span>Overview</button>
  <button data-view="models"><span class="nav-icon">◫</span>Models<span id="models-count" class="count"></span></button>
  <button data-view="sources"><span class="nav-icon">⌘</span>Source catalog</button>
  <button data-view="lineage"><span class="nav-icon">⑂</span>Lineage explorer</button>
  <button data-view="review"><span class="nav-icon">⚑</span>To review<span id="review-count" class="count"></span></button>
 </nav>
 <div style="border-top:1px solid var(--line);padding:20px 13px"><p class="eyebrow">Follow the dependencies</p><p class="muted" style="font-size:11px;margin-top:9px">Start with a source.<br>Find every affected model.<br>See the whole picture.</p></div>
 <section class="aside-note"><strong>Your inventory stays here.</strong>All exploration runs in this browser. No uploads, services, or external resources.</section>
</aside>
<main id="main" aria-busy="true"><div class="loading-state" role="status"><h2>Opening your inventory</h2><p>Preparing the local data. This may take a moment.</p><p>If this message stays, this page cannot run here (scripts blocked, a preview pane, or an old browser): download the file and open it in a current version of Microsoft Edge, Google Chrome, Firefox, or Safari 16.4 or later.</p></div><noscript>This inventory requires JavaScript.</noscript></main>
</div>
<dialog id="dialog" aria-labelledby="dialog-title"><div class="dialog-head flex between"><div><span class="eyebrow" id="dialog-kicker">Source details</span><h2 id="dialog-title"></h2></div><button id="dialog-close" aria-label="Close dialog">✕</button></div><div class="dialog-body" id="dialog-body"></div></dialog>
<div class="toast" id="toast" role="status" hidden></div>
<!--MODEL_SOURCES_DATA-->
<script>
'use strict';
// Production loader: the inert payload is never interpreted as HTML or JavaScript.
async function loadModelSources() {
  const payload = document.getElementById('model-sources-data');
  if (!payload) throw new Error('No inventory data was found in this page. Run Parse-SemanticModels.ps1 again to rebuild it.');
  if (payload.dataset.encoding !== 'gzip-base64') throw new Error('The embedded inventory uses an unsupported encoding. Expected gzip-base64.');
  if (typeof DecompressionStream === 'undefined') throw new Error('This browser is too old to open the inventory. Open the file in a current version of Microsoft Edge, Google Chrome, Firefox, or Safari 16.4 or later.');
  // Yield one paint before decoding a potentially large payload.
  await new Promise(resolve => setTimeout(resolve, 30));
  try {
    const binary = atob(payload.textContent.replace(/\s/g, ''));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
    return JSON.parse(await new Response(stream).text());
  } catch {
    throw new Error('The embedded inventory could not be decoded. It may be incomplete or damaged. Regenerate the report and try again.');
  }
}
(async function startAtlas() {
// All inventory text goes through textContent. Composite JSON keys preserve arbitrary names.
const $ = id => document.getElementById(id);
function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = String(text);
  // A pill shortens long names with an ellipsis: the full name stays readable on hover.
  if (text !== undefined && node.classList.contains('pill')) node.title = String(text);
  return node;
}
function button(text, action, className = '') {
  const node = el('button', className, text); node.type = 'button'; node.addEventListener('click', action); return node;
}
function append(parent, ...nodes) { nodes.forEach(node => parent.append(node)); return parent; }
const key = (...parts) => JSON.stringify(parts);
const fmt = n => n.toLocaleString();
const compare = (a, b) => a.localeCompare(b);
const sqlObjectTypes = new Set(['Table', 'View', 'Table/View', 'Stored procedure', 'Function']);
function schemaOf(row) {
  // Bracket/double-quoted identifiers can contain dots and escaped closing quotes.
  if (!sqlObjectTypes.has(row.ObjectType) || /databricks/i.test(row.SourceType)) return '';
  const name = row.Object || ''; let quote = '';
  for (let i = 0; i < name.length; i++) {
    const char = name[i];
    if (quote) { if (char === quote) { if (name[i + 1] === quote) i++; else quote = ''; } }
    else if (char === '[') quote = ']';
    else if (char === '"') quote = '"';
    else if (char === '.') {
      const part = name.slice(0, i).trim();
      if (part.startsWith('[') && part.endsWith(']')) return part.slice(1, -1).replace(/]]/g, ']');
      if (part.startsWith('"') && part.endsWith('"')) return part.slice(1, -1).replace(/""/g, '"');
      return part;
    }
  }
  return '';
}
function buildIndex(data) {
  const maps = Object.fromEntries(['model','table','schema','object','workspace','server','database','type'].map(type => [type,new Map()]));
  const rows = data.rows.map((raw, index) => {
    const r = {...raw, index};
    r.modelKey = key(r.Workspace, r.SemanticModel);
    r.tableKey = key(r.Workspace, r.SemanticModel, r.Table);
    r.isSql = sqlObjectTypes.has(r.ObjectType);
    r.schemaLabel = schemaOf(r);
    r.schema = r.schemaLabel.toLowerCase();
    r.serverKey = r.Server.toLowerCase();
    r.databaseKey = key(r.serverKey, r.Database.toLowerCase());
    r.objectKey = key(r.SourceType, r.serverKey, r.Database.toLowerCase(), r.isSql ? r.Object.toLowerCase() : r.Object);
    const values = {
      model:[r.modelKey,r.SemanticModel,r.Workspace], table:[r.tableKey,r.Table,r.SemanticModel+' · '+r.Workspace],
      schema:[r.schema,r.schemaLabel,'SQL schema · all servers'], object:[r.objectKey,r.Object || '(unnamed source)',[r.Server,r.Database,r.SourceType].filter(Boolean).join(' · ')],
      workspace:[r.Workspace,r.Workspace,'Workspace'], server:[r.serverKey,r.Server,'Server'],
      database:[r.databaseKey,r.Database,r.Server || 'No server'], type:[r.SourceType,r.SourceType,'Source type']
    };
    for (const [type, [id, label, detail]] of Object.entries(values)) {
      if ((type === 'schema' || type === 'server') && !id || type === 'database' && !r.Database) continue;
      let item = maps[type].get(id);
      if (!item) { item = {type,id,label:label || '(unspecified)',detail,rows:[],models:new Set(),spellings:new Map()}; maps[type].set(id,item); }
      const spelling=label || '(unspecified)';item.spellings.set(spelling,(item.spellings.get(spelling)||0)+1);
      item.rows.push(r); item.models.add(r.modelKey);
    }
    return r;
  });
  const search = Object.values(maps).flatMap(map => [...map.values()]);
  // Map insertion order provides the first-seen tie break. Only copied rows are changed.
  search.forEach(item=>{let best=0;for(const [spelling,count] of item.spellings){if(count>best){item.label=spelling;best=count;}}});
  rows.forEach(r=>{
    r.Server=maps.server.get(r.serverKey)?.label || r.Server;
    r.Database=maps.database.get(r.databaseKey)?.label || r.Database;
    r.schemaLabel=maps.schema.get(r.schema)?.label || '';
    if(r.isSql&&r.Object)r.Object=maps.object.get(r.objectKey).label;
  });
  maps.database.forEach(item=>{item.detail=item.rows[0].Server || 'No server';});
  maps.object.forEach(item=>{const r=item.rows[0];item.detail=[r.Server,r.Database,r.SourceType].filter(Boolean).join(' · ');});
  search.forEach(item => item.search = (item.label+' '+item.detail).toLowerCase());
  return {rows,maps,search};
}
function normalizeId(type,id,indexMaps) {
  if(type==='schema'||type==='server')return id.toLowerCase();
  if(type==='database'||type==='object'){
    try{
      const parts=JSON.parse(id);
      if(type==='database'&&Array.isArray(parts)&&parts.length===2&&parts.every(p=>typeof p==='string'))return key(...parts.map(p=>p.toLowerCase()));
      if(type==='object'&&Array.isArray(parts)&&parts.length===4&&parts.every(p=>typeof p==='string')){
        const exact=key(parts[0],parts[1].toLowerCase(),parts[2].toLowerCase(),parts[3]);
        const folded=key(parts[0],parts[1].toLowerCase(),parts[2].toLowerCase(),parts[3].toLowerCase());
        return indexMaps.object.get(folded)?.rows.some(r=>r.isSql)?folded:exact;
      }
    }catch{/* Invalid composite identifiers remain unmatched. */}
  }
  return id;
}
// Hash contract: #<type>=encodeURIComponent(id)&with=encodeURIComponent(JSON.stringify([[type,id],...])).
// JSON tuples preserve arbitrary punctuation; only SQL identifiers are case-folded.
// Only these route types exist: anything else in the address (#constructor, #__proto__...) opens the overview.
const ROUTE_TYPES=new Set(['home','models','sources','lineage','review','model','table','schema','object','workspace','server','database','type']);
function canonicalRoute(route,indexMaps) {
  const type=ROUTE_TYPES.has(route.type)?route.type:'home',id=type==='home'?'':normalizeId(type,typeof route.id==='string'?route.id:'',indexMaps),seen=new Set([key(type,id)]);
  const withItems=[];
  for(const pair of route.withItems||[]){
    if(!Array.isArray(pair)||pair.length!==2||!['schema','type','object'].includes(pair[0])||typeof pair[1]!=='string')continue;
    const entry=[pair[0],normalizeId(pair[0],pair[1],indexMaps)],token=key(...entry);
    if(!seen.has(token)){seen.add(token);withItems.push(entry);}
  }
  return {type,id,withItems};
}
function encodeRoute(route,indexMaps) {
  const normalized=canonicalRoute(route,indexMaps);
  return '#'+normalized.type+(normalized.id?'='+encodeURIComponent(normalized.id):'')+(normalized.withItems.length?'&with='+encodeURIComponent(JSON.stringify(normalized.withItems)):'');
}
function decodeRoute(hash,indexMaps) {
  const [subject,...params]=hash.replace(/^#/,'').split('&'),split=subject.indexOf('=');
  let id='',withItems=[];
  try{id=split<0?'':decodeURIComponent(subject.slice(split+1));}catch{/* Invalid escape sequences are unmatched. */}
  const encoded=params.find(p=>p.startsWith('with='));
  if(encoded){try{const parsed=JSON.parse(decodeURIComponent(encoded.slice(5)));if(Array.isArray(parsed))withItems=parsed;}catch{/* Ignore malformed optional filters. */}}
  return canonicalRoute({type:(split<0?subject:subject.slice(0,split))||'home',id,withItems},indexMaps);
}
function intersectSources(indexMaps,subject,withItems,qualifies=()=>true) {
  const subjectRows=(indexMaps[subject.type]?.get(subject.id)?.rows||[]).filter(qualifies);
  let models=new Set(subjectRows.map(r=>r.modelKey));
  const additional=withItems.map(([type,id])=>indexMaps[type]?.get(id));
  additional.forEach(item=>{models=new Set([...models].filter(id=>item?.models.has(id)));});
  const selected=new Map();
  [subjectRows,...additional.map(item=>item?.rows||[])].forEach(group=>group.forEach(r=>{if(models.has(r.modelKey))selected.set(r.index,r);}));
  return [...selected.values()];
}
function coReadEntries(context,kind,selected,indexMaps) {
  const excluded=new Set(selected.map(([type,id])=>key(type,id))),groups=new Map();
  context.forEach(r=>{
    const type=kind==='object'?'object':r.schema?'schema':'type',id=kind==='object'?r.objectKey:r.schema||r.SourceType,token=key(type,id);
    if(excluded.has(token))return;
    if(!groups.has(token))groups.set(token,{item:indexMaps[type].get(id),models:new Set()});groups.get(token).models.add(r.modelKey);
  });
  return [...groups.values()].sort((a,b)=>b.models.size-a.models.size||compare(a.item.label,b.item.label));
}
const DATA = await loadModelSources();
if (!DATA || !Array.isArray(DATA.rows) || !DATA.files || typeof DATA.files !== 'object') throw new Error('The inventory has an invalid structure. Expected rows and files. Regenerate the report.');
const INDEX = buildIndex(DATA);
const {rows, maps} = INDEX;
// Assign category colors once, in descending source-row frequency, across the inventory.
const sourceGroupKey = r => key(r.schema ? 'schema' : 'type', r.schema || r.SourceType);
const sourceGroups = new Map();
rows.forEach(r => {
  const id = sourceGroupKey(r);
  if (!sourceGroups.has(id)) sourceGroups.set(id, {id,type:r.schema?'schema':'type',target:r.schema||r.SourceType,label:r.schemaLabel||r.SourceType||'(unspecified)',count:0});
  sourceGroups.get(id).count++;
});
const rankedGroups = [...sourceGroups.values()].sort((a,b)=>b.count-a.count||compare(a.id,b.id));
rankedGroups.forEach((group,index)=>{group.rank=index;group.color=`hsl(${(174+index*137.508)%360} ${index%2?36:43}% ${48+(index%3)*4}%)`;});
let lineageStyle = 'cards';
try { if (localStorage.getItem('source-atlas-lineage-style') === 'flow') lineageStyle = 'flow'; } catch { /* Optional file:// preference. */ }
const color = name => { let hash = 0; for (const c of name) hash = (hash * 31 + c.charCodeAt(0)) | 0; return `hsl(${(Math.abs(hash) * 137.508) % 360} 49% 54%)`; };
let state = {type:'home',id:'',withItems:[]};
let filters = {workspace:'',type:'',review:false};
let currentRows = rows;
let graphCleanup = () => {};
let toastTimer;
function toast(message) { $('toast').textContent = message; $('toast').hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => $('toast').hidden = true, 2600); }
async function copyText(text) {
  try { if (!navigator.clipboard) throw Error('Fallback'); await navigator.clipboard.writeText(text); toast('Copied to clipboard'); }
  catch {
    const input = el('textarea'); input.value = text; input.style.position = 'fixed'; input.style.opacity = '0';
    const parent = $('dialog').open ? $('dialog') : document.body; parent.append(input); input.select();
    const ok = document.execCommand('copy'); input.remove(); toast(ok ? 'Copied to clipboard' : 'Copy unavailable. Select and copy the text manually.');
  }
}
function navigate(type, id = '', withItems = []) {
  const hash = encodeRoute({type,id,withItems},maps);
  if (location.hash === hash) render(); else location.hash = hash;
  closeSearch();
}
function readRoute() {
  const previous=state;state=decodeRoute(location.hash,maps);const id=state.id;
  // Accept the brief's readable workspace|model example as well as lossless JSON identifiers.
  if (state.type === 'model' && !maps.model.has(id) && id.includes('|')) { const pos = id.indexOf('|'); state.id = key(id.slice(0,pos),id.slice(pos+1)); }
  const canonicalHash=encodeRoute(state,maps);
  if(location.hash!==canonicalHash){try{history.replaceState(null,'',canonicalHash);}catch{/* Canonical ids still apply when history updates are unavailable. */}}
  if(previous.type!==state.type||previous.id!==state.id)filters = {workspace:'',type:'',review:false};
  render();
  // After a change of page, the keyboard / screen reader focus moves to the new page title (not on filter changes).
  const heading=$('main').querySelector('h1');if(heading){heading.tabIndex=-1;heading.focus({preventScroll:true});}
}
function matchesFilter(r) { return (!filters.workspace || r.Workspace === filters.workspace) && (!filters.type || r.SourceType === filters.type) && (!filters.review || r.NeedsReview === 'yes'); }
function itemFor() { return maps[state.type === 'lineage' ? 'model' : state.type]?.get(state.id); }
function selectedRows() {
  if (itemFor()) return ['schema','object','server','database','type'].includes(state.type)?intersectSources(maps,state,state.withItems||[],matchesFilter):itemFor().rows.filter(matchesFilter);
  if (state.id) return [];
  return rows.filter(r => matchesFilter(r) && (state.type !== 'review' || r.NeedsReview === 'yes'));
}
function panel(title, subtitle, action) {
  const root = el('section','panel'); const head = el('div','panel-head'); const titleBox = el('div');
  append(titleBox,el('h2','',title)); if (subtitle) titleBox.append(el('p','',subtitle)); head.append(titleBox); if (action) head.append(action);
  root.append(head); return root;
}
function paged(parent, items, create, size = 25) {
  let count = 0; const more = button('', add, 'load');
  function add() {
    more.remove(); const chunk = document.createDocumentFragment();
    items.slice(count,count+size).forEach((item,index) => chunk.append(create(item,count+index)));
    count += size; parent.append(chunk);
    if (count < items.length) { more.textContent = `Show ${Math.min(size,items.length-count)} more · ${fmt(items.length-count)} remaining`; parent.append(more); }
  }
  if (!items.length) parent.append(el('div','empty','No results for this selection.')); else add();
}
function rankPanel(title, items, size = 6, subtitle = 'Distinct models using each source') {
  const root = panel(title,subtitle); const body = el('div','panel-body'); root.append(body);
  const sorted = [...items].sort((a,b) => b.models.size-a.models.size || compare(a.label,b.label)); const max = sorted[0]?.models.size || 1;
  paged(body, sorted, (item,index) => {
    const row = button('',() => navigate(item.type,item.id),'rank-row'); const name = el('div','rank-name');
    append(name,el('strong','',item.label),el('small','',item.type === 'object' ? item.detail : item.rows.length+' source rows'));
    const track = el('div','bar-track'); const fill = el('div','bar-fill'); fill.style.width = (100*item.models.size/max)+'%';
    if (item.type === 'workspace') fill.style.background = color(item.label);
    track.append(fill); name.append(track); append(row,el('span','rank-num',String(index+1).padStart(2,'0')),name,el('span','rank-value',fmt(item.models.size)+' models')); return row;
  },size); return root;
}
function renderFilters(parent) {
  const root = el('div','filters');
  function select(label, field, source) {
    const wrap = el('label','',label); const input = el('select'); input.setAttribute('aria-label',label);
    const all = el('option','',field === 'workspace' ? 'All workspaces' : 'All source types'); all.value = ''; input.append(all);
    [...source.values()].sort((a,b)=>compare(a.label,b.label)).forEach(item => { const opt = el('option','',item.label); opt.value = item.id; input.append(opt); });
    input.value = filters[field]; input.addEventListener('change',()=>{filters[field]=input.value;render();}); wrap.append(input); root.append(wrap);
  }
  select('Workspace','workspace',maps.workspace); select('Source','type',maps.type);
  const label = el('label'); const input = el('input'); input.type = 'checkbox'; input.checked = filters.review;
  input.addEventListener('change',()=>{filters.review=input.checked;render();}); append(label,input,document.createTextNode('Needs review only')); root.append(label);
  if (filters.workspace || filters.type || filters.review) root.append(button('Reset',()=>{filters={workspace:'',type:'',review:false};render();},'link'));
  root.append(el('span','explain','Filters qualify subject rows; co-read filters use complete model sources.')); parent.append(root);
}
function exportCSV() {
  const fields = ['Workspace','SemanticModel','Table','SourceType','Server','Database','Object','ObjectType','Mode','Via','Parameters','NeedsReview','QueryFile','SqlFile'];
  // Quote every cell and neutralize spreadsheet formula prefixes in untrusted names.
  const cell = value => { let s = String(value ?? ''); if (/^[\s]*[=+@-]/.test(s) || /^[\t\r\n]/.test(s)) s = "'"+s; return '"'+s.replace(/"/g,'""')+'"'; };
  const text = '\uFEFF'+[fields.map(cell).join(','),...currentRows.map(r=>fields.map(f=>cell(r[f])).join(','))].join('\r\n');
  const url = URL.createObjectURL(new Blob([text],{type:'text/csv;charset=utf-8'})); const a = el('a'); a.href = url; a.download = 'source-atlas-'+state.type+'.csv'; document.body.append(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1000); toast(`Exported ${fmt(currentRows.length)} source rows`);
}
function copyModels() { const ids = [...new Set(currentRows.map(r=>r.modelKey))]; copyText(ids.map(id=>{const r=maps.model.get(id).rows[0];return r.Workspace+' / '+r.SemanticModel;}).sort(compare).join('\n')); }
function render() {
  graphCleanup(); graphCleanup = () => {}; const main = $('main'); main.replaceChildren(); currentRows = selectedRows();
  const item = itemFor(); const isImpact = ['schema','object','server','database','type'].includes(state.type);
  const titles = {home:'Your data, connected.',models:'Semantic models',sources:'Source catalog',lineage:'Lineage explorer',review:'Sources to review'};
  const subtitles = {home:'Explore where your models begin — and understand what a change could affect.',models:'Explore tables, source dependencies, and the code behind every model.',sources:'Find a source to trace its reach across your semantic models.',lineage:'Follow the path from source infrastructure to semantic model.',review:'Uncertain source references that need a closer look.'};
  const breadcrumb = el('div','breadcrumb'); append(breadcrumb,button('Inventory',()=>navigate('home'),'link'),document.createTextNode('  /  '+(item ? state.type.charAt(0).toUpperCase()+state.type.slice(1) : titles[state.type] || 'Selection'))); main.append(breadcrumb);
  const heading = el('div','page-heading'); const title = el('div');
  if (item) title.append(el('span','eyebrow',isImpact ? 'Impact analysis · '+state.type : state.type));
  append(title,el('h1','',item?.label || titles[state.type] || 'Selection not found'),el('p','',item ? item.detail : subtitles[state.type] || 'Use search or return to the overview.'));
  const actions = el('div','heading-actions'); append(actions,button('Copy model names',copyModels),button('↓ Export CSV',exportCSV)); append(heading,title,actions); main.append(heading);
  document.querySelectorAll('nav button').forEach(b=>{ const active=b.dataset.view === state.type || b.dataset.view === 'models' && ['model','table','workspace'].includes(state.type) || b.dataset.view === 'sources' && isImpact; b.classList.toggle('active',active); if(active)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current'); });
  if (state.type === 'home') renderHome(main);
  else {
    renderFilters(main);
    if (state.id && !item) main.append(el('div','empty','This selection is not in the inventory.'));
    else if (isImpact) renderImpact(main,item);
    else if (state.type === 'model' || state.type === 'table') renderModel(main,item);
    else if (state.type === 'review') renderRows(main,currentRows,'Review queue');
    else if (state.type === 'sources') renderCatalog(main);
    else if (state.type === 'lineage') renderLineageExplorer(main);
    else if (state.type === 'models' || state.type === 'workspace') renderModels(main,currentRows);
    else main.append(el('div','empty','Choose a view from the navigation.'));
  }
  const foot = el('footer');const credit=el('span');const creditLink=el('a','','lazysnail.net');creditLink.href='https://lazysnail.net';creditLink.target='_blank';creditLink.rel='noopener';credit.append(creditLink);append(foot,el('span','',`Generated ${DATA?.generatedAtUtc ? new Date(DATA.generatedAtUtc).toLocaleString() : '—'} · ${DATA?.generator || 'Unknown generator'}`),credit); main.append(foot);
}
function renderHome(main) {
  const stats = el('div','stats');
  const values = [['Semantic models',maps.model.size,'models','Explore models →'],['Model tables',maps.table.size,'models','Inspect dependencies →'],['Source rows',rows.length,'sources','Browse inventory →'],['Distinct sources',maps.object.size,'sources','Trace an impact →'],['Needs review',rows.filter(r=>r.NeedsReview==='yes').length,'review','Open review queue →']];
  values.forEach(([label,value,type,hint])=>{ const b=button('',()=>navigate(type),'stat');append(b,el('span','label',label),el('span','value',fmt(value)),el('span','hint',hint));stats.append(b); }); main.append(stats);
  const top = [...maps.schema.values()].sort((a,b)=>b.models.size-a.models.size)[0] || [...maps.model.values()][0];
  if (top) main.append(lineagePanel(top.rows,'Dependency spotlight',`A closer look at ${top.label} · ${fmt(top.models.size)} connected models`,button('Explore impact ↗',()=>navigate(top.type,top.id),'link')));
  const grid = el('div','grid-two'); append(grid,rankPanel('Most connected schemas',maps.schema.values()),rankPanel('Most connected source objects',maps.object.values())); main.append(grid);
  const second = el('div','grid-two section-gap'); append(second,rankPanel('Models by workspace',maps.workspace.values(),6,'Workspace coverage across this inventory'),rankPanel('Source types',maps.type.values(),6,'Connectors and generated sources')); main.append(second);
}
function renderCatalog(main) {
  const tabs = el('div','flex wrap'); const content = el('div','section-gap');
  function show(type) {
    content.replaceChildren(); [...tabs.children].forEach(b=>b.classList.toggle('primary',b.dataset.type===type));
    const ids = new Set(currentRows.map(r=>({schema:r.schema,object:r.objectKey,server:r.serverKey,database:r.databaseKey,type:r.SourceType})[type]));
    const items = [...maps[type].values()].filter(i=>ids.has(i.id)).map(i=>({...i,rows:i.rows.filter(matchesFilter),models:new Set(i.rows.filter(matchesFilter).map(r=>r.modelKey))}));
    content.append(rankPanel('Browse '+({schema:'schemas',object:'source objects',server:'servers',database:'databases',type:'source types'})[type],items,30));
  }
  ['schema','object','server','database','type'].forEach(type=>{const b=button(({schema:'Schemas',object:'Objects',server:'Servers',database:'Databases',type:'Source types'})[type],()=>show(type));b.dataset.type=type;tabs.append(b);});
  append(main,tabs,content);show('schema');
}
function fullModelRows(matching) { return [...new Set(matching.map(r=>r.modelKey))].flatMap(id=>maps.model.get(id).rows); }
function renderImpact(main,item) {
  const impacted = [...new Set(currentRows.map(r=>r.modelKey))];
  const context = fullModelRows(currentRows);
  const selected=[[state.type,state.id],...(state.withItems||[])];
  const selection=el('div','context-banner wrap');selection.setAttribute('aria-label','Combined source selection');
  append(selection,el('span','eyebrow','Models reading all of'),el('span','pill teal',item?.label||state.id));
  (state.withItems||[]).forEach(([type,id],index)=>{
    const label=maps[type]?.get(id)?.label||id;
    const chip=button('+ '+label+' ×',()=>navigate(state.type,state.id,state.withItems.filter((_,i)=>i!==index)),'pill');chip.setAttribute('aria-label','Remove co-read filter '+label);selection.append(chip);
  });main.append(selection);
  const banner = el('div','context-banner'); append(banner,el('strong','',fmt(impacted.length)+' impacted models'),el('span','',`from ${fmt(currentRows.length)} matching source rows. Expand a model to see everything else it reads.`)); main.append(banner);
  const graphHost=el('div');main.append(graphHost);
  function showImpactGraph(allReads){
    graphCleanup();graphHost.replaceChildren();const scope=el('div','flex');
    append(scope,button('Matching paths',()=>showImpactGraph(false),allReads?'':'primary'),button('All model reads',()=>showImpactGraph(true),allReads?'primary':''));
    graphHost.append(lineagePanel(allReads?context:currentRows,'Source → model lineage',allReads?'Complete source context of impacted models · includes reads outside the filters':'Matching dependencies · select any node to refocus',scope));
  }
  showImpactGraph(false);
  const grid = el('div','grid-impact'); const list = panel('Impacted models',`${fmt(impacted.length)} models · grouped by workspace`);
  const matched = new Set(currentRows.map(r=>r.index));
  const models = impacted.map(id=>maps.model.get(id)).sort((a,b)=>compare(a.rows[0].Workspace,b.rows[0].Workspace)||compare(a.label,b.label));
  paged(list,models,m=>modelCard(m,matched),15);
  const side = el('div','stack'); side.style.alignContent='start';
  const coPanel = panel('Often used together','Click to require another source · shares of remaining models'); const body = el('div','panel-body'); coPanel.append(body);
  const controls = el('div','flex'); const ranking = el('div');
  function coRead(kind) {
    ranking.replaceChildren(); [...controls.children].forEach(b=>b.classList.toggle('primary',b.dataset.kind===kind));
    const entries = coReadEntries(context,kind,selected,maps);
    paged(ranking,entries,entry=>{
      const row=el('div','flex');const filterButton=button('',()=>navigate(state.type,state.id,[...(state.withItems||[]),[entry.item.type,entry.item.id]]),'rank-row');
      filterButton.title='Require models to also read '+entry.item.label;
      const name=el('div','rank-name');append(name,el('strong','',entry.item.label),el('small','',`${entry.models.size} of ${impacted.length} models`));append(filterButton,name,el('span','pill teal',Math.round(entry.models.size/impacted.length*100)+'%'));
      const open=button('↗',()=>navigate(entry.item.type,entry.item.id),'icon-btn');open.title='Open '+entry.item.label+' impact';open.setAttribute('aria-label',open.title);append(row,filterButton,open);return row;
    },10);
  }
  ['schema','object'].forEach(kind=>{const b=button(kind==='schema'?'Schemas / types':'Objects',()=>coRead(kind));b.dataset.kind=kind;controls.append(b);}); append(body,controls,ranking);coRead('schema');side.append(coPanel);append(grid,list,side);main.append(grid);
}
function modelCard(model,matched) {
  const root=el('article','model-card'); const r=model.rows[0];const top=el('div','flex between wrap');const name=el('div','flex');const dot=el('span','dot');dot.style.background=color(r.Workspace);append(name,dot,button(model.label,()=>navigate('model',model.id),'link model-title'));
  append(top,name,el('span','pill',r.Workspace)); root.append(top);
  const sources=[...new Map(model.rows.map(row=>[row.objectKey,row])).values()];const matching=matched ? model.rows.filter(row=>matched.has(row.index)) : model.rows;
  const other=sources.filter(row=>!matching.some(m=>m.objectKey===row.objectKey));
  root.append(el('small','',`${new Set(model.rows.map(row=>row.tableKey)).size} tables · ${sources.length} sources${matched ? ' · '+other.length+' other sources' : ''}`));
  const chips=el('div','chips');const groups=new Map();
  // Count every source row, including repeated references; selected categories come first.
  model.rows.forEach(row=>{const id=sourceGroupKey(row);if(!groups.has(id))groups.set(id,{...sourceGroups.get(id),rows:0,selected:false});const group=groups.get(id);group.rows++;if(matched?.has(row.index))group.selected=true;});
  [...groups.values()].sort((a,b)=>Number(b.selected)-Number(a.selected)||b.rows-a.rows||a.rank-b.rank).forEach(group=>{
    const chip=button('',()=>navigate(group.type,group.target),'source-chip'+(group.selected?' selected':''));chip.style.setProperty('--group-color',group.color);
    chip.title=(group.selected?'Includes the selected source · ':'')+group.rows+' source rows';
    append(chip,el('span','dot'),document.createTextNode(`${group.label} · ${group.rows}${group.selected?' ✓':''}`));chips.append(chip);
  });root.append(chips);
  const details=el('details');details.append(el('summary','',matched ? `Everything this model reads · ${sources.length} sources` : `Inspect all ${sources.length} sources`));
  details.addEventListener('toggle',()=>{if(!details.open||details.dataset.loaded)return;details.dataset.loaded='yes';const list=el('div','model-sources');paged(list,sources,row=>{
    const line=el('div','source-row');append(line,el('span',matching.some(m=>m.objectKey===row.objectKey)&&matched?'pill teal':'pill',matched ? matching.some(m=>m.objectKey===row.objectKey)?'Matching':'Also reads':'Source'));
    const source=el('div');append(source,button(row.Object||row.SourceType||'(unnamed)',()=>navigate('object',row.objectKey),'link'),el('small','',[row.Server,row.Database,row.schemaLabel||row.SourceType].filter(Boolean).join(' / ')));line.append(source);
    line.append(button('Inspect table →',()=>navigate('table',row.tableKey),'link'));return line;
  },25);details.append(list);});root.append(details);return root;
}
function renderModels(main,subset) {
  const models=[...new Set(subset.map(r=>r.modelKey))].map(id=>maps.model.get(id)).sort((a,b)=>compare(a.rows[0].Workspace,b.rows[0].Workspace)||compare(a.label,b.label));
  const root=panel('Model inventory',fmt(models.length)+' models · workspace order');paged(root,models,m=>modelCard(m),20);main.append(root);
}
function renderModel(main,item) {
  if(!item)return;
  main.append(lineagePanel(currentRows,'Model dependencies','All selected table-to-source paths · select a style to explore'));
  const tables=[...new Set(currentRows.map(r=>r.tableKey))].map(id=>maps.table.get(id));
  const root=panel(state.type==='table'?'Table sources & query':'Tables & sources',`${tables.length} tables · click a table to inspect its query code`);
  const wrap=el('div','table-wrap'); const table=el('table');const head=el('thead');const hr=el('tr');['Model table','Source objects','Storage mode','Status'].forEach(t=>hr.append(el('th','',t)));head.append(hr);table.append(head);const body=el('tbody');table.append(body);wrap.append(table);root.append(wrap);
  let shown=0;const more=button('Show more tables',add,'load');
  function add(){more.remove();tables.slice(shown,shown+35).forEach(t=>{
    const tr=el('tr');const name=el('td');name.append(button(t.label,()=>openQuery(t),'link'));name.append(el('small','',new Set(t.rows.map(r=>r.objectKey)).size+' distinct sources'));
    const sources=el('td');const unique=[...new Map(t.rows.map(r=>[r.objectKey,r])).values()];
    const sourceList=el('div');paged(sourceList,unique,r=>{const d=el('div');d.append(button(r.Object||r.SourceType||'(unnamed)',()=>navigate('object',r.objectKey),'link'));d.append(el('small','',[r.Server,r.Database].filter(Boolean).join(' / ')));return d;},8);sources.append(sourceList);
    const modes=el('td','',[...new Set(t.rows.map(r=>r.Mode||'unspecified'))].join(', '));const status=el('td');status.append(el('span',t.rows.some(r=>r.NeedsReview==='yes')?'pill warn':'pill teal',t.rows.some(r=>r.NeedsReview==='yes')?'Needs review':'Resolved'));
    append(tr,name,sources,modes,status);body.append(tr);
  });shown+=35;if(shown<tables.length)root.append(more);}
  add();if(!tables.length)root.append(el('div','empty','No tables match the filters.'));main.append(root);
  if(state.type==='table')main.append(button('Open query code',()=>openQuery(item),'primary section-gap'));
}
function renderRows(main,subset,title) {
  const root=panel(title,fmt(subset.length)+' source rows · select a table to inspect the extraction');
  paged(root,subset,r=>{const article=el('article','model-card');const top=el('div','flex between wrap');append(top,button(r.Table,()=>openQuery(maps.table.get(r.tableKey)),'link model-title'),el('span','pill warn','Needs review'));article.append(top);
    const path=el('div','flex wrap');append(path,button(r.SemanticModel,()=>navigate('model',r.modelKey),'link'),el('small','',r.Workspace));article.append(path);
    append(article,el('p','',r.Object||'(source object unresolved)'),el('small','',[r.SourceType,r.Server,r.Database,r.Parameters,r.Via].filter(Boolean).join(' · ')));return article;
  },30);main.append(root);
}
function openDialog(title,kicker) { $('dialog-title').textContent=title;$('dialog-kicker').textContent=kicker;$('dialog-body').replaceChildren();if(!$('dialog').open)$('dialog').showModal();return $('dialog-body'); }
function openQuery(table) {
  const body=openDialog(table.label,'Query inspector · '+table.rows[0].SemanticModel);
  const paths=[...new Set(table.rows.flatMap(r=>[r.QueryFile,r.SqlFile]).filter(Boolean))];
  const metadata=el('p','muted',table.rows.map(r=>[r.SourceType,r.Mode,r.Via,r.Parameters].filter(Boolean).join(' · ')).filter((v,i,a)=>a.indexOf(v)===i).join('\n'));body.append(metadata);
  if(!paths.length){body.append(el('div','empty','No query file is associated with this table.'));return;}
  const tabs=el('div','code-tabs');const toolbar=el('div','flex between wrap');const pathLabel=el('small');const code=el('div','code-block');code.tabIndex=0;code.setAttribute('aria-label','Query source code');const next=el('div');let activeText='';
  append(toolbar,pathLabel,button('Copy code',()=>copyText(activeText)));append(body,tabs,toolbar,code,next);
  function show(path){
    [...tabs.children].forEach(b=>b.classList.toggle('active',b.dataset.path===path));pathLabel.textContent=path;
    activeText=Object.hasOwn(DATA.files||{},path)?String(DATA.files[path]):'';code.replaceChildren();next.replaceChildren();
    if(!Object.hasOwn(DATA.files||{},path)){code.append(el('div','empty','This file is not included in the inventory.'));return;}
    // Tokenize lazily, only when opened. Retain multiline comments/strings across line breaks.
    const tokenPattern=/(\/\*[\s\S]*?\*\/|\/\/[^\n]*|--[^\n]*|"(?:""|[^"])*"|'(?:''|[^'])*'|\b(?:let|in|each|if|then|else|try|otherwise|true|false|null|and|or|not|select|from|where|join|left|right|inner|outer|on|as|group|by|order|having|union|all|case|when|end|distinct|with|return|var|calculate|filter|sum|sumx|date|calendar|table|values|related)\b)/gi;
    const lines=[[]];let offset=0;function segment(text,cls){const parts=text.split('\n');parts.forEach((part,i)=>{if(i)lines.push([]);if(part)lines[lines.length-1].push({text:part,cls});});}
    for(const match of activeText.matchAll(tokenPattern)){segment(activeText.slice(offset,match.index),'');const text=match[0];segment(text,/^(\/\/|--|\/\*)/.test(text)?'tok-comment':/^["']/.test(text)?'tok-string':'tok-key');offset=match.index+text.length;}segment(activeText.slice(offset),'');
    let count=0;const more=button('Show next 300 lines',add,'load');function add(){more.remove();const frag=document.createDocumentFragment();lines.slice(count,count+300).forEach((parts,i)=>{const line=el('div','code-line');const num=el('span','line-no',count+i+1);num.setAttribute('aria-hidden','true');const content=el('span','line-code');parts.forEach(part=>content.append(el('span',part.cls,part.text)));append(line,num,content);frag.append(line);});count+=300;code.append(frag);if(count<lines.length)next.append(more);}add();
  }
  paths.forEach(path=>{const b=button(path.split(/[\\/]/).pop(),()=>show(path));b.dataset.path=path;tabs.append(b);});show(paths[0]);
}
// Graph construction keeps real row memberships on every node and edge. Hover highlights
// only paths represented by those rows, avoiding false transitive paths at shared nodes.
function graphData(subset,limit) {
  const lanes=Array.from({length:6},()=>new Map());const paths=[];
  subset.forEach(r=>{
    const specs=[
      [key(r.serverKey||'',r.Server?'':r.SourceType),r.Server||r.SourceType||'Unspecified source',r.Server?'server':'type',r.serverKey||r.SourceType],
      [r.databaseKey,r.Database||'No database',r.Database?'database':null,r.databaseKey],
      [key(r.serverKey,r.databaseKey,r.schema?'schema':'type',r.schema||r.SourceType),r.schemaLabel||r.SourceType||'Unspecified group',r.schema?'schema':'type',r.schema||r.SourceType],
      [r.objectKey,r.Object||'(unnamed source)','object',r.objectKey],
      [r.tableKey,r.Table,'table',r.tableKey],
      [r.modelKey,r.SemanticModel,'model',r.modelKey]
    ];
    const path=specs.map(([id,label,type,target],lane)=>{
      let node=lanes[lane].get(id);if(!node){node={id:lane+':'+id,label,type,target,lane,rows:new Set(),workspace:lane>=4?r.Workspace:'',detail:lane===4?r.SemanticModel:lane===5?r.Workspace:''};lanes[lane].set(id,node);}node.rows.add(r.index);return node;
    });paths.push({r,nodes:path});
  });
  const replacements=new Map();const columns=lanes.map((map,lane)=>{
    const sorted=[...map.values()].sort((a,b)=>b.rows.size-a.rows.size||compare(a.label,b.label));
    if(sorted.length<=limit)return sorted;
    const kept=sorted.slice(0,limit);const hidden=sorted.slice(limit);const aggregate={id:'aggregate:'+lane,label:`+ ${hidden.length} more ${['servers','databases','groups','objects','tables','models'][lane]}`,lane,rows:new Set(),hidden,workspace:'',detail:'Click to explore all members'};
    hidden.forEach(node=>{replacements.set(node.id,aggregate);node.rows.forEach(id=>aggregate.rows.add(id));});return [...kept,aggregate];
  });
  const edges=new Map();paths.forEach(({r,nodes})=>{
    const mapped=nodes.map(n=>replacements.get(n.id)||n);
    for(let i=0;i<5;i++){const a=mapped[i],b=mapped[i+1],id=key(a.id,b.id);let edge=edges.get(id);if(!edge){edge={a,b,rows:new Set(),workspaces:new Set()};edges.set(id,edge);}edge.rows.add(r.index);edge.workspaces.add(r.Workspace);}
  });
  // Alternating barycentric sweeps reduce crossings while aggregate nodes stay at the bottom.
  const edgeList=[...edges.values()];
  for(let pass=0;pass<4;pass++){
    const forward=pass%2===0;const order=forward?[1,2,3,4,5]:[4,3,2,1,0];
    order.forEach(lane=>{
      const adjacent=columns[lane+(forward?-1:1)];const rank=new Map(adjacent.map((node,i)=>[node.id,i]));
      const scores=new Map();columns[lane].forEach(node=>{let total=0,weight=0;edgeList.forEach(edge=>{const near=forward?edge.a:edge.b;const far=forward?edge.b:edge.a;if(far===node&&rank.has(near.id)){total+=rank.get(near.id)*edge.rows.size;weight+=edge.rows.size;}});scores.set(node.id,weight?total/weight:0);});
      columns[lane].sort((a,b)=>Number(!!a.hidden)-Number(!!b.hidden)||scores.get(a.id)-scores.get(b.id)||compare(a.label,b.label));
    });
  }
  const max=Math.max(1,...columns.map(c=>c.length));const height=Math.max(290,max*66+88);
  columns.forEach((column,lane)=>column.forEach((node,i)=>{node.x=25+lane*232;node.y=62+(height-92-column.length*66)/2+i*66;}));
  return {columns,edges:edgeList,width:1390,height,total:lanes.reduce((n,lane)=>n+lane.size,0),visible:columns.reduce((n,lane)=>n+lane.length,0)};
}
// A single pixels-per-row scale preserves mass through all six lanes. Each endpoint
// has independently stacked category/neighbor slots, so bands never overlap at bars.
function layoutFlow(graph, subset) {
  const byIndex=new Map(subset.map(r=>[r.index,r]));
  const massHeight=360, gap=42, scale=massHeight/Math.max(1,subset.length);
  const maxNodes=Math.max(1,...graph.columns.map(column=>column.length));
  graph.height=100+massHeight+(maxNodes-1)*gap;
  const bands=[];
  graph.columns.forEach(column=>{
    let y=62+(maxNodes-column.length)*gap/2;
    column.forEach(node=>{
      node.y=y;node.h=node.rows.size*scale;node.segments=[];
      const counts=new Map();node.rows.forEach(id=>{const category=sourceGroupKey(byIndex.get(id));counts.set(category,(counts.get(category)||0)+1);});
      let offset=0;[...counts].sort((a,b)=>sourceGroups.get(a[0]).rank-sourceGroups.get(b[0]).rank).forEach(([category,count])=>{
        node.segments.push({category,y:offset,height:count*scale});offset+=count*scale;
      });
      y+=node.h+gap;
    });
  });
  graph.edges.forEach(edge=>{
    const categories=new Map();edge.rows.forEach(id=>{const category=sourceGroupKey(byIndex.get(id));if(!categories.has(category))categories.set(category,new Set());categories.get(category).add(id);});
    categories.forEach((members,category)=>bands.push({a:edge.a,b:edge.b,rows:members,category,width:members.size*scale}));
  });
  graph.columns.flat().forEach(node=>{
    for(const outgoing of [true,false]){
      const attached=bands.filter(band=>(outgoing?band.a:band.b)===node);
      attached.sort((a,b)=>sourceGroups.get(a.category).rank-sourceGroups.get(b.category).rank||(outgoing?a.b.y-b.b.y:a.a.y-b.a.y));
      let offset=0;attached.forEach(band=>{band[outgoing?'sourceY':'targetY']=node.y+offset;offset+=band.width;});
    }
  });
  return bands;
}
function svgEl(tag,attrs={},text) { const node=document.createElementNS('http://www.w3.org/2000/svg',tag);Object.entries(attrs).forEach(([k,v])=>node.setAttribute(k,String(v)));if(text!==undefined)node.textContent=text;return node; }
function lineagePanel(subset,title,subtitle,extra) {
  const root=panel(title,subtitle);root.classList.add('lineage-panel');const head=root.firstChild;
  const tools=el('div','diagram-toolbar');if(extra)tools.append(extra);const density=el('select');density.setAttribute('aria-label','Diagram detail');[['6','Compact'],['10','More detail']].forEach(([value,label])=>{const o=el('option','',label);o.value=value;density.append(o);});tools.append(density);
  const styleSwitch=el('div','style-switch');styleSwitch.setAttribute('role','group');styleSwitch.setAttribute('aria-label','Lineage drawing style');
  ['cards','flow'].forEach(style=>{const control=button(style==='cards'?'Cards':'Flow',()=>{lineageStyle=style;try{localStorage.setItem('source-atlas-lineage-style',style);}catch{/* Preference storage is optional. */}draw();});control.dataset.style=style;styleSwitch.append(control);});tools.prepend(styleSwitch);
  const minus=button('−',()=>zoom(.8),'icon-btn'),plus=button('+',()=>zoom(1.25),'icon-btn'),fitButton=button('Fit',()=>fit(),'icon-btn');minus.setAttribute('aria-label','Zoom out');plus.setAttribute('aria-label','Zoom in');append(tools,minus,plus,fitButton);head.append(tools);
  const shell=el('div','diagram-shell');const svg=svgEl('svg',{class:'diagram',role:'group','aria-label':'Interactive lineage. Tab to nodes, Enter to refocus. Drag to pan; zoom with the controls or Ctrl + scroll.'});shell.append(svg);root.append(shell);
  const foot=el('div','diagram-footer');const legend=el('div','legend');const hint=el('span','','Hover to trace · Click to explore · Drag to pan · Ctrl + scroll to zoom');append(foot,legend,hint);root.append(foot);const note=el('div','diagram-note');root.append(note);
  function drawLegend(){
    legend.replaceChildren();
    const entries=lineageStyle==='flow'?[...new Set(subset.map(sourceGroupKey))].map(id=>sourceGroups.get(id)).sort((a,b)=>a.rank-b.rank):[...new Set(subset.map(r=>r.Workspace))].sort(compare).map(name=>({label:name||'(unspecified workspace)',color:color(name)}));
    const label=lineageStyle==='flow'?'schema / source-type':'workspace';
    const host=entries.length>5?el('details'):legend;
    if(entries.length>5){host.append(el('summary','',`${entries.length} ${label} colors`));legend.append(host);}else legend.append(el('span','',label+':'));
    entries.forEach(entry=>{const item=el('span','flex');item.style.gap='5px';const dot=el('span','dot');dot.style.background=entry.color;append(item,dot,document.createTextNode(entry.label));host.append(item);});
  }
  let flowAnimated=false;
  let graph,scene,v={x:0,y:0,w:1390,h:500};let dragging=null,moved=false;
  function view(){svg.setAttribute('viewBox',`${v.x} ${v.y} ${v.w} ${v.h}`);}
  function fit(){if(!graph)return;const ratio=(svg.clientWidth||1100)/(svg.clientHeight||455);const w=Math.max(graph.width,graph.height*ratio);const h=w/ratio;v={x:(graph.width-w)/2,y:(graph.height-h)/2,w,h};view();}
  function zoom(factor,px=.5,py=.5){const next=Math.max(220,Math.min(5000,v.w/factor));const scale=next/v.w;v.x+=v.w*px*(1-scale);v.y+=v.h*py*(1-scale);v.w=next;v.h*=scale;view();}
  function focusNode(node){
    if(moved)return;
    if(node.hidden){const body=openDialog(node.label,'Aggregated lineage nodes');paged(body,node.hidden,n=>{const b=button('',()=>{if(n.type){$('dialog').close();navigate(n.type,n.target);}else toast('This is a placeholder for sources without a database.');},'rank-row');append(b,el('span','rank-name',n.label),el('span','rank-value',n.rows.size+' rows'));return b;},30);}
    else if(node.type)navigate(node.type,node.target);
    else toast('This source has no database in the inventory.');
  }
  function draw(){
    [...styleSwitch.children].forEach(control=>control.setAttribute('aria-pressed',String(control.dataset.style===lineageStyle)));drawLegend();
    const flow=lineageStyle==='flow';
    svg.replaceChildren();graph=graphData(subset,Number(density.value));const bands=flow?layoutFlow(graph,subset):[];scene=svgEl('g',{class:'diagram-animate'});svg.append(scene);
    if(!subset.length){scene.append(svgEl('text',{x:500,y:160,'font-size':18},'No dependencies match this selection.'));fit();return;}
    ['SERVER','DATABASE','SCHEMA / TYPE','SOURCE OBJECT','MODEL TABLE','SEMANTIC MODEL'].forEach((label,i)=>{
      scene.append(svgEl('text',{x:25+i*232,y:27,class:'lane-label'},String(i+1).padStart(2,'0')+'  '+label));
      scene.append(svgEl('line',{x1:25+i*232,y1:39,x2:215+i*232,y2:39,stroke:'var(--line)'}));
    });
    const members=[];
    if(flow){
      const layer=svgEl('g',{class:flowAnimated?'':'flow-reveal'});flowAnimated=true;scene.append(layer);
      bands.forEach(band=>{
        const x1=band.a.x+12,x2=band.b.x,c=x1+(x2-x1)/2,y1=band.sourceY,y2=band.targetY,w=band.width;
        const path=svgEl('path',{class:'flow-band',fill:sourceGroups.get(band.category).color,d:`M${x1},${y1} C${c},${y1} ${c},${y2} ${x2},${y2} L${x2},${y2+w} C${c},${y2+w} ${c},${y1+w} ${x1},${y1+w} Z`});
        path.append(svgEl('title',{},sourceGroups.get(band.category).label+' · '+band.rows.size+' rows'));layer.append(path);members.push({element:path,rows:band.rows,edge:true});
      });
    }else graph.edges.forEach(edge=>{const x1=edge.a.x+188,y1=edge.a.y+23,x2=edge.b.x,y2=edge.b.y+23;const c=(x1+x2)/2;const path=svgEl('path',{d:`M${x1},${y1} C${c},${y1} ${c},${y2} ${x2},${y2}`,class:'edge',stroke:edge.workspaces.size===1?color([...edge.workspaces][0]):'var(--accent)','stroke-width':Math.min(10,1+Math.log2(edge.rows.size+1)*1.3)});scene.append(path);members.push({element:path,rows:edge.rows,edge:true});});
    // Row -> drawn members, built once per drawing: a hover marks the members of the hovered rows, instead of testing
    // every member against every row.
    let byRow=null;
    function highlight(active){
      if(!active){members.forEach(member=>{member.element.classList.remove('dim');if(member.edge)member.element.classList.remove('hot');});return;}
      if(!byRow){byRow=new Map();members.forEach((member,i)=>member.rows.forEach(id=>{let list=byRow.get(id);if(!list){list=[];byRow.set(id,list);}list.push(i);}));}
      const hit=new Uint8Array(members.length);
      for(const id of active){const list=byRow.get(id);if(list)for(const i of list)hit[i]=1;}
      members.forEach((member,i)=>{member.element.classList.toggle('dim',!hit[i]);if(member.edge)member.element.classList.toggle('hot',hit[i]===1);});
    }
    graph.columns.flat().forEach(node=>{
      const group=svgEl('g',{class:'node'+(flow?' flow-node':'')+(node.hidden?' aggregate':''),transform:`translate(${node.x},${node.y})`,tabindex:0,role:'button','aria-label':node.label+' · '+node.rows.size+' source rows'});
      group.append(svgEl('title',{},node.label+'\n'+node.detail+'\n'+node.rows.size+' source rows'));
      if(flow){
        group.append(svgEl('rect',{class:'flow-hit',x:-4,y:Math.min(0,node.h/2-17),width:205,height:Math.max(34,node.h),rx:3,fill:'transparent'}));
        node.segments.forEach(segment=>group.append(svgEl('rect',{width:12,y:segment.y,height:segment.height,fill:sourceGroups.get(segment.category).color})));
        if(node.hidden)group.append(svgEl('rect',{x:-2,y:-2,width:16,height:node.h+4,rx:2,fill:'none',stroke:'var(--muted)','stroke-dasharray':'3 3'}));
      }else append(group,svgEl('rect',{class:'main',width:188,height:46,rx:8}),svgEl('rect',{x:0,y:10,width:3,height:26,rx:1.5,fill:node.workspace?color(node.workspace):'var(--accent)'}));
      const label=node.label.length>24?node.label.slice(0,23)+'…':node.label;group.append(svgEl('text',{x:flow?21:12,y:flow?node.h/2-3:19,'font-size':11,'font-weight':550},label));
      const sub=flow?`${node.rows.size} rows${node.detail?' · '+node.detail:''}`:node.detail||`${node.rows.size} source ${node.rows.size===1?'row':'rows'}`;group.append(svgEl('text',{x:flow?21:12,y:flow?node.h/2+12:34,class:'node-sub'},sub.length>29?sub.slice(0,28)+'…':sub));
      group.addEventListener('pointerenter',()=>highlight(node.rows));group.addEventListener('pointerleave',()=>highlight(null));group.addEventListener('focus',()=>highlight(node.rows));group.addEventListener('blur',()=>highlight(null));group.addEventListener('click',()=>focusNode(node));group.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();moved=false;focusNode(node);}});
      scene.append(group);members.push({element:group,rows:node.rows});
    });
    note.textContent=`${fmt(subset.length)} source rows · ${fmt(graph.total)} underlying nodes · ${graph.visible} drawn${graph.total>graph.visible?' · Dashed nodes collect remaining dependencies; click to inspect them.':''} · ${flow?'Bar heights and band widths are proportional to source rows. Colors identify schemas / source types.':'Link width reflects row count. Colors identify workspaces.'} Full labels on hover.`;fit();
  }
  // Ctrl / ⌘ + wheel zooms the diagram; a plain wheel keeps scrolling the page.
  svg.addEventListener('wheel',event=>{if(!event.ctrlKey&&!event.metaKey)return;event.preventDefault();const box=svg.getBoundingClientRect();zoom(event.deltaY<0?1.12:1/1.12,(event.clientX-box.left)/box.width,(event.clientY-box.top)/box.height);},{passive:false});
  svg.addEventListener('pointerdown',event=>{if(event.button!==0)return;moved=false;dragging={x:event.clientX,y:event.clientY,vx:v.x,vy:v.y};});
  svg.addEventListener('pointermove',event=>{if(!dragging)return;const dx=event.clientX-dragging.x,dy=event.clientY-dragging.y;if(Math.abs(dx)+Math.abs(dy)>4){moved=true;svg.setPointerCapture(event.pointerId);}if(moved){v.x=dragging.vx-dx*v.w/svg.clientWidth;v.y=dragging.vy-dy*v.h/svg.clientHeight;view();}});
  const release=()=>{dragging=null;};window.addEventListener('pointerup',release);svg.addEventListener('pointercancel',release);
  density.addEventListener('change',draw);draw();const observer=new ResizeObserver(()=>fit());observer.observe(svg);graphCleanup=()=>{observer.disconnect();window.removeEventListener('pointerup',release);};return root;
}
function renderLineageExplorer(main) {
  const controls=el('div','context-banner');const label=el('label','','Start with a model ');const select=el('select');select.style.maxWidth='60vw';select.setAttribute('aria-label','Select a model for lineage');
  [...maps.model.values()].sort((a,b)=>compare(a.label,b.label)).forEach(m=>{const o=el('option','',m.label+' · '+m.rows[0].Workspace);o.value=m.id;select.append(o);});label.append(select);controls.append(label);main.append(controls);
  if(state.id && maps.model.has(state.id)) select.value=state.id;
  const host=el('div');main.append(host);const all=button('Explore selected model',()=>navigate('model',select.value),'link');controls.append(all);
  function show(){graphCleanup();host.replaceChildren();const item=maps.model.get(select.value);currentRows=(item?.rows||[]).filter(matchesFilter);host.append(lineagePanel(currentRows,'Source → model','Use global search to start from a schema, object, server, or database.'));}
  select.addEventListener('change',()=>navigate('lineage',select.value));show();
}
// Debounced, bounded suggestions use a precomputed text index; query file contents are never searched.
let searchTimer,searchResults=[],searchIndex=-1;
function closeSearch(){clearTimeout(searchTimer);searchTimer=null;$('suggestions').hidden=true;$('search').setAttribute('aria-expanded','false');$('search').removeAttribute('aria-activedescendant');searchIndex=-1;}
function searchNow(){
  const q=$('search').value.toLowerCase().trim();const host=$('suggestions');host.replaceChildren();searchIndex=-1;$('search').removeAttribute('aria-activedescendant');
  if(!q){searchResults=[];closeSearch();return;}
  const found=INDEX.search.filter(item=>item.search.includes(q));
  found.sort((a,b)=>Number(b.label.toLowerCase()===q)-Number(a.label.toLowerCase()===q)||Number(b.label.toLowerCase().startsWith(q))-Number(a.label.toLowerCase().startsWith(q))||b.models.size-a.models.size);
  searchResults=found.slice(0,30);host.hidden=false;$('search').setAttribute('aria-expanded','true');
  if(!found.length){host.append(el('div','empty','No matching sources or models.'));return;}
  host.append(el('div','eyebrow',`${fmt(found.length)} matches · showing ${searchResults.length}`));
  searchResults.forEach((item,i)=>{const b=button('',()=>{navigate(item.type,item.id);$('search').value='';},'suggestion');b.id='suggestion-'+i;b.setAttribute('role','option');b.setAttribute('aria-selected','false');b.tabIndex=-1;
    const name=el('span');append(name,el('strong','',item.label),el('small','',item.detail));name.lastChild.style.display='block';append(b,el('span','pill teal',item.type),name);host.append(b);
  });
}
$('search').addEventListener('input',()=>{clearTimeout(searchTimer);searchTimer=setTimeout(()=>{searchTimer=null;searchNow();},140);});
$('search').addEventListener('focus',()=>{if($('search').value)searchNow();});
$('search').addEventListener('keydown',event=>{
  if(event.key==='Escape'){closeSearch();return;}
  if(event.key==='Enter'||event.key==='ArrowDown'||event.key==='ArrowUp'){clearTimeout(searchTimer);if(searchTimer){searchTimer=null;searchNow();}}
  if(event.key==='ArrowDown'||event.key==='ArrowUp'){event.preventDefault();if($('suggestions').hidden)searchNow();if(!searchResults.length)return;searchIndex=(searchIndex+(event.key==='ArrowDown'?1:-1)+searchResults.length)%searchResults.length;
    $('suggestions').querySelectorAll('[role=option]').forEach((b,i)=>{b.classList.toggle('active',i===searchIndex);b.setAttribute('aria-selected',String(i===searchIndex));});const active=$('suggestion-'+searchIndex);if(active){$('search').setAttribute('aria-activedescendant',active.id);active.scrollIntoView({block:'nearest'});}
  }else if(event.key==='Enter'&&!$('suggestions').hidden&&searchResults.length){event.preventDefault();clearTimeout(searchTimer);const item=searchResults[Math.max(0,searchIndex)];navigate(item.type,item.id);$('search').value='';}
});
document.addEventListener('pointerdown',event=>{if(!event.target.closest('.search-wrap'))closeSearch();});
document.addEventListener('keydown',event=>{if(event.key==='/'&&!$('dialog').open&&!['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName)){event.preventDefault();$('search').focus();}});
$('dialog-close').addEventListener('click',()=>$('dialog').close());
document.querySelectorAll('nav button').forEach(b=>b.addEventListener('click',()=>navigate(b.dataset.view)));
let theme='system';try{theme=localStorage.getItem('source-atlas-theme')||'system';}catch{/* Storage may be disabled for file URLs. */}
function applyTheme(){if(theme==='system')delete document.documentElement.dataset.theme;else document.documentElement.dataset.theme=theme;$('theme').title='Theme: '+theme+' · click to cycle';$('theme').setAttribute('aria-label','Theme: '+theme+'. Switch theme');}
$('theme').addEventListener('click',()=>{theme=({system:'light',light:'dark',dark:'system'})[theme]||'system';applyTheme();try{localStorage.setItem('source-atlas-theme',theme);}catch{/* Optional preference persistence. */}});
applyTheme();$('models-count').textContent=fmt(maps.model.size);$('review-count').textContent=fmt(rows.filter(r=>r.NeedsReview==='yes').length);
window.addEventListener('hashchange',readRoute);readRoute();
$('main').setAttribute('aria-busy','false');
})().catch(error => {
  const main=document.getElementById('main');main.replaceChildren();main.setAttribute('aria-busy','false');
  const box=document.createElement('section');box.className='loading-state';box.setAttribute('role','alert');
  const title=document.createElement('h2');title.textContent='Unable to open this inventory';
  const message=document.createElement('p');message.textContent=error.message || 'An unexpected error occurred. Regenerate the report and try again.';
  box.append(title,message);main.append(box);
});
</script>
</body>
</html>
'@

# ------------------------------------------------------------------ main
# The analysis runs under the invariant culture: case-insensitive regular expressions, sorting and date formats must not
# depend on the machine (under Turkish settings, (?i:TRIM) does not match "trim" and created certain, wrong objects).
# The CSV delimiter was taken from the user's settings above, before the switch; the user's culture is restored at the end.
$userCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [cultureinfo]::InvariantCulture
# The status line is stopped on every way out, Ctrl+C included.
try {
    Write-UiHeader 'Source analysis' 'offline'
    $ExportPath = [IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath))
    $OutputPath = [IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath))
    # Plain drive-letter paths only. \\?\C:\..., \\.\... and UNC spellings can name the same folder as C:\... while
    # defeating the path comparisons below (export inside output, scripts inside output, OneLake segments).
    foreach ($checked in @{ Name = '-ExportPath'; Value = $ExportPath }, @{ Name = '-OutputPath'; Value = $OutputPath }) {
        if ($checked.Value -cnotmatch '\A[A-Za-z]:\\') { throw "Refusing $($checked.Name) '$($checked.Value)': use a plain drive-letter path such as C:\Exports (no \\?\, \\.\ or UNC form)." }
    }
    if (-not (Test-Path -LiteralPath $ExportPath -PathType Container)) { throw "Export folder not found: $ExportPath (run Extract-SemanticModels.ps1 first, or point -ExportPath to the folder that holds your exported models)." }
    # OneLake File Explorer syncs local writes and deletes back to Fabric: never write there.
    function Assert-NotOneLake([string]$Path) {
        # Checked on -OutputPath and on every model folder derived from it (a workspace NAMED like the sync root would
        # otherwise land inside it). LIMIT: a subst drive, UNC share or 8.3 short name aliasing the sync folder is not detected.
        $full = [IO.Path]::GetFullPath($Path)
        foreach ($segment in $full.Split([char[]]@('\', '/'))) { if ($segment -like 'OneLake - *' -or $segment -eq 'OneLake') { throw "Refusing to write inside a OneLake File Explorer folder (local changes there are synced to Fabric): $full" } }
    }
    Assert-NotOneLake $OutputPath
    $exportRoot = $ExportPath.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $outputRoot = $OutputPath.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    # Neither folder may contain the other: the per-model output folders are deleted and rewritten on each run,
    # which must never reach the exported definitions this script reads.
    if ($outputRoot.StartsWith($exportRoot, [StringComparison]::OrdinalIgnoreCase)) { throw '-OutputPath must not be inside -ExportPath.' }
    if ($exportRoot.StartsWith($outputRoot, [StringComparison]::OrdinalIgnoreCase)) { throw '-ExportPath must not be inside -OutputPath.' }
    function Test-UnderOneDrive([string]$Path) {
        # True when $Path is, or is under, one of the OneDrive folders of the current user (personal or work / school).
        $full = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/') + '\'
        foreach ($root in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
            if ([string]::IsNullOrWhiteSpace($root) -or $root -cnotmatch '\A[A-Za-z]:\\') { continue }
            $prefix = [IO.Path]::GetFullPath($root).TrimEnd('\', '/') + '\'
            if ($prefix.Length -gt 3 -and $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
        return $false
    }
    function Assert-NotLinked($Item, [string]$Label) {
        # A junction or symbolic link would let an ordinary-looking path resolve somewhere else (into a synchronised folder,
        # into the folder that is deleted and rewritten...): refused. OneDrive marks the files and folders it syncs as
        # reparse points too (cloud files), without making them links: such an entry is accepted inside the OneDrive folder only.
        $linkType = if ($Item -is [IO.DirectoryInfo]) { [string]$Item.LinkType } else { '' }   # files: a hard link redirects nothing
        if ($linkType -or $null -ne $Item.LinkTarget) {
            $kind = if ($linkType) { $linkType } else { 'link' }
            throw "$Label, refused: '$($Item.FullName)' is a $kind$(if ($Item.LinkTarget) { " to '$($Item.LinkTarget)'" }). Use a plain local folder, without junction or symbolic link on the way."
        }
        if (($Item.Attributes -band [IO.FileAttributes]::ReparsePoint) -and -not (Test-UnderOneDrive $Item.FullName)) {
            throw "$Label, refused: '$($Item.FullName)' is a reparse point (a mount point, or a cloud / sync folder other than OneDrive). Use a plain local folder."
        }
    }
    # Folders already checked during the analysis of the current model (cleared for each model).
    $script:SafeFolders = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    function Assert-SafeDirectory([string]$Path) {
        # No existing ancestor may be a file or a link (see Assert-NotLinked).
        $cursor = $Path
        while ($cursor) {
            if ($script:SafeFolders.Contains($cursor)) { return }      # it and its ancestors were checked
            $entry = [IO.DirectoryInfo]::new($cursor)
            if ($entry.Exists) { Assert-NotLinked $entry 'Unsafe directory'; [void]$script:SafeFolders.Add($cursor) }
            elseif ([IO.File]::Exists($cursor)) { throw "Unsafe directory, refused: '$cursor' is a file where a folder is expected." }
            $cursor = [IO.Path]::GetDirectoryName($cursor)
        }
    }
    function Assert-OwnedOutputRoot([string]$Root, [bool]$Adopt) {
        # Work only inside a folder these scripts created, found empty, or that the user explicitly adopted:
        # it must never be let loose on a folder that holds anything else.
        $marker = Join-Path $Root '.semantic-model-tools'
        if (-not (Test-Path -LiteralPath $Root)) { [IO.Directory]::CreateDirectory($Root) | Out-Null }
        elseif (-not (Test-Path -LiteralPath $marker -PathType Leaf) -and -not $Adopt -and @(Get-ChildItem -LiteralPath $Root -Force | Select-Object -First 1).Count) {
            throw "Refusing to use '$Root': it already contains files and was not created by these scripts. Use a new or empty folder, or pass -AdoptOutputFolder once if you are certain it is a plain local folder that is not synchronised to OneLake / Fabric."
        }
        if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) { [IO.File]::WriteAllText($marker, "Folder managed by the semantic model export scripts. Its sub-folders are deleted and rewritten on each run.`r`n") }
    }
    Assert-SafeDirectory $OutputPath
    if ($PSScriptRoot -and ([IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar).StartsWith($outputRoot, [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing -OutputPath '$OutputPath': the scripts themselves are inside it. Keep the scripts outside the output folder." }
    # The output holds company code (M, SQL), server and database names: say so when it lands where it could be shared.
    if (Test-UnderOneDrive $OutputPath) { Write-UiWarning "$OutputPath is in your OneDrive folder: the inventory and the query files (SQL and M code, server names) will be synced to OneDrive." }
    $cursor = $OutputPath
    while ($cursor) {
        if (Test-Path -LiteralPath (Join-Path $cursor '.git')) {
            # The scripts' own repository ignores its data\ folder (see .gitignore): nothing to say for an output under
            # that folder. Anywhere else (another repository, or another folder of this one), warn.
            $own = $PSScriptRoot -and [string]::Equals([IO.Path]::GetFullPath($cursor).TrimEnd('\', '/'), [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase) -and
                ($OutputPath.TrimEnd('\', '/') + '\').StartsWith([IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'data')).TrimEnd('\', '/') + '\', [StringComparison]::OrdinalIgnoreCase)
            if (-not $own) { Write-UiWarning "$OutputPath is inside a git repository ($cursor). The inventory and query files hold your SQL and M code and your server names: keep them out of commits (add the folder to .gitignore)." }
            break
        }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }

    Set-UiStatus 'Looking for exported models'
    # Every <Model>.SemanticModel folder under -ExportPath, in a fixed order (ordinal, ignoring case).
    $allFolders = [IO.DirectoryInfo[]]@(Get-ChildItem -LiteralPath $ExportPath -Directory -Recurse -Filter '*.SemanticModel')
    $folderKeys = [string[]]@($allFolders | ForEach-Object { $_.FullName })
    [Array]::Sort([Array]$folderKeys, [Array]$allFolders, [Collections.IComparer][StringComparer]::OrdinalIgnoreCase)
    $describe = { param($Folder)
        $name = $Folder.Name.Substring(0, $Folder.Name.Length - '.SemanticModel'.Length)
        $platform = Join-Path $Folder.FullName '.platform'
        if (Test-Path -LiteralPath $platform) {
            try { $display = (Get-Content -LiteralPath $platform -Raw | ConvertFrom-Json).metadata.displayName; if ($display) { $name = [string]$display } } catch { Write-Verbose "Unreadable .platform in $($Folder.FullName)" }
        }
        $workspaceFolderName = if ($Folder.Parent.FullName.TrimEnd('\', '/') -ieq $ExportPath.TrimEnd('\', '/')) { '' } else { $Folder.Parent.Name }
        [pscustomobject]@{ Folder = $Folder; Name = $name; Workspace = $workspaceFolderName; Relative = ''
            Tmdl = Test-Path -LiteralPath (Join-Path $Folder.FullName 'definition\tables') -PathType Container
            Bim = Test-Path -LiteralPath (Join-Path $Folder.FullName 'model.bim') -PathType Leaf }
    }
    $allModels = @(foreach ($folder in $allFolders) { & $describe $folder })
    $modelFolders = @($allModels | Where-Object Tmdl)
    # Output folder names are given over EVERY model of the export, in the fixed order above, so that a filtered run
    # (-Workspace, -Model, -NewOnly...) gives each model the same folder as a full run: two models whose names give the
    # same folder name can never swap folders.
    $usedModelFolders = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $modelFolders) {
        $workspaceFolder = ConvertTo-FileName $entry.Workspace
        $leaf = Split-Path (Get-UniqueName (Join-Path $workspaceFolder (ConvertTo-FileName $entry.Name)) $usedModelFolders) -Leaf
        $entry.Relative = if ($entry.Workspace) { Join-Path $workspaceFolder $leaf } else { $leaf }
    }
    $isSelected = { param($Entry)
        if ($Workspace -and -not @($Workspace | Where-Object { Test-NameMatch $Entry.Workspace $_ }).Count) { return $false }
        if ($Model -and -not @($Model | Where-Object { Test-NameMatch $Entry.Name $_ }).Count) { return $false }
        if ($Exclude -and @($Exclude | Where-Object { Test-NameMatch $Entry.Workspace $_ }).Count) { return $false }
        return $true
    }
    # Models saved without TMDL (model.bim, the older TMSL format of Power BI projects) cannot be read: say how many, once.
    $notTmdl = @($allModels | Where-Object { -not $_.Tmdl } | Where-Object { & $isSelected $_ })
    if ($notTmdl.Count) {
        $bim = @($notTmdl | Where-Object Bim).Count
        $names = @($notTmdl | Select-Object -First 3 | ForEach-Object { if ($_.Workspace) { "$($_.Workspace)\$($_.Name)" } else { $_.Name } }) -join ', '
        Write-UiWarning "$(Format-UiCount $notTmdl.Count 'model folder') skipped ($names$(if ($notTmdl.Count -gt 3) { ', ...' })): no TMDL definition (definition\tables)$(if ($bim) { "; $bim saved as model.bim (TMSL format)" }). This tool reads TMDL only: models exported by Extract-SemanticModels.ps1 always are; in Power BI Desktop, turn on the preview feature 'Store semantic model using TMDL format' and save the project again."
    }
    if (-not $modelFolders.Count) { throw "No exported semantic model in TMDL format found under $ExportPath (expected <Model>.SemanticModel\definition\tables$(if ($notTmdl.Count) { "; $($notTmdl.Count) model folder(s) without TMDL were skipped" }))." }
    Assert-OwnedOutputRoot $OutputPath $AdoptOutputFolder.IsPresent
    if (-not $NoTSqlParser) { Set-UiStatus "Loading Microsoft's SQL parser"; Initialize-TSqlParser $ScriptDomPath }
    if ($null -ne $script:TSql) { Write-Verbose "SQL parser: Microsoft ScriptDom $($script:TSql.Version) ($($script:TSql.ParserType.Name)), $($script:TSql.Path)" }
    $csvPath = Join-Path $OutputPath 'model_sources.csv'
    # -NewOnly: the models already in the inventory (same key as $processed below) are left as they are.
    $alreadyParsed = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); $skippedNew = 0
    if ($NewOnly -and (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
        foreach ($row in @(Import-Csv -LiteralPath $csvPath -Delimiter $Delimiter[0])) { if ($row.PSObject.Properties['SemanticModel']) { [void]$alreadyParsed.Add("$($row.Workspace)|$($row.SemanticModel)") } }
    }
    $csvRows = [Collections.Generic.List[object]]::new()
    $processed = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $modelCount = 0; $tableTotal = 0
    # The models to analyse are selected first, so that the counters of the status line know their totals.
    $selectedModels = @(foreach ($entry in $modelFolders) {
            if (-not (& $isSelected $entry)) { continue }
            if ($NewOnly -and $alreadyParsed.Contains("$(Protect-CsvValue $entry.Workspace)|$(Protect-CsvValue $entry.Name)")) { $skippedNew++; continue }
            $entry
        })
    $tableFilesTotal = 0
    foreach ($entry in $selectedModels) { $tableFilesTotal += @([IO.Directory]::GetFiles((Join-Path $entry.Folder.FullName 'definition\tables'), '*.tmdl')).Count }
    $tablesDone = 0
    $nameWidth = [math]::Min(40, [math]::Max(12, [int](@($selectedModels | ForEach-Object { $_.Name.Length }) | Measure-Object -Maximum).Maximum))
    $shownWorkspace = $null
    $utf8 = [Text.UTF8Encoding]::new($false)
    foreach ($selectedModel in $selectedModels) {
        $folder = $selectedModel.Folder; $modelName = $selectedModel.Name; $workspaceName = $selectedModel.Workspace; $relativeModel = $selectedModel.Relative
        Set-UiStatus "Reading model $($modelCount + 1) of $($selectedModels.Count)" $tablesDone $tableFilesTotal $modelName
        $script:SafeFolders.Clear(); $script:MaskCache.Clear(); $script:MarkCache.Clear()   # per model: bounded memory
        # Input side too: a junction or link on the way to the exported model could make it the same physical folder
        # as the output folder that is deleted and rewritten below.
        # Checked down to the folder that is actually read: a junction BELOW the model folder (its 'definition' or 'tables'
        # sub-folder) could equally point into the output folder.
        Assert-SafeDirectory (Join-Path $folder.FullName 'definition\tables')
        $modelCount++
        [void]$processed.Add("$(Protect-CsvValue $workspaceName)|$(Protect-CsvValue $modelName)")
        $definition = Join-Path $folder.FullName 'definition'
        $info = [pscustomobject]@{ Expressions = [ordered]@{}; Queries = [ordered]@{}; Literals = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal); Cache = @{} }
        foreach ($file in 'expressions.tmdl', 'model.tmdl') { $path = Join-Path $definition $file; if (Test-Path -LiteralPath $path) { Read-TmdlExpressionsFile $path $info.Expressions } }
        $dataSources = Read-TmdlDataSources (Join-Path $definition 'dataSources.tmdl')
        $script:ModelParameters.Clear()
        foreach ($name in @($info.Expressions.Keys)) { if ($info.Expressions[$name] -match 'IsParameterQuery\s*=\s*true') { [void]$script:ModelParameters.Add($name) } }
        foreach ($name in @($info.Expressions.Keys)) {   # text / number parameters: "value" meta [...]
            $m = [regex]::Match($info.Expressions[$name], '\A\s*(?:"((?:[^"]|"")*)"|(-?\d+(?:\.\d+)?))\s*(?:meta\b[\s\S]*)?\z')
            if ($m.Success) { $info.Literals[$name] = if ($m.Groups[1].Success) { ConvertFrom-MString $m.Groups[1].Value } else { $m.Groups[2].Value } }
        }
        $modelOut = [IO.Path]::GetFullPath((Join-Path $OutputPath $relativeModel))
        if (-not ([IO.Path]::GetFullPath($modelOut) + '\').StartsWith($OutputPath.TrimEnd('\', '/') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing output folder outside -OutputPath: $modelOut" }
        Assert-NotOneLake $modelOut
        Assert-SafeDirectory $modelOut
        if (Test-Path -LiteralPath $modelOut) { Remove-Item -LiteralPath $modelOut -Recurse -Force }   # stale tables must not linger
        [IO.Directory]::CreateDirectory($modelOut) | Out-Null
        if ($info.Expressions.Count) {
            $usedShared = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($name in $info.Expressions.Keys) {
                $text = "// Semantic model : $modelName`n// Shared expression : $name`n`n$($info.Expressions[$name])`n"
                Write-ModelQueryFile (Join-Path '_shared' ((Get-UniqueName (ConvertTo-FileName $name) $usedShared) + '.pq')) $text
            }
        }
        $usedFiles = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $tableCount = 0; $before = $csvRows.Count
        $modelBlocks = [Collections.Generic.List[object]]::new(); $nativeTableCount = 0
        # Two passes: a table's query can reference ANOTHER table of the model (Source = VersionTable), so every table
        # is read first, and its query becomes referenceable next to the shared expressions.
        $tableFiles = [IO.FileInfo[]]@(Get-ChildItem -LiteralPath (Join-Path $definition 'tables') -Filter '*.tmdl' -File)
        $tableKeys = [string[]]@($tableFiles | ForEach-Object { $_.Name })
        [Array]::Sort([Array]$tableKeys, [Array]$tableFiles, [Collections.IComparer][StringComparer]::OrdinalIgnoreCase)
        $parsedTables = @(foreach ($tableFile in $tableFiles) {
                Assert-NotLinked $tableFile 'Unsafe input file'
                Read-TmdlTableFile $tableFile.FullName
            })
        foreach ($name in $info.Expressions.Keys) { $info.Queries[$name] = $info.Expressions[$name] }
        foreach ($parsed in $parsedTables) {
            if ($info.Queries.Contains($parsed.Name)) { continue }
            $tableQuery = if ($parsed.PolicySource) { $parsed.PolicySource } else { $null; foreach ($p in $parsed.Partitions) { if ($p.Type -eq 'm' -and $p.Source) { $p.Source; break } } }
            if ($tableQuery) { $info.Queries[$parsed.Name] = [string]$tableQuery }
        }
        foreach ($table in $parsedTables) {
            $tablesDone++
            Set-UiStatus "Analysing model $modelCount of $($selectedModels.Count)" $tablesDone $tableFilesTotal "$modelName $($script:Ui.Dot) $($table.Name)"
            if (-not $IncludeAutoDateTables -and $table.Name -match '\A(LocalDateTable|DateTableTemplate)_[0-9a-fA-F-]{36}\z') { continue }
            $tableCount++
            # One file per table. Distinct partition queries (incremental refresh, hybrid tables) are listed one after the other.
            $queries = [Collections.Generic.List[object]]::new()
            if ($table.PolicySource) { $queries.Add([pscustomobject]@{ Label = 'refresh policy source expression'; Type = 'm'; Text = $table.PolicySource }) }
            foreach ($p in $table.Partitions) {
                if (-not $p.Source) { continue }
                $duplicate = $false; foreach ($q in $queries) { if ($q.Text -ceq $p.Source -and $q.Type -eq $p.Type) { $duplicate = $true; break } }
                if ($duplicate) { continue }
                $queries.Add([pscustomobject]@{ Label = "partition '$($p.Name)' ($($p.Type)$(if ($p.Mode) { ", $($p.Mode)" }))"; Type = $p.Type; Text = $p.Source })
            }
            $queryFile = ''; $sqlFile = ''
            $tableFileName = Get-UniqueName (ConvertTo-FileName $table.Name) $usedFiles
            $nativeQueries = [Collections.Generic.List[object]]::new()
            $seenSql = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
            if ($queries.Count) {
                $types = @($queries | ForEach-Object Type)
                $extension = if ($types -cne 'calculated') { if ($types -ceq 'query') { '.sql' } else { '.pq' } } else { '.dax' }
                $comment = if ($extension -eq '.sql') { '--' } else { '//' }
                $text = [Text.StringBuilder]::new("$comment Semantic model : $modelName`n$comment Table          : $($table.Name)`n")
                foreach ($q in $queries) { [void]$text.Append("$comment Query          : $($q.Label)`n`n$($q.Text)`n`n") }
                $queryRelative = Join-Path $(if ($extension -eq '.dax') { 'dax' } else { 'power-query' }) ($tableFileName + $extension)
                $queryFile = Join-Path $relativeModel $queryRelative
                Write-ModelQueryFile $queryRelative ($text.ToString().TrimEnd() + "`n")
            }
            # Source rows.
            $found = [Collections.Generic.List[object]]::new()
            $modes = @($table.Partitions | ForEach-Object { $_.Mode } | Where-Object { $_ } | Select-Object -Unique) -join ', '
            $add = { param($SourceType, $Server, $Database, $Object, $ObjectType, $Via, $Parameters, $Review)
                # db.schema.object or server.db.schema.object: the database (and linked server) written in front of a SQL
                # object goes to its own column, so that Object always reads schema.object and Database says which
                # database is really read (a prefix wins over the database of the connection).
                if ($Object -and $ObjectType -in $script:SqlObjectTypes) {
                    $parts = @(Split-SqlName $Object)
                    if ($parts.Count -eq 4) { if ($parts[0] -ne '') { $Server = $parts[0] }; $parts = @($parts[1..3]) }
                    if ($parts.Count -eq 3) { if ($parts[0] -ne '') { $Database = $parts[0] }; $Object = Format-SqlName @($parts[1..2] | Where-Object { $_ -ne '' }) }
                }
                # No password or key ever reaches the inventory (a URL or a parameter value may carry one).
                $found.Add([pscustomobject][ordered]@{ Workspace = $workspaceName; SemanticModel = $modelName; Table = $table.Name; SourceType = $SourceType; Server = (Hide-Secret $Server)
                        Database = (Hide-Secret $Database); Object = (Hide-Secret $Object); ObjectType = $ObjectType; Mode = $modes; Via = $Via
                        Parameters = (Hide-Secret $Parameters); NeedsReview = $(if ($Review) { 'yes' } else { '' }); QueryFile = $queryFile; SqlFile = '' }) }
            # Braces are kept by the ODBC "driver={…}" key of a connection string; anywhere else in a server, database or
            # object name they are an unresolved M fragment: never certain.
            $hasPlaceholder = { param([string]$Text) ($Text -replace '(?i)\bdriver=\{[^{}]*\}', '').Contains('{') }
            $convert = { param($Rows, [string]$QueryOrigin)
                foreach ($row in $Rows) {
                    $c = $row.Context
                    if ($row.ContainsKey('Sql') -and $seenSql.Add($row.Sql)) {
                        $origin = $row.SqlOrigin
                        if ($row.Via) {
                            $viaKind = if ($info.Expressions.Contains($row.Via)) { 'shared expression' } elseif ($info.Queries.Contains($row.Via)) { 'table' } else { '' }
                            $origin = "$viaKind $($row.Via) / $origin".Trim()
                        } elseif ($origin) { $origin = "$QueryOrigin / $origin" } else { $origin = $QueryOrigin }
                        $nativeQueries.Add(@{ Text = $row.Sql; Origin = $origin; Context = $c; Parser = $row.Parser; Resolved = $row.SqlResolved })
                    }
                    $database = if ($row.Database) { $row.Database } elseif ($c.Database) { $c.Database } else { $c.Containers -join ', ' }
                    $parameters = @(@($c.Params) + @($row.Params) | Select-Object -Unique) -join '; '
                    $review = $row.Review -or $c.Review -or (& $hasPlaceholder "$($c.Server)$database$($row.Object)")
                    & $add $c.SourceType $c.Server $database $row.Object $row.ObjectType $row.Via $parameters $review
                }
            }
            $mTexts = @(foreach ($q in $queries) { if ($q.Type -eq 'm') { $q.Text } })
            foreach ($q in $queries) { if ($q.Type -eq 'm') { & $convert (Get-QuerySources $q.Text $info @($table.Name)).Rows $q.Label } }   # a table never follows itself
            foreach ($p in $table.Partitions) {
                if ($p.Type -eq 'entity') {
                    $context = $null
                    if ($p.ExpressionSource -and $info.Expressions.Contains($p.ExpressionSource)) { $context = (Get-QuerySources $info.Expressions[$p.ExpressionSource] $info @($p.ExpressionSource)).Context }
                    $object = if ($p.SchemaName) { "$($p.SchemaName).$($p.EntityName)" } else { $p.EntityName }
                    # 'entity' partitions are used by Direct Lake AND by composite models (DirectQuery to another semantic model / Analysis Services).
                    $kind = if ($p.Mode -ieq 'directLake') { 'Direct Lake entity' } else { 'Entity of another model / source' }
                    if ($null -ne $context) {
                        $database = if ($context.Database) { $context.Database } else { $context.Containers -join ', ' }
                        $doubt = [bool]$context.Review -or (& $hasPlaceholder "$($context.Server)$database")
                        & $add "$kind ($($context.SourceType))" $(if ($context.Server) { $context.Server } else { $context.BasePath }) $database $object 'Table' $p.ExpressionSource (@($context.Params | Select-Object -Unique) -join '; ') $doubt
                    } else { & $add $kind '' '' $object 'Table' $p.ExpressionSource '' (-not $p.ExpressionSource) }
                } elseif ($p.Type -eq 'query' -and $p.Source) {
                    # Legacy (provider) data source: the partition holds native SQL, read like any native query (same doubts),
                    # with the server and database of its data source when dataSources.tmdl names them.
                    $context = New-SourceContext "Legacy data source ($($p.DataSource))" 'db'
                    $known = if ($p.DataSource -and $dataSources.ContainsKey($p.DataSource)) { $dataSources[$p.DataSource] } else { $null }
                    if ($null -ne $known) { $context.Server = $known.Server; $context.Database = $known.Database }
                    # Legacy data sources are mostly SQL Server: unless the provider says otherwise, text the T-SQL grammar
                    # rejects is not certain.
                    $context.ExpectsTSql = $null -eq $known -or -not $known.Provider -or $known.Provider -match '(?i)SQLNCLI|MSOLEDBSQL|SQLOLEDB|SqlClient|SQL Server'
                    $legacyRows = [Collections.Generic.List[object]]::new()
                    Add-SqlTextRows $legacyRows $context ([pscustomobject]@{ Text = $p.Source; Resolved = $true; Opaque = $false; Params = [string[]]@(); Placeholders = [string[]]@() }) ''
                    & $convert $legacyRows "legacy query partition '$($p.Name)'"
                }
            }
            if (-not $found.Count) {
                $types = @($table.Partitions | ForEach-Object { $_.Type } | Select-Object -Unique)
                if ($types -contains 'calculated' -and $types.Count -eq 1) { & $add 'Calculated table (DAX)' '' '' '' '' '' '' $false }
                elseif ($mTexts.Count -and @($mTexts | Where-Object { $_ -match 'Table\.FromRows\s*\(\s*Json\.Document\s*\(\s*Binary\.Decompress\s*\(\s*Binary\.FromText' }).Count) { & $add 'Entered data (embedded in the model)' '' '' '' '' '' '' $false }
                # No connector call, no navigation, no reference to a query that has one (DateTime.LocalNow(), #table(...), a list...):
                # nothing looks like a source, so this is not a doubt. Unknown connectors get their own row, flagged for review.
                elseif ($mTexts.Count) { $references = @($mTexts | ForEach-Object { Get-ReferencedExpressions $_ $info.Queries $info.Literals } | Where-Object { $_ -and $_ -ne $table.Name } | Select-Object -Unique) -join ', '; & $add 'No external source (generated in M)' '' '' '' '' $references '' $false }
                elseif ($types.Count) { & $add "Partition type: $($types -join ', ')" '' '' '' '' '' '' $false }
                else { & $add 'No partition' '' '' '' '' '' '' $false }
            }
            if ($nativeQueries.Count) {
                $nativeTableCount++
                $sqlRelative = Join-Path 'sql' ($tableFileName + '.sql')
                $sqlFile = Join-Path $relativeModel $sqlRelative
                $blocks = foreach ($native in $nativeQueries) {
                    $c = $native.Context
                    $header = "-- Semantic model : $modelName`n-- Table : $($table.Name)`n-- Source : $($c.Function) | server: $($c.Server) | database: $($c.Database)`n-- Origin : $($native.Origin)`n-- Analysed by : $($native.Parser)`n"
                    if (-not $native.Resolved) { $header += "-- WARNING : unresolved M fragments are shown as {name} (one value) or {?name} (opaque text); this SQL is not runnable as is.`n" }
                    $header + "`n" + $native.Text + "`n"
                }
                Write-ModelQueryFile $sqlRelative ($blocks -join "`n")
            }
            foreach ($row in $found) { $row.SqlFile = $sqlFile }
            $block = [Text.StringBuilder]::new("-- ============================================================`n-- Table : $($table.Name)`n-- Mode  : $modes`n-- ============================================================`n-- Power Query :`n")
            foreach ($q in $queries) {
                foreach ($line in ($q.Text -split '\r\n|\n|\r')) { [void]$block.Append('--     ').Append($line).Append("`n") }
            }
            [void]$block.Append("--`n")
            if ($nativeQueries.Count) { [void]$block.Append((@(foreach ($native in $nativeQueries) { $native.Text }) -join "`n`n") + "`n") }
            else { [void]$block.Append("-- (no native SQL: navigation / calculated table / generated in M / entity partition …)`n") }
            $modelBlocks.Add([pscustomobject]@{ Table = $table.Name; Text = $block.ToString() })
            # One row per object: the same object can be reached by several partitions, or both directly and through a
            # shared expression (the paths are then merged in Via). A merged row is to review when ANY of its paths is:
            # a merge never clears a flag.
            $merged = [ordered]@{}
            foreach ($row in $found) {
                $key = "$($row.SourceType)|$($row.Server)|$($row.Database)|$($row.Object)"
                if (-not $merged.Contains($key)) { $merged[$key] = $row; continue }
                $kept = $merged[$key]
                $kept.Via = @(@($kept.Via -split ', ') + @($row.Via -split ', ') | Where-Object { $_ } | Select-Object -Unique) -join ', '
                $kept.Parameters = @(@($kept.Parameters -split '; ') + @($row.Parameters -split '; ') | Where-Object { $_ } | Select-Object -Unique) -join '; '
                if ($row.NeedsReview) { $kept.NeedsReview = 'yes' }
            }
            foreach ($row in $merged.Values) { $csvRows.Add($row) }
        }
        $modelText = [Text.StringBuilder]::new("-- Semantic model : $modelName`n-- Workspace : $workspaceName`n-- Generated by Parse-SemanticModels.ps1 on $([DateTime]::UtcNow.ToString('o', [cultureinfo]::InvariantCulture))`n-- Tables : $tableCount, with native SQL : $nativeTableCount`n`n")
        $blockOrder = $modelBlocks.ToArray(); $blockKeys = [string[]]@(foreach ($b in $blockOrder) { $b.Table })
        [Array]::Sort([Array]$blockKeys, [Array]$blockOrder, [Collections.IComparer][StringComparer]::OrdinalIgnoreCase)
        [void]$modelText.Append((@(foreach ($b in $blockOrder) { $b.Text }) -join "`n"))
        Write-ModelQueryFile ((Split-Path $modelOut -Leaf) + '.sql') $modelText.ToString()
        $tableTotal += $tableCount
        $added = $csvRows.GetRange($before, $csvRows.Count - $before)
        $review = 0; foreach ($row in $added) { if ($row.NeedsReview -eq 'yes') { $review++ } }
        if ($workspaceName -and -not [string]::Equals($workspaceName, $shownWorkspace, [StringComparison]::Ordinal)) { Write-UiLine "    $(Format-UiText Bold $workspaceName)"; $shownWorkspace = $workspaceName }
        $mark = if ($review) { Format-UiText Yellow $script:Ui.WarnMark } else { Format-UiText Green $script:Ui.OkMark }
        Write-UiLine "      $mark  $(Format-UiName $modelName $nameWidth)  $(Format-UiCount $tableCount 'table' 4)  $(Format-UiCount $added.Count 'source' 5)$(if ($review) { '  ' + (Format-UiText Yellow "$review to review") })"
    }
    if (-not $modelCount -and $skippedNew) {
        Stop-UiStatus
        Write-UiLine "  $(Format-UiText Green $script:Ui.OkMark)  Nothing new: the $(Format-UiCount $skippedNew 'selected model') $(if ($skippedNew -eq 1) { 'is' } else { 'are' }) already in the inventory."
        Write-UiLine
        return
    }
    if (-not $modelCount) {
        $available = @($modelFolders | ForEach-Object { if ($_.Workspace) { "'$($_.Folder.Parent.Name)'" } else { '(models directly under -ExportPath)' } } | Sort-Object -Unique)
        throw "No exported semantic model matches -Workspace / -Exclude / -Model under $ExportPath. Workspace(s) exported there so far: $(@($available | Select-Object -First 8) -join ', ')$(if ($available.Count -gt 8) { " and $($available.Count - 8) more" }). This script is offline and never contacts Fabric: export the workspace first with .\Extract-SemanticModels.ps1 -Workspace '<name>', or run both steps at once with .\Extract-And-Parse-SemanticModels.ps1 -Workspace '<name>'."
    }
    $dot = " $($script:Ui.Dot) "
    $modelsText = "$modelCount analysed$dot$(Format-UiCount $tableTotal 'table')$(if ($skippedNew) { "$dot$skippedNew already in the inventory" })$(if ($notTmdl.Count) { "$dot$($notTmdl.Count) skipped (no TMDL)" })"
    $parserText = if ($null -ne $script:TSql) { "Microsoft ScriptDom $($script:TSql.Version.Major).$($script:TSql.Version.Minor)" }
                  elseif ($NoTSqlParser) { 'built-in scanner (-NoTSqlParser)' }
                  else { "built-in scanner $(Format-UiText Dim "(more precise T-SQL with Microsoft's parser: Install-Module SqlServer -Scope CurrentUser)")" }
    if ($QueryFilesOnly) {
        Stop-UiStatus
        [Threading.Thread]::CurrentThread.CurrentCulture = $userCulture   # the summary in the user's number format
        Write-UiSummary ([ordered]@{ Models = $modelsText; 'SQL parser' = $parserText; 'Query files' = (Format-UiPath $OutputPath) })
        if ($PassThru) { $csvRows }
        return
    }
    Set-UiStatus 'Writing the CSV file'
    $columns = @('Workspace', 'SemanticModel', 'Table', 'SourceType', 'Server', 'Database', 'Object', 'ObjectType', 'Mode', 'Via', 'Parameters', 'NeedsReview', 'QueryFile', 'SqlFile')
    # A filtered run refreshes only the selected models: the rows of the other models already in the CSV are kept. The CSV
    # holds values protected against formulas (see Protect-CsvValue); the kept rows get their plain values back here.
    $kept = @()
    if (($Workspace -or $Model -or $Exclude -or $NewOnly) -and (Test-Path -LiteralPath $csvPath -PathType Leaf)) {
        $existing = @(Import-Csv -LiteralPath $csvPath -Delimiter $Delimiter[0])
        $sameLayout = $existing.Count -and -not (Compare-Object @($existing[0].PSObject.Properties.Name) $columns)
        if ($sameLayout) {
            $kept = @(foreach ($row in $existing) {
                    if ($processed.Contains("$($row.Workspace)|$($row.SemanticModel)")) { continue }
                    $copy = [ordered]@{}; foreach ($column in $columns) { $copy[$column] = [regex]::Replace([string]$row.$column, "\A'(?=[=+\-@\t\r])", '') }; [pscustomobject]$copy
                })
        } elseif ($existing.Count) { Write-UiWarning "The existing $csvPath has another layout or delimiter: it is rebuilt with the selected models only." }
    }
    # Fixed order: workspace, model, table, object (ordinal, ignoring case), whatever the machine.
    $finalRows = [object[]]@(@($kept) + @($csvRows))
    $sortKeys = [string[]]::new($finalRows.Count)
    for ($i = 0; $i -lt $finalRows.Count; $i++) { $row = $finalRows[$i]; $sortKeys[$i] = "$($row.Workspace)$([char]1)$($row.SemanticModel)$([char]1)$($row.Table)$([char]1)$($row.Object)$([char]1)$($i.ToString('D9'))" }   # last part: stable
    [Array]::Sort([Array]$sortKeys, [Array]$finalRows, [Collections.IComparer][StringComparer]::OrdinalIgnoreCase)
    $stale = Get-Item -LiteralPath $csvPath -Force -ErrorAction Ignore
    if ($stale) { $stale.Delete() }
    # Written by hand (Export-Csv is slow on large inventories): every value quoted, formulas neutralised, UTF-8 with BOM
    # so that Excel reads the accents. Operators only in the loop (see Get-InventorySheets); the formula test is the one
    # of Protect-CsvValue.
    $separator = [string]$Delimiter[0]
    $lines = foreach ($row in $finalRows) {
        @(foreach ($column in $columns) { $value = [string]$row.$column; if ($value -match '^[=+\-@\t\r]') { $value = "'" + $value }; '"' + ($value -replace '"', '""') + '"' }) -join $separator
    }
    $csv = (@(@(foreach ($column in $columns) { "`"$column`"" }) -join $separator) + @($lines) -join "`r`n") + "`r`n"
    $lines = $null
    $stream = [IO.File]::Open($csvPath, [IO.FileMode]::CreateNew)
    try { $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($true)); try { $writer.Write($csv) } finally { $writer.Dispose() } } finally { $stream.Dispose() }
    $csv = $null
    # Same rows as an Excel workbook. Cells are written as text, so nothing can be evaluated as a formula.
    Set-UiStatus 'Writing the Excel workbook'
    $xlsxPath = Join-Path $OutputPath 'model_sources.xlsx'
    $excelText = Format-UiPath $xlsxPath
    try { Export-SimpleXlsx $xlsxPath @(Get-InventorySheets $finalRows) }
    catch [IO.IOException], [UnauthorizedAccessException] { Write-UiWarning "model_sources.xlsx was not updated (is it open in Excel?): $($_.Exception.Message) The CSV is up to date."; $excelText = Format-UiText Red 'not updated (is it open in Excel?)' }
    # Same rows again as a web page to explore them (search, impact, lineage, query code), opened offline by double-click.
    Set-UiStatus 'Writing the web page'
    $htmlPath = Join-Path $OutputPath 'model_sources.html'
    $htmlText = Format-UiPath $htmlPath
    try { Export-HtmlReport $htmlPath $finalRows $OutputPath }
    catch [Management.Automation.PipelineStoppedException] { throw }
    catch {
        # Never lose the CSV and the workbook because of the page (a locked file, not enough memory for a huge inventory...).
        Write-UiWarning "model_sources.html was not updated: $($_.Exception.Message) The CSV and the workbook are up to date."
        $htmlText = Format-UiText Red 'not updated'
    }
    Stop-UiStatus
    $reviewCount = 0; foreach ($row in $csvRows) { if ($row.NeedsReview -eq 'yes') { $reviewCount++ } }
    $sourcesText = "$(Format-UiCount $csvRows.Count 'row')$dot$(if ($reviewCount) { Format-UiText Yellow "$reviewCount to review" } else { '0 to review' })$(if ($kept.Count) { "$dot$($kept.Count) rows of other models kept" })"
    [Threading.Thread]::CurrentThread.CurrentCulture = $userCulture   # the summary in the user's number format
    Write-UiSummary ([ordered]@{ Models = $modelsText; Sources = $sourcesText; 'SQL parser' = $parserText; Excel = $excelText; 'Web page' = $htmlText; CSV = (Format-UiPath $csvPath); 'Query files' = (Format-UiPath $OutputPath) })
    if ($PassThru) { $csvRows }
} finally {
    Stop-UiStatus
    [Threading.Thread]::CurrentThread.CurrentCulture = $userCulture
}

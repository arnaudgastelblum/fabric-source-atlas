# Offline test of Extract-SemanticModels.ps1 with HIGH FIDELITY: a fake Fabric API runs on a local
# HttpListener (localhost only) and the script goes through the REAL Invoke-WebRequest cmdlet, so real
# parameter validation, real header shapes and real 202/429 handling are exercised. Only `az` and
# Start-Sleep are stubbed. Nothing leaves the machine.
$ErrorActionPreference = 'Stop'
$target = Join-Path (Split-Path $PSScriptRoot -Parent) 'Extract-SemanticModels.ps1'
$out = Join-Path ([IO.Path]::GetTempPath()) "pstest_$(Get-Random)"
$api = 'https://api.fabric.microsoft.com'
$port = Get-Random -Minimum 20000 -Maximum 40000
$wsId = '11111111-2222-3333-4444-555555555555'
$wsName = '[TEST] WS   Name (acc)'
$opId = '99999999-8888-7777-6666-555555555555'
$shared = [hashtable]::Synchronized(@{ calls = [Collections.ArrayList]::Synchronized([Collections.ArrayList]::new())
        evil = ''; throttled = $false; retryAfter = '7'; failures = 0; polls = 0; stop = $false; ready = $false; error = $null })

$server = Start-ThreadJob -ArgumentList $port, $shared, $api, $wsId, $wsName, $opId -ScriptBlock {
    param($port, $shared, $api, $wsId, $wsName, $opId)
    function b64([string]$s) { [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s)) }
    function Get-Parts([string]$tag) {
        @{ definition = @{ parts = @(
            @{ path = 'definition/model.tmdl'; payload = (b64 "model Model`n`t/// $tag é€`n"); payloadType = 'InlineBase64' }
            @{ path = 'definition/tables/Sales.tmdl'; payload = (b64 "table Sales`n"); payloadType = 'InlineBase64' }
            @{ path = 'definition.pbism'; payload = (b64 '{}'); payloadType = 'InlineBase64' }
        ) } }
    }
    $models = @(
        @{ id = 'aaaaaaaa-0000-0000-0000-000000000001'; displayName = 'Sales: EU/US?' }
        @{ id = 'bbbbbbbb-0000-0000-0000-000000000002'; displayName = 'Sales_ EU_US_' }
        @{ id = 'cccccccc-0000-0000-0000-000000000003'; displayName = 'Broken' }
        @{ id = 'dddddddd-0000-0000-0000-000000000004'; displayName = '[Fin] Budget' }
        @{ id = 'eeeeeeee-0000-0000-0000-000000000005'; displayName = 'Evil' }
    )
    function Get-Route([string]$sig) {
        # Transient failures, once ($shared.failures counts them): token refused, server error, connection dropped (-1).
        if ($sig -eq 'GET /v1/workspaces' -and $shared.failures -eq 0) {
            switch ($shared.evil) {
                '401once' { $shared.failures++; return 401, @{ errorCode = 'TokenExpired'; message = 'Access token has expired' }, @{} }
                '503once' { $shared.failures++; return 503, @{ errorCode = 'ServiceUnavailable'; message = 'try later' }, @{} }
                'drop' { $shared.failures++; return -1, $null, @{} }
            }
        }
        if ($shared.evil -eq '401always') { return 401, @{ errorCode = 'TokenExpired'; message = 'Access token has expired' }, @{} }
        if ($shared.evil -eq 'status204' -and $sig -eq "GET /v1/workspaces/$wsId/semanticModels") { return 204, $null, @{} }
        if ($shared.evil -eq 'reverse' -and $sig -eq "GET /v1/workspaces/$wsId/semanticModels") { $list = @($models); [array]::Reverse($list); return 200, @{ value = $list }, @{} }
        if ($shared.evil -eq 'long' -and $sig -eq "GET /v1/workspaces/$wsId/semanticModels") { return 200, @{ value = @(@{ id = $models[0].id; displayName = ('L' * 200) + '1' }, @{ id = $models[3].id; displayName = ('L' * 200) + '2' }) }, @{} }
        # 'Flaky': visible and exportable, but listing its models fails.
        $flaky = @{ id = '00000000-0000-0000-0000-00000000000c'; displayName = 'Flaky'; type = 'Workspace' }
        if ($sig -eq 'GET /v1/workspaces/00000000-0000-0000-0000-00000000000c/semanticModels') { return 404, @{ errorCode = 'WorkspaceNotFound'; message = 'gone' }, @{} }
        switch ($sig) {
            # Workspaces where the login is Admin, Member or Contributor: not 'Other' (Viewer only).
            'GET /v1/workspaces?roles=Admin,Member,Contributor' {
                if ($shared.evil -eq 'noroles') { return 400, @{ errorCode = 'InvalidParameter'; message = 'roles' }, @{} }
                if ($shared.evil -eq 'emptyroles') { return 200, @{ value = @() }, @{} }
                if ($shared.evil -eq 'rolesid') { return 200, @{ value = @(@{ id = "$wsId/../x"; displayName = 'x' }) }, @{} }
                $mine = @(@{ id = $wsId; displayName = $wsName; type = 'Workspace' }, @{ id = '00000000-0000-0000-0000-00000000000b'; displayName = 'My workspace'; type = 'Personal' })
                if ($shared.evil -eq 'listfail') { $mine += $flaky }
                return 200, @{ value = $mine }, @{}
            }
            'GET /v1/workspaces' {
                if ($shared.evil -eq 'write') { return 200, @{ value = @(); continuationUri = "$api/v1/workspaces/$wsId/semanticModels/$($models[0].id)/updateDefinition" }, @{} }
                if ($shared.evil -eq 'onelake-ws') { return 200, @{ value = @(@{ id = '77777777-0000-0000-0000-000000000007'; displayName = 'OneLake - Microsoft' }) }, @{} }
                if ($shared.evil -eq 'cycle') { return 200, @{ value = @(); continuationUri = "$api/v1/workspaces?page=2" }, @{} }
                if ($shared.evil -eq 'host') { return 200, @{ value = @(); continuationUri = "https://api.fabric.micro$([char]0x200B)soft.com/v1/workspaces?page=2" }, @{} }
                $visible = @(@{ id = '00000000-0000-0000-0000-00000000000a'; displayName = 'Other'; type = 'Workspace' }, @{ id = '00000000-0000-0000-0000-00000000000b'; displayName = 'My workspace'; type = 'Personal' })
                if ($shared.evil -eq 'listfail') { $visible += $flaky }
                return 200, @{ value = $visible; continuationUri = "$api/v1/workspaces?page=2" }, @{}
            }
            # 'Other': the login is Viewer only, Fabric refuses every definition. 'My workspace' must never be read by -AllWorkspaces.
            'GET /v1/workspaces/00000000-0000-0000-0000-00000000000a/semanticModels' { return 200, @{ value = @(@{ id = 'ffffffff-0000-0000-0000-000000000006'; displayName = 'Viewed 1' }, @{ id = 'ffffffff-0000-0000-0000-000000000007'; displayName = 'Viewed 2' }) }, @{} }
            'POST /v1/workspaces/00000000-0000-0000-0000-00000000000a/semanticModels/ffffffff-0000-0000-0000-000000000006/getDefinition?format=TMDL' { return 403, @{ errorCode = 'InsufficientPrivileges'; message = 'The caller does not have sufficient permissions' }, @{} }
            'POST /v1/workspaces/00000000-0000-0000-0000-00000000000a/semanticModels/ffffffff-0000-0000-0000-000000000007/getDefinition?format=TMDL' { return 403, @{ errorCode = 'InsufficientPrivileges'; message = 'The caller does not have sufficient permissions' }, @{} }
            'GET /v1/workspaces/00000000-0000-0000-0000-00000000000b/semanticModels' { return 200, @{ value = @(@{ id = 'ffffffff-0000-0000-0000-000000000008'; displayName = 'Mine' }) }, @{} }
            'GET /v1/workspaces?page=2' { if ($shared.evil -eq 'cycle') { return 200, @{ value = @(); continuationUri = "$api/v1/workspaces" }, @{} }; return 200, @{ value = @(@{ id = $wsId; displayName = $wsName }) }, @{} }
            "GET /v1/workspaces/$wsId/semanticModels" {
                if ($shared.evil -eq 'device') { return 200, @{ value = @(@{ id = 'aaaaaaaa-0000-0000-0000-000000000001'; displayName = 'NUL' }) }, @{} }
                if ($shared.evil -eq 'id') { return 200, @{ value = @(@{ id = "$($models[0].id)/getDefinition?format=TMDL#\..\..\victim"; displayName = ' ' }) }, @{} }
                if ($shared.evil -eq 'throttle-always') { return 429, @{ errorCode = 'RequestBlocked'; message = 'slow down' }, @{ 'Retry-After' = '600' } }
                if (-not $shared.throttled) {
                    $shared.throttled = $true
                    # 'date': Retry-After as an HTTP date, 90 s from now.
                    $after = if ($shared.retryAfter -eq 'date') { [datetime]::UtcNow.AddSeconds(90).ToString('r') } else { $shared.retryAfter }
                    return 429, @{ errorCode = 'RequestBlocked'; message = 'slow down' }, @{ 'Retry-After' = $after }
                }
                return 200, @{ value = $models }, @{}
            }
            "GET /v1/operations/$opId" {
                $shared.polls++
                return 200, @{ status = $(if ($shared.polls -lt 3) { 'Running' } else { 'Succeeded' }) }, @{ 'Retry-After' = '2' }
            }
            "GET /v1/operations/$opId/result" { return 200, (Get-Parts 'LRO'), @{} }
        }
        if ($shared.evil -like 'part-*' -and $sig -like 'POST *getDefinition?format=TMDL') {
            $bad = switch ($shared.evil) { 'part-script' { 'Parse-SemanticModels.ps1' } 'part-cmd' { 'definition/az.cmd' } 'part-noext' { 'definition/runme' } 'part-device' { 'definition/tables/NUL.tmdl' } 'part-space' { 'definition/NUL /x.tmdl' } 'part-dot' { 'definition/evil.cmd./x.tmdl' } }
            return 200, @{ definition = @{ parts = @(@{ path = 'definition/model.tmdl'; payload = (b64 'model Model'); payloadType = 'InlineBase64' }, @{ path = $bad; payload = (b64 'Write-Host pwned'); payloadType = 'InlineBase64' }) } }, @{}
        }
        foreach ($m in $models) {
            if ($sig -eq "POST /v1/workspaces/$wsId/semanticModels/$($m.id)/getDefinition?format=TMDL") {
                switch ($m.displayName) {
                    'Broken' { return 400, @{ errorCode = 'OperationNotSupportedForItem'; message = 'not supported' }, @{} }
                    'Sales_ EU_US_' { return 202, $null, @{ 'x-ms-operation-id' = $opId; 'Retry-After' = '3'; Location = "$api/v1/operations/$opId" } }
                    'Evil' { return 200, @{ definition = @{ parts = @(@{ path = '../../escape.txt'; payload = (b64 'x'); payloadType = 'InlineBase64' }) } }, @{} }
                    default { return 200, (Get-Parts $m.displayName), @{} }
                }
            }
        }
        return 404, @{ errorCode = 'TestRouteNotFound'; message = $sig }, @{}
    }
    try {
        $listener = [System.Net.HttpListener]::new()
        $listener.Prefixes.Add("http://localhost:$port/")
        $listener.Start(); $shared.ready = $true
        while (-not $shared.stop) {
            $task = $listener.GetContextAsync()
            while (-not $task.Wait(100)) { if ($shared.stop) { $listener.Stop(); return } }
            $ctx = $task.Result; $req = $ctx.Request
            $sig = "$($req.HttpMethod) $($req.Url.PathAndQuery)"
            $bodyLength = 0
            if ($req.HasEntityBody) { $reader = [IO.StreamReader]::new($req.InputStream); $bodyLength = $reader.ReadToEnd().Length }
            [void]$shared.calls.Add([pscustomobject]@{ Sig = $sig; Auth = $req.Headers['Authorization']; BodyLength = $bodyLength })
            $status, $body, $headers = Get-Route $sig
            if ($status -eq -1) { $ctx.Response.Abort(); continue }
            $ctx.Response.StatusCode = $status
            foreach ($k in $headers.Keys) { $ctx.Response.Headers[$k] = $headers[$k] }
            if ($null -ne $body) {
                $bytes = [Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Depth 10 -Compress))
                $ctx.Response.ContentType = 'application/json; charset=utf-8'
                $ctx.Response.ContentLength64 = $bytes.Length
                $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
            }
            $ctx.Response.Close()
        }
    } catch { $shared.error = $_.ToString() }
}
while (-not $shared.ready) { if ($shared.error) { throw "listener failed: $($shared.error)" }; [Threading.Thread]::Sleep(50) }

$global:sleeps = [Collections.Generic.List[double]]::new(); $global:azCalls = 0; $global:timeouts = [Collections.Generic.List[int]]::new()
# Fake az: answers like `az account get-access-token -o json`, with a fake JWT (unsigned) whose payload names a test
# user and tenant. $global:azMode selects a failure or an expiry shape. The real az is never called.
$tid = 'abcdef01-2345-6789-abcd-ef0123456789'
function ConvertTo-Base64Url([string]$Text) { [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
function New-FakeJwt([hashtable]$Claims) { "$(ConvertTo-Base64Url '{"alg":"none","typ":"JWT"}').$(ConvertTo-Base64Url ($Claims | ConvertTo-Json -Compress)).FAKESIGNATURE" }
$global:fakeToken = New-FakeJwt @{ upn = 'tester@example.test'; tid = $tid; exp = [DateTimeOffset]::UtcNow.AddHours(1).ToUnixTimeSeconds() }
$fakePayload = $global:fakeToken.Split('.')[1]
$global:azMode = 'ok'; $global:azArgs = @()
function az {
    $global:azCalls++; $global:azArgs = @($args)
    $fail = switch ($global:azMode) {
        'notloggedin' { "ERROR: Please run 'az login' to setup account." }
        'mfa' { 'ERROR: AADSTS50076: Due to a configuration change made by your administrator, or because you moved to a new location, you must use multi-factor authentication to access the resource.' }
        'aadsts' { 'ERROR: AADSTS700082: The refresh token has expired due to inactivity.' }
        'nosub' { 'ERROR: No subscriptions found for tester@example.test.' }
        'leak' { "ERROR: unexpected failure near $global:fakeToken" }
    }
    if ($fail) { Write-Error $fail; $global:LASTEXITCODE = 1; return }
    $global:LASTEXITCODE = 0
    $now = [DateTimeOffset]::Now
    $answer = [ordered]@{ accessToken = $global:fakeToken; expiresOn = $now.AddHours(1).ToString('yyyy-MM-dd HH:mm:ss.ffffff', [cultureinfo]::InvariantCulture); expires_on = $now.AddHours(1).ToUnixTimeSeconds(); tenant = $tid; tokenType = 'Bearer' }
    switch ($global:azMode) {
        # Close to its end (20 s): must be refreshed before every request.
        'expiring' { $answer.expires_on = $now.AddSeconds(20).ToUnixTimeSeconds(); $answer.expiresOn = $now.AddSeconds(20).ToString('yyyy-MM-dd HH:mm:ss.ffffff', [cultureinfo]::InvariantCulture) }
        # Only the local-time string, and a token without exp: parsed with the invariant culture.
        'localonly' { $answer.Remove('expires_on'); $answer.accessToken = New-FakeJwt @{ upn = 'tester@example.test'; tid = $tid } }
        # No expiry at all, and an application token (appid, no upn).
        'noexpiry' { $answer.Remove('expires_on'); $answer.Remove('expiresOn'); $answer.accessToken = New-FakeJwt @{ appid = 'app-test-id'; tid = $tid } }
    }
    $answer | ConvertTo-Json
}
function Start-Sleep { param($Seconds) $global:sleeps.Add([double]$Seconds) }
# Thin wrapper: same parameters, REAL cmdlet, host swapped to the local listener.
function Invoke-WebRequest {
    [CmdletBinding()]
    param($Method, $Uri, $Headers, [switch]$SkipHttpErrorCheck, $MaximumRedirection, $TimeoutSec, $ContentType, $Body)
    if (-not $Uri.StartsWith("$api/", [StringComparison]::Ordinal)) { throw "script tried to reach an unexpected host: $Uri" }
    $global:timeouts.Add([int]$TimeoutSec)
    $p = @{} + $PSBoundParameters
    $p.Uri = "http://localhost:$port" + $Uri.Substring($api.Length)
    Microsoft.PowerShell.Utility\Invoke-WebRequest @p
}
function Assert($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
function Get-Sigs { @($shared.calls | ForEach-Object Sig) }

try {
    # ---- Run 1: name with brackets, wrong case, wildcard across the triple space ----
    # -Verbose: no live status line, so each wait is ONE Start-Sleep call whose length can be checked.
    $r = @(& $target -Workspace '[test] ws*name (ACC)' -OutputPath $out -Verbose 4>$null 6>$null)
    $ws = Join-Path $out $wsName
    Assert ($r.Count -eq 5) "5 results, got $($r.Count)"
    $st = @{}; $r | ForEach-Object { $st[$_.Model] = $_.Status }
    Assert ($st['Sales: EU/US?'] -eq 'Exported' -and $st['Sales_ EU_US_'] -eq 'Exported' -and $st['[Fin] Budget'] -eq 'Exported') "exported statuses: $($r | ForEach-Object { "$($_.Model)=$($_.Status) $($_.Error)" })"
    Assert ($st['Broken'] -eq 'Failed' -and $st['Evil'] -eq 'Failed') 'failed statuses'
    Assert (($r | Where-Object Model -EQ 'Broken').Error -match 'HTTP 400: OperationNotSupportedForItem') "fabric error surfaced: $(($r | Where-Object Model -EQ 'Broken').Error)"
    Assert (($r | Where-Object Model -EQ 'Evil').Error -match 'Unsafe|escapes') 'traversal refused'
    $lro = Join-Path $ws 'Sales_ EU_US_ (bbbbbbbb).SemanticModel\definition\model.tmdl'
    Assert (Test-Path -LiteralPath (Join-Path $ws 'Sales_ EU_US_.SemanticModel\definition\tables\Sales.tmdl')) 'sanitized model folder'
    Assert (Test-Path -LiteralPath $lro) 'duplicate folder name gets id suffix'
    Assert (Test-Path -LiteralPath (Join-Path $ws '[Fin] Budget.SemanticModel\definition.pbism')) 'bracket model name'
    $expected = [Text.Encoding]::UTF8.GetBytes("model Model`n`t/// LRO é€`n")
    Assert ([Convert]::ToBase64String([IO.File]::ReadAllBytes($lro)) -eq [Convert]::ToBase64String($expected)) 'bytes preserved exactly (no BOM, no re-encoding)'
    Assert ($global:sleeps -contains 7) "Retry-After honoured on 429: $($global:sleeps -join ',')"
    Assert ($shared.polls -eq 3 -and (Get-Sigs) -contains "GET /v1/operations/$opId/result") 'LRO polled then result fetched'
    Assert ((Get-Sigs) -contains 'GET /v1/workspaces?page=2') 'paging followed'
    Assert ($global:azCalls -eq 1) "token cached, az called $($global:azCalls)x"
    Assert (-not ($shared.calls | Where-Object Auth -NE "Bearer $global:fakeToken")) 'every request carried the bearer token'
    Assert (-not ($shared.calls | Where-Object BodyLength -GT 0)) 'no request ever carried a body'
    Assert (-not ($global:timeouts | Where-Object { $_ -lt 1 -or $_ -gt 300 })) "per-request timeouts within 1..300 s: $($global:timeouts | Sort-Object -Unique)"
    $verbs = @($shared.calls | ForEach-Object { ($_.Sig -split ' ')[0] } | Sort-Object -Unique)
    Assert (($verbs -join ',') -eq 'GET,POST') "verbs on the wire: $($verbs -join ',')"
    Assert (-not ($shared.calls | Where-Object { $_.Sig -like 'POST *' -and $_.Sig -notlike '*/getDefinition?format=TMDL' })) 'the only POST is getDefinition'
    $csv = @(Import-Csv -LiteralPath (Join-Path $ws '_export_report.csv'))
    Assert ($csv.Count -eq 5 -and ($csv | Where-Object Model -EQ '[Fin] Budget').TmdlFiles -eq '2') 'report csv'
    'RUN 1 OK  -> requests seen by the fake server:'; Get-Sigs | Group-Object { $_ -replace '[0-9a-f]{8}-[0-9a-f-]{27}', '{id}' } | ForEach-Object { "     x$($_.Count)  $($_.Name)" }

    # ---- Run 2: re-export. Stale file disappears; a failed model keeps its previous export ----
    Set-Content -LiteralPath (Join-Path $ws '[Fin] Budget.SemanticModel\definition\tables\Removed.tmdl') -Value 'stale'
    New-Item -ItemType Directory -Path (Join-Path $ws 'Broken.SemanticModel') | Out-Null
    Set-Content -LiteralPath (Join-Path $ws 'Broken.SemanticModel\previous.tmdl') -Value 'previous good export'
    $shared.polls = 0
    $r2 = @(& $target -Workspace $wsId -OutputPath $out -Model '[fin]*', 'Broken' 6>$null)
    Assert ($r2.Count -eq 2) "model filter with literal brackets, got $($r2.Count)"
    Assert (-not (Test-Path -LiteralPath (Join-Path $ws '[Fin] Budget.SemanticModel\definition\tables\Removed.tmdl'))) 'stale file removed'
    Assert (Test-Path -LiteralPath (Join-Path $ws 'Broken.SemanticModel\previous.tmdl')) 'failed model keeps previous export'
    # -Model run: the report keeps the lines of the other models (merged by model id).
    $csv2 = @(Import-Csv -LiteralPath (Join-Path $ws '_export_report.csv'))
    Assert ($csv2.Count -eq 5 -and @($csv2.ModelId | Sort-Object -Unique).Count -eq 5 -and ($csv2 | Where-Object Model -EQ 'Sales: EU/US?').Status -eq 'Exported') "report merged by model id: $($csv2 | Format-Table Model, Status | Out-String)"
    'RUN 2 OK'

    # ---- Run 3: unknown workspace is a terminating error ----
    try { & $target -Workspace 'does not exist' -OutputPath $out 6>$null | Out-Null; throw 'expected failure' }
    catch { Assert ($_.Exception.Message -match 'No workspace matches') $_.Exception.Message; 'RUN 3 OK' }

    # ---- Run 4: hostile continuationUri pointing at a write endpoint is refused before anything is sent ----
    $shared.evil = 'write'; $shared.calls.Clear()
    try { & $target -Workspace $wsId -OutputPath $out 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'outside the read-only allow-list: GET /v1/workspaces/.+/updateDefinition') $_.Exception.Message }
    Assert (-not (Get-Sigs | Where-Object { $_ -match 'updateDefinition' })) 'write endpoint was never called'
    'RUN 4 OK'

    # ---- Run 5: lookalike host (zero-width space inside "microsoft") must never receive the token ----
    $shared.evil = 'host'; $shared.calls.Clear()
    try { & $target -Workspace $wsId -OutputPath $out 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'unexpected characters in the URL') $_.Exception.Message }
    Assert (@(Get-Sigs).Count -eq 1) "only the first GET /v1/workspaces was sent: $(Get-Sigs)"
    'RUN 5 OK'

    # ---- Run 6: hostile item id (path traversal hidden behind ?query#fragment) is rejected ----
    $shared.evil = 'id'; $shared.calls.Clear()
    try { & $target -Workspace $wsId -OutputPath $out 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'Unexpected item id in Fabric response') $_.Exception.Message }
    Assert (-not (Get-Sigs | Where-Object { $_ -like 'POST *' })) 'no getDefinition sent for a hostile id'
    $shared.evil = ''
    'RUN 6 OK'

    # ---- Run 7: -OutputPath inside a OneLake File Explorer folder is refused before any request or write ----
    $oneLake = Join-Path ([IO.Path]::GetTempPath()) "pstest_onelake_$(Get-Random)\OneLake - Microsoft\SomeWorkspace"
    $shared.calls.Clear()
    try { & $target -Workspace $wsId -OutputPath $oneLake 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'OneLake File Explorer') $_.Exception.Message }
    Assert (@(Get-Sigs).Count -eq 0 -and -not (Test-Path -LiteralPath $oneLake)) 'nothing sent, nothing created'
    'RUN 7 OK'

    # ---- Run 8: -Debug must not print the bearer token (Invoke-WebRequest dumps request headers in debug mode) ----
    $all = & $target -Workspace $wsId -OutputPath $out -Model '[fin]*' -Debug *>&1 | Out-String
    # Colours and check marks appear when the harness runs in an interactive terminal: compare without them.
    Assert (($all -replace "\[[0-9;?]*[A-Za-z]", '') -match '(\[OK\]|✓)\s+\[Fin\] Budget') "export still works under -Debug: $all"
    Assert ($all -notmatch "Bearer|FAKE|$fakePayload") "TOKEN LEAKED in debug output:`n$(($all -split "`n" | Select-String "Bearer|FAKE|$fakePayload" | Select-Object -First 3) -join "`n")"
    'RUN 8 OK'

    # ---- Run 9: a workspace NAMED like the sync root must not become <OutputPath>\OneLake - Microsoft ----
    $shared.evil = 'onelake-ws'; $shared.calls.Clear()
    $plain = Join-Path $out 'run9'
    try { & $target -Workspace '77777777-0000-0000-0000-000000000007' -OutputPath $plain 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'OneLake File Explorer') $_.Exception.Message }
    Assert (-not (Test-Path -LiteralPath (Join-Path $plain 'OneLake - Microsoft'))) 'no folder created under a OneLake-named workspace'
    Assert (-not (Get-Sigs | Where-Object { $_ -like 'POST *' -or $_ -like '*semanticModels*' })) 'refused before listing or exporting any model'
    $shared.evil = ''
    'RUN 9 OK'

    # ---- Run 10: one-command launcher = export (fake server only) + offline analysis of what was exported ----
    $shared.calls.Clear(); $shared.polls = 0
    $launcher = Join-Path (Split-Path $target) 'Extract-And-Parse-SemanticModels.ps1'
    $e10 = Join-Path $out 'w-export'; $s10 = Join-Path $out 'w-sources'
    $warnings = @(& $launcher -Workspace '[test] ws*name (ACC)' -ExportPath $e10 -OutputPath $s10 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
    Assert (@($warnings | Where-Object { "$_" -match '2 of 5 model\(s\) could not be exported' }).Count -eq 1) "failed exports are reported: $warnings"
    $rows10 = @(Import-Csv -LiteralPath (Join-Path $s10 'model_sources.csv') -Delimiter ((Get-Culture).TextInfo.ListSeparator[0]))
    $models10 = @($rows10.SemanticModel | Sort-Object -Unique) -join ' | '
    Assert ($models10 -eq '[Fin] Budget | Sales_ EU_US_ | Sales_ EU_US_ (bbbbbbbb)') "launcher analysed exactly the exported models: $models10"
    Assert (-not ($rows10 | Where-Object Workspace -NE $wsName)) 'workspace folder name passed through (brackets + triple space)'
    Assert (Test-Path -LiteralPath (Join-Path $s10 'model_sources.xlsx')) 'workbook written by the launcher'
    Assert (@(Get-Sigs | Where-Object { $_ -like 'POST *' }).Count -eq 5 -and -not (Get-Sigs | Where-Object { $_ -like 'POST *' -and $_ -notlike '*/getDefinition?format=TMDL' })) 'launcher adds no request of its own'
    'RUN 10 OK'

    # ---- Run 11: a non-empty folder that these scripts did not create is refused before any request ----
    Assert (Test-Path -LiteralPath (Join-Path $out '.semantic-model-tools') -PathType Leaf) 'marker written in a folder created by the script'
    $foreign = Join-Path $out 'foreign'; New-Item -ItemType Directory -Path $foreign | Out-Null
    Set-Content -LiteralPath (Join-Path $foreign 'precious.txt') -Value 'do not touch'
    $shared.calls.Clear()
    try { & $target -Workspace $wsId -OutputPath $foreign 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'was not created by these scripts') $_.Exception.Message }
    Assert (@(Get-Sigs).Count -eq 0 -and -not (Test-Path -LiteralPath (Join-Path $foreign '.semantic-model-tools'))) 'refused before any request, nothing written'
    $adopted = @(& $target -Workspace $wsId -OutputPath $foreign -Model '[fin]*' -AdoptOutputFolder 6>$null)
    Assert ($adopted.Count -eq 1 -and $adopted[0].Status -eq 'Exported' -and (Test-Path -LiteralPath (Join-Path $foreign 'precious.txt')) -and (Test-Path -LiteralPath (Join-Path $foreign '.semantic-model-tools'))) 'explicit adoption works and leaves existing files alone'
    'RUN 11 OK'

    # ---- Run 12: pagination that loops back on itself is stopped ----
    $shared.evil = 'cycle'; $shared.calls.Clear()
    try { & $target -Workspace $wsId -OutputPath $out 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'pagination does not terminate') $_.Exception.Message }
    Assert (@(Get-Sigs).Count -eq 2) "stopped after the repeated continuationUri: $(@(Get-Sigs).Count) requests"
    $shared.evil = ''
    'RUN 12 OK'

    # ---- Run 13: the scripts must not live inside the output folder (an exported part could replace them) ----
    $tools = Join-Path $out 'tools13'; New-Item -ItemType Directory -Path $tools | Out-Null
    Copy-Item -LiteralPath $target -Destination $tools
    $shared.calls.Clear()
    try { & (Join-Path $tools (Split-Path $target -Leaf)) -Workspace $wsId -OutputPath $out 6>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'the scripts themselves are inside it') $_.Exception.Message }
    Assert (@(Get-Sigs).Count -eq 0) 'refused before any request'
    'RUN 13 OK'

    # ---- Run 14: a model named like a Windows device gets a safe folder name ----
    $shared.evil = 'device'
    $dev = @(& $target -Workspace $wsId -OutputPath (Join-Path $out 'run14') 6>$null)
    Assert ($dev.Count -eq 1 -and $dev[0].Status -eq 'Exported' -and $dev[0].Folder.EndsWith('\_NUL.SemanticModel')) "device name sanitised: $($dev | Out-String)"
    $shared.evil = ''
    'RUN 14 OK'

    # ---- Run 15: only plain drive-letter output paths (\\?\C:\... names the same folder but defeats path comparisons) ----
    foreach ($spelling in ('\\?\' + $out), ('\\.\' + $out), '\\localhost\c$\pstest-never-created') {
        $shared.calls.Clear()
        try { & $target -Workspace $wsId -OutputPath $spelling 6>$null | Out-Null; throw 'expected refusal' }
        catch { Assert ($_.Exception.Message -match 'plain drive-letter path') "$spelling -> $($_.Exception.Message)" }
        Assert (@(Get-Sigs).Count -eq 0) "refused before any request: $spelling"
    }
    'RUN 15 OK'

    # ---- Run 16: a hostile definition cannot plant a script, an executable, an extension-less file or a device name ----
    foreach ($mode in 'part-script', 'part-cmd', 'part-noext', 'part-device', 'part-space', 'part-dot') {
        $shared.evil = $mode; $dir16 = Join-Path $out "run16-$mode"
        $r16 = @(& $target -Workspace $wsId -OutputPath $dir16 -Model '[fin]*' 6>$null)
        Assert ($r16.Count -eq 1 -and $r16[0].Status -eq 'Failed' -and $r16[0].Error -match 'Unexpected file type|reserved device name') "$mode -> $($r16 | Out-String)"
        Assert (-not @(Get-ChildItem -LiteralPath $dir16 -Recurse -File | Where-Object { $_.Name -notin '.semantic-model-tools', '_export_report.csv' }).Count) "$mode : nothing of that model was written"
    }
    $shared.evil = ''
    'RUN 16 OK'

    # ---- Run 17: -ListWorkspaces / -ListModels only read the listings, never ask for a definition, never write ----
    Push-Location $out
    try {
        # Snapshot of everything under the scratch folder (names, sizes, contents): the listings must change nothing.
        $snapshot = { (@(Get-ChildItem -LiteralPath $out -Recurse -Force | ForEach-Object { if ($_.PSIsContainer) { $_.FullName } else { "$($_.FullName)|$($_.Length)|$([Convert]::ToBase64String([IO.File]::ReadAllBytes($_.FullName)))" } } | Sort-Object) -join "`n") }
        $before17 = & $snapshot
        $shared.calls.Clear()
        $ws17 = @(& $target -ListWorkspaces 6>$null)
        Assert ($ws17.Count -eq 3 -and ((@($ws17.Workspace) | Sort-Object) -join '|') -eq ((@('My workspace', 'Other', $wsName) | Sort-Object) -join '|') -and ($ws17 | Where-Object Workspace -EQ $wsName).Id -eq $wsId) "workspaces listed with their id: $($ws17 | Out-String)"
        Assert ((@(Get-Sigs) -join ' ') -eq 'GET /v1/workspaces GET /v1/workspaces?page=2') "only the workspace listing was requested: $(Get-Sigs)"
        $shared.calls.Clear()
        $m17 = @(& $target -Workspace '[test] ws*' -ListModels -Model 'Sales*', '[fin]*' 6>$null)
        Assert ($m17.Count -eq 3 -and -not ($m17 | Where-Object Workspace -NE $wsName) -and (@($m17.Model) -join '|') -eq 'Sales: EU/US?|Sales_ EU_US_|[Fin] Budget' -and $m17[0].ModelId -eq 'aaaaaaaa-0000-0000-0000-000000000001' -and -not ($m17 | Where-Object WorkspaceId -NE $wsId)) "models listed and filtered: $($m17 | Out-String)"
        Assert ((@(Get-Sigs) -join ' ') -eq "GET /v1/workspaces GET /v1/workspaces?page=2 GET /v1/workspaces/$wsId/semanticModels") "exactly the two listings, no definition: $(Get-Sigs)"
        # A switch passed as false still selects the list parameter set: it must list, never fall through to an export.
        foreach ($arguments in @{ Workspace = $wsId; ListModels = $false }, @{ ListWorkspaces = $false }) {
            $shared.calls.Clear()
            $r17 = @(& $target @arguments 6>$null)
            Assert ($r17.Count -ge 2 -and -not (Get-Sigs | Where-Object { $_ -like 'POST *' -or $_ -like '*operations*' })) "a false list switch still only lists: $(Get-Sigs)"
        }
        Assert ((& $snapshot) -ceq $before17 -and -not (Test-Path -LiteralPath (Join-Path $out 'data'))) 'the listings wrote nothing anywhere under the scratch folder'
        try { & $target -ListWorkspaces -Workspace $wsId 6>$null | Out-Null; throw 'expected failure' }
        catch { Assert ($_.Exception.Message -match 'Parameter set cannot be resolved|expected failure' -and $_.Exception.Message -ne 'expected failure') "-ListWorkspaces takes no -Workspace: $($_.Exception.Message)" }
    } finally { Pop-Location }
    'RUN 17 OK'

    # ---- Run 18: /? -h --help help, or no argument at all, show the short help and send nothing ----
    $shared.calls.Clear()
    foreach ($word in '/?', '-h', '--help', 'help') {
        $help18 = & $target $word 6>&1 | Out-String
        Assert ($help18 -match '-AllWorkspaces' -and $help18 -match '-NewOnly' -and $help18 -match 'Examples') "'$word' shows the help of the exporter"
    }
    Assert ((& $target 6>&1 | Out-String) -match '-ListWorkspaces') 'no argument shows the help of the exporter'
    Assert ((& $target -h 6>&1 | Out-String) -match '-ListWorkspaces' -and (& $target -Workspace $wsId -Help 6>&1 | Out-String) -match '-ListWorkspaces') '-h / -Help switch shows the help'
    Assert ((& $launcher -h 6>&1 | Out-String) -match '-NewOnly') '-h shows the help of the launcher'
    Assert ((& $launcher '/?' 6>&1 | Out-String) -match '-Exclude') "'/?' shows the help of the launcher"
    Assert ((& $launcher 6>&1 | Out-String) -match '-AllWorkspaces') 'no argument shows the help of the launcher'
    Assert (@(Get-Sigs).Count -eq 0) "asking for help never sends a request: $(Get-Sigs)"
    foreach ($arguments in @{ Workspace = $wsId; AllWorkspaces = $true }, @{ Model = 'Sales*' }, @{ Exclude = 'Other' }) {
        try { & $target @arguments -OutputPath $out 6>$null | Out-Null; throw 'expected failure' }
        catch { Assert ($_.Exception.Message -match 'or use -AllWorkspaces') "$($arguments.Keys -join ', ') -> $($_.Exception.Message)" }
        try { & $launcher @arguments -ExportPath $out 6>$null | Out-Null; throw 'expected failure' }
        catch { Assert ($_.Exception.Message -match 'or use -AllWorkspaces') "launcher $($arguments.Keys -join ', ') -> $($_.Exception.Message)" }
    }
    Assert (@(Get-Sigs).Count -eq 0) "a wrong combination sends nothing: $(Get-Sigs)"
    'RUN 18 OK'

    # ---- Run 19: -AllWorkspaces leaves out My workspace (never read) and the Viewer-only workspace (one refused request, nothing written) ----
    $shared.calls.Clear(); $shared.polls = 0
    $d19 = Join-Path $out 'run19'
    $r19 = @(& $target -AllWorkspaces -OutputPath $d19 6>$null)
    Assert ($r19.Count -eq 5 -and -not ($r19 | Where-Object WorkspaceId -NE $wsId)) "only the test workspace is exported: $($r19 | Out-String)"
    Assert (-not (Get-Sigs | Where-Object { $_ -like '*00000000000b*' })) "My workspace is never read: $(Get-Sigs)"
    Assert ((Get-Sigs) -contains 'GET /v1/workspaces?roles=Admin,Member,Contributor') "workspaces filtered on the roles that can read definitions: $(Get-Sigs)"
    Assert (-not (Get-Sigs | Where-Object { $_ -like '*00000000000a*' })) "the Viewer-only workspace is left out before any request: $(Get-Sigs)"
    Assert (-not (Test-Path -LiteralPath (Join-Path $d19 'Other'))) 'nothing written for the Viewer-only workspace'
    Assert (Test-Path -LiteralPath (Join-Path $d19 "$wsName\_export_report.csv")) 'report of the exported workspace'
    $shared.calls.Clear()
    $m19 = @(& $target -AllWorkspaces -ListModels 6>$null)
    Assert ($m19.Count -eq 7 -and -not ($m19 | Where-Object Workspace -EQ 'My workspace')) "-AllWorkspaces -ListModels lists every workspace but My workspace: $($m19 | Out-String)"
    Assert (-not (Get-Sigs | Where-Object { $_ -like 'POST *' })) 'listing asks for no definition'
    'RUN 19 OK'

    # ---- Run 20: -Exclude (name pattern or id), with -AllWorkspaces or -Workspace ----
    $shared.calls.Clear()
    $r20 = @(& $target -AllWorkspaces -Exclude '[test]*' -OutputPath (Join-Path $out 'run20') 6>$null)
    Assert ($r20.Count -eq 0 -and -not (Get-Sigs | Where-Object { $_ -like "*$wsId*" })) "excluded workspace never read: $(Get-Sigs)"
    $shared.calls.Clear()
    $r20 = @(& $target -Workspace '[test]*', 'Other' -Exclude $wsId, 'oth*' -OutputPath (Join-Path $out 'run20') 6>$null)
    Assert ($r20.Count -eq 0 -and @(Get-Sigs).Count -eq 2) "-Exclude by id and by pattern, nothing but the workspace listing sent: $(Get-Sigs)"
    $m20 = @(& $target -AllWorkspaces -Exclude 'Other' -ListModels 6>$null)
    Assert ($m20.Count -eq 5 -and -not ($m20 | Where-Object Workspace -NE $wsName)) '-Exclude applies to -ListModels'
    'RUN 20 OK'

    # ---- Run 21: -NewOnly skips the models already exported, retries the others, keeps their report line ----
    $shared.calls.Clear(); $shared.polls = 0
    $r21 = @(& $target -Workspace $wsId -NewOnly -OutputPath $d19 6>$null)
    $st21 = @{}; $r21 | ForEach-Object { $st21[$_.Model] = $_.Status }
    Assert ($r21.Count -eq 5 -and $st21['Sales: EU/US?'] -eq 'AlreadyExported' -and $st21['Sales_ EU_US_'] -eq 'AlreadyExported' -and $st21['[Fin] Budget'] -eq 'AlreadyExported' -and $st21['Broken'] -eq 'Failed' -and $st21['Evil'] -eq 'Failed') "statuses: $($r21 | ForEach-Object { "$($_.Model)=$($_.Status)" })"
    $posts21 = @(Get-Sigs | Where-Object { $_ -like 'POST *' })
    Assert ($posts21.Count -eq 2 -and -not ($posts21 | Where-Object { $_ -notmatch 'cccccccc|eeeeeeee' })) "only the missing models are requested: $posts21"
    $csv21 = @(Import-Csv -LiteralPath (Join-Path $d19 "$wsName\_export_report.csv"))
    $budget21 = $csv21 | Where-Object Model -EQ '[Fin] Budget'
    Assert ($csv21.Count -eq 5 -and $budget21.Status -eq 'AlreadyExported' -and $budget21.TmdlFiles -eq '2' -and $budget21.ExportedAtUtc) "report keeps the previous export of skipped models: $($budget21 | Out-String)"
    'RUN 21 OK'

    # ---- Run 22: a Viewer-only workspace named explicitly: tried (never left out on the role filter alone), nothing written ----
    $d22 = Join-Path $out 'run22'
    # Named explicitly: no role listing needed, its first definition is refused, a clear message (not "no model matches").
    $shared.calls.Clear()
    try { & $target -Workspace 'Other' -OutputPath $d22 6>$null | Out-Null; throw 'expected failure' }
    catch { Assert ($_.Exception.Message -match "Viewer access only to 'Other'.+Contributor") $_.Exception.Message }
    Assert (@(Get-Sigs | Where-Object { $_ -like 'POST *' }).Count -eq 1 -and -not (Get-Sigs | Where-Object { $_ -like '*roles=*' }) -and -not (Test-Path -LiteralPath (Join-Path $d22 'Other'))) "Viewer-only workspace named: one refused request, nothing written: $(Get-Sigs)"
    # Same without the role filter (fallback): recognised on the first refused definition, same message.
    $shared.evil = 'noroles'; $shared.calls.Clear()
    try { & $target -Workspace 'Other' -OutputPath $d22 6>$null | Out-Null; throw 'expected failure' }
    catch { Assert ($_.Exception.Message -match "Viewer access only to 'Other'") $_.Exception.Message }
    Assert (@(Get-Sigs | Where-Object { $_ -like 'POST *' }).Count -eq 1 -and -not (Test-Path -LiteralPath (Join-Path $d22 'Other'))) "fallback: one refused request, nothing written: $(Get-Sigs)"
    $shared.evil = ''
    'RUN 22 OK'

    # ---- Run 23: launcher with -AllWorkspaces -NewOnly, twice: the second run requests and analyses nothing new ----
    $shared.calls.Clear(); $shared.polls = 0
    $e23 = Join-Path $out 'a-export'; $s23 = Join-Path $out 'a-sources'
    & $launcher -AllWorkspaces -NewOnly -ExportPath $e23 -OutputPath $s23 3>$null 6>$null | Out-Null
    $csv23 = Join-Path $s23 'model_sources.csv'
    $rows23 = @(Import-Csv -LiteralPath $csv23 -Delimiter ((Get-Culture).TextInfo.ListSeparator[0]))
    Assert ((@($rows23.SemanticModel | Sort-Object -Unique) -join ' | ') -eq '[Fin] Budget | Sales_ EU_US_ | Sales_ EU_US_ (bbbbbbbb)') "first run analyses the exported models: $(@($rows23.SemanticModel | Sort-Object -Unique) -join ' | ')"
    Assert (-not (Get-Sigs | Where-Object { $_ -like '*00000000000b*' })) 'My workspace never read by the launcher'
    $bytes23 = [IO.File]::ReadAllBytes($csv23)
    $shared.calls.Clear()
    $text23 = & $launcher -AllWorkspaces -NewOnly -ExportPath $e23 -OutputPath $s23 3>$null *>&1 | Out-String
    Assert ($text23 -match 'Nothing new') "second run: nothing new to analyse: $text23"
    Assert ([Convert]::ToBase64String([IO.File]::ReadAllBytes($csv23)) -ceq [Convert]::ToBase64String($bytes23)) 'inventory untouched when nothing is new'
    Assert (-not (Get-Sigs | Where-Object { $_ -like "POST /v1/workspaces/$wsId/*" -and $_ -notmatch 'cccccccc|eeeeeeee' })) "second run requests no model already exported: $(Get-Sigs)"
    try { & $launcher -AllWorkspaces -Exclude '[test]*' -ExportPath $e23 -OutputPath $s23 6>$null | Out-Null; throw 'expected failure' }
    catch { Assert ($_.Exception.Message -match 'nothing to analyse') $_.Exception.Message }
    'RUN 23 OK'

    # ---- Run 24: the console display block is the same in the three scripts ----
    $blocks = foreach ($script in $target, $launcher, (Join-Path (Split-Path $target) 'Parse-SemanticModels.ps1')) {
        $text = [IO.File]::ReadAllText($script).Replace("`r`n", "`n")
        $begin = $text.IndexOf('console display (begin)')
        $end = $text.IndexOf('console display (end)')
        Assert ($begin -ge 0 -and $end -gt $begin) "display block markers in $script"
        $text.Substring($begin, $end - $begin)
    }
    Assert (@($blocks | Select-Object -Unique).Count -eq 1) 'console display block identical in the three scripts'
    'RUN 24 OK'

    # ---- Run 25: sign-in line from the token (decoded locally), -TenantId passed to az and validated ----
    $global:azCalls = 0; $shared.calls.Clear()
    $text25 = & $target -ListWorkspaces 6>&1 | Out-String
    Assert ($text25 -match "Signed in as tester@example\.test, tenant $tid") "signed-in line: $text25"
    Assert ($text25 -notmatch "FAKE|$fakePayload") 'the token is never printed'
    Assert ($global:azArgs -notcontains '--tenant') 'no --tenant without -TenantId'
    $null = & $target -ListWorkspaces -TenantId $tid 6>$null
    Assert (($global:azArgs -join ' ') -match "--tenant $tid") "-TenantId passed to az: $($global:azArgs -join ' ')"
    $warn25 = @(& $target -ListWorkspaces -TenantId '00000000-1111-2222-3333-444444444444' 3>&1 6>$null | Where-Object { $_ -is [Management.Automation.WarningRecord] })
    Assert (@($warn25 | Where-Object { "$_" -match "token is for tenant $tid" }).Count -eq 1) "tenant mismatch warned: $warn25"
    $shared.calls.Clear()
    foreach ($bad in 'contoso.onmicrosoft.com', '{abcdef01-2345-6789-abcd-ef0123456789}', "$tid --debug") {
        try { & $target -ListWorkspaces -TenantId $bad 6>$null | Out-Null; throw 'expected failure' }
        catch { Assert ($_.Exception.Message -match '-TenantId expects the tenant id') "$bad -> $($_.Exception.Message)" }
    }
    Assert (@(Get-Sigs).Count -eq 0) 'an invalid -TenantId sends nothing'
    $global:azMode = 'noexpiry'
    $text25 = & $target -ListWorkspaces 6>&1 | Out-String
    Assert ($text25 -match "Signed in as application app-test-id, tenant $tid") "application sign-in line: $text25"
    $global:azMode = 'ok'
    'RUN 25 OK'

    # ---- Run 26: az failures are explained (not installed is not tested: the real az would be found) ----
    $cases = [ordered]@{
        notloggedin = 'not signed in to Azure CLI.+Run: az login,'
        mfa         = 'AADSTS50076.+az login --scope https://api\.fabric\.microsoft\.com/\.default'
        aadsts      = 'AADSTS700082.+Run: az login,'
        nosub       = 'az login --allow-no-subscriptions'
    }
    foreach ($case in $cases.GetEnumerator()) {
        $global:azMode = $case.Key; $shared.calls.Clear()
        try { & $target -ListWorkspaces 6>$null | Out-Null; throw 'expected failure' }
        catch { Assert ($_.Exception.Message -match "(?s)$($case.Value)") "$($case.Key) -> $($_.Exception.Message)" }
        Assert (@(Get-Sigs).Count -eq 0) "$($case.Key): nothing sent without a token"
    }
    $global:azMode = 'notloggedin'
    try { & $target -ListWorkspaces -TenantId $tid 6>$null | Out-Null; throw 'expected failure' }
    catch { Assert ($_.Exception.Message -match "az login --tenant $tid") $_.Exception.Message }
    $global:azMode = 'leak'
    try { & $target -ListWorkspaces 6>$null | Out-Null; throw 'expected failure' }
    catch { Assert ($_.Exception.Message -notmatch "FAKE|$fakePayload" -and $_.Exception.Message -match '<hidden>') "a token in az's message is hidden: $($_.Exception.Message)" }
    $global:azMode = 'ok'
    'RUN 26 OK'

    # ---- Run 27: token lifetime: refreshed before expiry, local-time expiry parsed, 20 min fallback ----
    $global:azMode = 'expiring'; $global:azCalls = 0; $shared.calls.Clear()
    $null = & $target -ListWorkspaces 6>$null
    Assert ($global:azCalls -eq @(Get-Sigs).Count + 1) "a token 20 s from its end is renewed before every request: az x$($global:azCalls), $(@(Get-Sigs).Count) requests"
    foreach ($mode in 'localonly', 'noexpiry', 'ok') {
        $global:azMode = $mode; $global:azCalls = 0
        $null = & $target -Workspace $wsId -ListModels 6>$null
        Assert ($global:azCalls -eq 1) "$mode : one token for the whole run, az x$($global:azCalls)"
    }
    'RUN 27 OK'

    # ---- Run 28: HTTP 401 -> one new token, same request again; twice 401 -> clear error ----
    $shared.evil = '401once'; $shared.failures = 0; $global:azCalls = 0; $shared.calls.Clear()
    $r28 = @(& $target -ListWorkspaces 6>$null)
    Assert ($r28.Count -eq 3 -and $global:azCalls -eq 2) "401 then success after one refresh: $($r28.Count) workspaces, az x$($global:azCalls)"
    Assert ((@(Get-Sigs) -join ' ') -eq 'GET /v1/workspaces GET /v1/workspaces GET /v1/workspaces?page=2') "same request sent again once: $(Get-Sigs)"
    $shared.evil = '401always'; $global:azCalls = 0; $shared.calls.Clear()
    try { & $target -ListWorkspaces 6>$null | Out-Null; throw 'expected failure' }
    catch { Assert ($_.Exception.Message -match 'HTTP 401.+az login') $_.Exception.Message }
    Assert ($global:azCalls -eq 2 -and @(Get-Sigs).Count -eq 2) "one refresh, one retry only: az x$($global:azCalls), $(@(Get-Sigs).Count) requests"
    $shared.evil = ''
    'RUN 28 OK'

    # ---- Run 29: HTTP 5xx and dropped connections are retried with a back-off ----
    foreach ($mode in '503once', 'drop') {
        $shared.evil = $mode; $shared.failures = 0; $global:sleeps.Clear(); $shared.calls.Clear()
        $r29 = @(& $target -ListWorkspaces -Verbose 4>$null 6>$null)
        Assert ($r29.Count -eq 3 -and $shared.failures -eq 1 -and ($global:sleeps -join ',') -eq '5') "$mode : retried after 5 s, then listed: $($r29.Count) workspaces, sleeps $($global:sleeps -join ',')"
    }
    $shared.evil = ''
    'RUN 29 OK'

    # ---- Run 30: Retry-After honoured up to 600 s (number or HTTP date), with a visible warning for long waits ----
    foreach ($case in @(@('300', 300, 300), @('900', 600, 600), @('date', 85, 90))) {
        $shared.throttled = $false; $shared.retryAfter = $case[0]; $global:sleeps.Clear()
        $out30 = @(& $target -Workspace $wsId -ListModels -Verbose 4>$null 3>&1 6>$null)
        $w30 = @($out30 | Where-Object { $_ -is [Management.Automation.WarningRecord] })
        Assert (@($out30 | Where-Object { $_ -isnot [Management.Automation.WarningRecord] }).Count -eq 5) "$($case[0]) : listed after the wait"
        Assert ($global:sleeps.Count -eq 1 -and $global:sleeps[0] -ge $case[1] -and $global:sleeps[0] -le $case[2]) "Retry-After $($case[0]) -> waited $($global:sleeps -join ',') s"
        Assert (@($w30 | Where-Object { "$_" -match 'throttling' }).Count -eq 1) "long wait announced: $w30"
    }
    $shared.retryAfter = '7'
    'RUN 30 OK'

    # ---- Run 31: a workspace whose models cannot be listed is reported, the others are exported ----
    $shared.evil = 'listfail'; $shared.calls.Clear(); $shared.polls = 0
    $d31 = Join-Path $out 'run31'
    $r31 = @(& $target -AllWorkspaces -OutputPath $d31 6>$null)
    $fail31 = @($r31 | Where-Object Workspace -EQ 'Flaky')
    Assert ($fail31.Count -eq 1 -and $fail31[0].Status -eq 'Failed' -and -not $fail31[0].ModelId -and $fail31[0].Error -match 'Could not list the semantic models: HTTP 404') "listing failure row: $($fail31 | Out-String)"
    Assert (@($r31 | Where-Object WorkspaceId -EQ $wsId).Count -eq 5 -and (Test-Path -LiteralPath (Join-Path $d31 "$wsName\_export_report.csv"))) 'the other workspace is still exported'
    Assert (-not (Test-Path -LiteralPath (Join-Path $d31 'Flaky'))) 'nothing written for the workspace that could not be listed'
    $m31 = @(& $target -AllWorkspaces -ListModels 3>$null 6>$null)
    Assert ($m31.Count -eq 7) "-ListModels goes on past the failing workspace: $($m31.Count)"
    $shared.evil = ''
    'RUN 31 OK'

    # ---- Run 32: colliding folder names do not depend on the listing order; long names are cut ----
    $shared.evil = 'reverse'; $shared.polls = 0
    $r32 = @(& $target -Workspace $wsId -Model 'Sales*' -OutputPath (Join-Path $out 'run32') 6>$null)
    Assert ((($r32 | Where-Object Model -EQ 'Sales: EU/US?').Folder).EndsWith('\Sales_ EU_US_.SemanticModel') -and (($r32 | Where-Object Model -EQ 'Sales_ EU_US_').Folder).EndsWith('\Sales_ EU_US_ (bbbbbbbb).SemanticModel')) "same folders whatever the order: $($r32 | Format-Table Model, Folder | Out-String)"
    $shared.evil = 'long'
    $r32 = @(& $target -Workspace $wsId -OutputPath (Join-Path $out 'run32') 6>$null)
    $leaves = @($r32 | ForEach-Object { Split-Path $_.Folder -Leaf })
    Assert ($r32.Count -eq 2 -and -not ($r32 | Where-Object Status -NE 'Exported')) "long names exported: $($r32 | Out-String)"
    Assert (($leaves -join '|') -ceq ("$('L' * 80) (aaaaaaaa).SemanticModel|$('L' * 80) (dddddddd).SemanticModel")) "cut at 80 characters plus the id: $($leaves -join '|')"
    $shared.evil = ''
    'RUN 32 OK'

    # ---- Run 33: -NewOnly trusts a model folder only when the previous report ties it to the same model id ----
    $report33 = Join-Path $d19 "$wsName\_export_report.csv"
    $rows33 = @(Import-Csv -LiteralPath $report33)
    ($rows33 | Where-Object Model -EQ '[Fin] Budget').ModelId = '12345678-0000-0000-0000-000000000000'
    $rows33 | Export-Csv -LiteralPath $report33 -NoTypeInformation -Encoding utf8
    $shared.calls.Clear()
    $r33 = @(& $target -Workspace $wsId -NewOnly -Model '[fin]*', 'Sales*' -OutputPath $d19 6>$null)
    $st33 = @{}; $r33 | ForEach-Object { $st33[$_.Model] = $_.Status }
    Assert ($st33['[Fin] Budget'] -eq 'Exported' -and $st33['Sales: EU/US?'] -eq 'AlreadyExported') "mismatched id exported again, the others kept: $($r33 | ForEach-Object { "$($_.Model)=$($_.Status)" })"
    Assert (@(Get-Sigs | Where-Object { $_ -like 'POST *' }).Count -eq 1 -and (Get-Sigs | Where-Object { $_ -like 'POST *dddddddd*' })) "only the mismatched model asked: $(Get-Sigs)"
    $csv33 = @(Import-Csv -LiteralPath $report33)
    Assert ($csv33.Count -eq 5 -and ($csv33 | Where-Object Model -EQ '[Fin] Budget').ModelId -eq 'dddddddd-0000-0000-0000-000000000004') "report repaired and merged: $($csv33 | Format-Table Model, ModelId, Status | Out-String)"
    'RUN 33 OK'

    # ---- Run 34: -NewOnly in a workspace found Viewer-only by the fallback: nothing of it is counted, not even kept models ----
    $d34 = Join-Path $out 'run34'; $viewed = Join-Path $d34 'Other\Viewed 1.SemanticModel\definition'
    New-Item -ItemType Directory -Path $viewed | Out-Null
    Set-Content -LiteralPath (Join-Path $d34 '.semantic-model-tools') -Value 'test'
    Set-Content -LiteralPath (Join-Path $viewed 'model.tmdl') -Value 'model Model'
    [pscustomobject]@{ Workspace = 'Other'; WorkspaceId = '00000000-0000-0000-0000-00000000000a'; Model = 'Viewed 1'; ModelId = 'ffffffff-0000-0000-0000-000000000006'; Status = 'Exported'; Files = 1; TmdlFiles = 1; Folder = (Split-Path $viewed); Error = ''; ExportedAtUtc = '2026-01-01T00:00:00Z' } | Export-Csv -LiteralPath (Join-Path $d34 'Other\_export_report.csv') -NoTypeInformation -Encoding utf8
    $shared.evil = 'noroles'; $shared.calls.Clear()
    $r34 = @(& $target -AllWorkspaces -NewOnly -OutputPath $d34 6>$null)
    Assert (-not ($r34 | Where-Object Workspace -EQ 'Other') -and @(Get-Sigs | Where-Object { $_ -like 'POST */00000000-0000-0000-0000-00000000000a/*' }).Count -eq 1) "Viewer-only workspace left out entirely: $($r34 | Format-Table Workspace, Model, Status | Out-String)"
    $shared.evil = ''
    'RUN 34 OK'

    # ---- Run 35: OneDrive. COVERED: the warning, a junction refused inside OneDrive too, a drive-root variable ignored.
    # NOT COVERED: a real files-on-demand folder (ReparsePoint attribute, no link) being accepted - it cannot be created here.
    $oneDriveBefore = $env:OneDrive
    try {
        $root35 = Join-Path $out 'od35\OneDrive - Test'; New-Item -ItemType Directory -Path $root35 | Out-Null
        $env:OneDrive = $root35
        $w35 = @(& $target -Workspace $wsId -Model '[fin]*' -OutputPath (Join-Path $root35 'exports') 3>&1 6>$null)
        Assert (@($w35 | Where-Object { $_ -isnot [Management.Automation.WarningRecord] -and $_.Status -eq 'Exported' }).Count -eq 1) 'export inside OneDrive works'
        Assert (@($w35 | Where-Object { $_ -is [Management.Automation.WarningRecord] -and "$_" -match 'is in OneDrive.+synced to OneDrive' }).Count -eq 1) "OneDrive warning: $($w35 | Where-Object { $_ -is [Management.Automation.WarningRecord] })"
        $elsewhere = Join-Path $out 'od35\elsewhere'; New-Item -ItemType Directory -Path $elsewhere | Out-Null
        New-Item -ItemType Junction -Path (Join-Path $root35 'link') -Target $elsewhere | Out-Null
        $shared.calls.Clear()
        try { & $target -Workspace $wsId -OutputPath (Join-Path $root35 'link\exports') 6>$null | Out-Null; throw 'expected refusal' }
        catch { Assert ($_.Exception.Message -match 'is a Junction to .+-OutputPath C:\\') $_.Exception.Message }
        Assert (@(Get-Sigs).Count -eq 0 -and -not @(Get-ChildItem -LiteralPath $elsewhere -Force).Count) 'refused before any request, nothing written through the junction'
        $env:OneDrive = [IO.Path]::GetPathRoot($out)
        $w35 = @(& $target -Workspace $wsId -Model '[fin]*' -OutputPath (Join-Path $out 'run35') 3>&1 6>$null | Where-Object { $_ -is [Management.Automation.WarningRecord] })
        Assert (-not ($w35 | Where-Object { "$_" -match 'OneDrive' })) 'a OneDrive variable naming a drive root is ignored'
    } finally { $env:OneDrive = $oneDriveBefore }
    'RUN 35 OK'

    # ---- Run 36: a listing answered with anything but 200 is an error, never an empty list ----
    $shared.evil = 'status204'
    $r36 = @(& $target -Workspace $wsId -OutputPath (Join-Path $out 'run36') 6>$null)
    Assert ($r36.Count -eq 1 -and $r36[0].Status -eq 'Failed' -and $r36[0].Error -match 'HTTP 204') "HTTP 204 on a listing is a failure: $($r36 | Out-String)"
    $shared.evil = ''
    'RUN 36 OK'

    # ---- Run 37: an EMPTY answer of the role filter while workspaces are visible is not believed (first-403 fallback) ----
    $shared.evil = 'emptyroles'; $shared.calls.Clear(); $shared.polls = 0
    $d37 = Join-Path $out 'run37'
    $r37 = @(& $target -AllWorkspaces -OutputPath $d37 6>$null)
    Assert (@($r37 | Where-Object { $_.WorkspaceId -eq $wsId -and $_.Status -eq 'Exported' }).Count -eq 3 -and -not ($r37 | Where-Object Workspace -EQ 'Other')) "the test workspace is still exported: $($r37 | Format-Table Workspace, Model, Status | Out-String)"
    Assert (@(Get-Sigs | Where-Object { $_ -like 'POST */00000000-0000-0000-0000-00000000000a/*' }).Count -eq 1 -and -not (Test-Path -LiteralPath (Join-Path $d37 'Other'))) "Viewer-only found by its first refused definition: $(Get-Sigs)"
    $shared.evil = ''
    'RUN 37 OK'

    # ---- Run 38: a model whose folder name changed: its earlier folder (same model id in the report) is removed after the
    # new export; never a folder that a current model uses, never through a junction ----
    $d38 = Join-Path $out 'run38'; $ws38 = Join-Path $d38 $wsName; $shared.polls = 0
    $null = @(& $target -Workspace $wsId -Model 'Sales*', '[fin]*' -OutputPath $d38 6>$null)
    $old38 = Join-Path $ws38 'Sales EU US old long name.SemanticModel'; New-Item -ItemType Directory -Path (Join-Path $old38 'definition') | Out-Null
    Set-Content -LiteralPath (Join-Path $old38 'definition\model.tmdl') -Value 'model Old'
    $outside38 = Join-Path $out 'outside38'; New-Item -ItemType Directory -Path $outside38 | Out-Null
    Set-Content -LiteralPath (Join-Path $outside38 'keep.txt') -Value 'keep'
    New-Item -ItemType Junction -Path (Join-Path $ws38 'Linked.SemanticModel') -Target $outside38 | Out-Null
    $report38 = Join-Path $ws38 '_export_report.csv'
    $rows38 = @(Import-Csv -LiteralPath $report38)
    $sales38 = $rows38 | Where-Object Model -EQ 'Sales: EU/US?'
    $extra = @(
        [pscustomobject]@{ Workspace = $wsName; WorkspaceId = $wsId; Model = 'Sales: EU/US?'; ModelId = $sales38.ModelId; Status = 'Exported'; Files = 3; TmdlFiles = 2; Folder = $old38; Error = ''; ExportedAtUtc = '2026-01-01T00:00:00Z' }
        [pscustomobject]@{ Workspace = $wsName; WorkspaceId = $wsId; Model = 'Sales: EU/US?'; ModelId = $sales38.ModelId; Status = 'Exported'; Files = 3; TmdlFiles = 2; Folder = (Join-Path $ws38 'Linked.SemanticModel'); Error = ''; ExportedAtUtc = '2026-01-01T00:00:00Z' }
        # Points to the folder of ANOTHER current model: never deleted.
        [pscustomobject]@{ Workspace = $wsName; WorkspaceId = $wsId; Model = '[Fin] Budget'; ModelId = 'dddddddd-0000-0000-0000-000000000004'; Status = 'Exported'; Files = 3; TmdlFiles = 2; Folder = (Join-Path $ws38 'Sales_ EU_US_ (bbbbbbbb).SemanticModel'); Error = ''; ExportedAtUtc = '2026-01-01T00:00:00Z' }
    )
    @(@($rows38 | Where-Object Model -NE '[Fin] Budget') + $extra) | Export-Csv -LiteralPath $report38 -NoTypeInformation -Encoding utf8
    $w38 = @(& $target -Workspace $wsId -Model 'Sales: EU*', '[fin]*' -OutputPath $d38 3>&1 6>$null)
    Assert (@($w38 | Where-Object { $_ -isnot [Management.Automation.WarningRecord] -and $_.Status -eq 'Exported' }).Count -eq 2) "re-export: $($w38 | Out-String)"
    Assert (-not (Test-Path -LiteralPath $old38)) 'the earlier folder of the same model is removed'
    Assert (Test-Path -LiteralPath (Join-Path $ws38 'Sales_ EU_US_ (bbbbbbbb).SemanticModel\definition\model.tmdl')) 'the folder of another current model is kept'
    Assert ((Test-Path -LiteralPath (Join-Path $outside38 'keep.txt')) -and (Test-Path -LiteralPath (Join-Path $ws38 'Linked.SemanticModel'))) 'nothing deleted through a junction'
    Assert (@($w38 | Where-Object { $_ -is [Management.Automation.WarningRecord] -and "$_" -match 'Linked\.SemanticModel.+kept' }).Count -eq 1) "the junction is reported: $($w38 | Where-Object { $_ -is [Management.Automation.WarningRecord] })"
    [IO.Directory]::Delete((Join-Path $ws38 'Linked.SemanticModel'))   # the junction itself, not its target
    'RUN 38 OK'

    # ---- Run 39: a refusal of the script's guards inside a per-workspace catch still stops the whole run ----
    $shared.evil = 'id'; $shared.calls.Clear()
    try { & $target -Workspace $wsId -ListModels 6>$null 3>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'Unexpected item id in Fabric response') "bad model id while listing models: $($_.Exception.Message)" }
    $shared.evil = 'rolesid'; $shared.calls.Clear()
    try { & $target -AllWorkspaces -OutputPath (Join-Path $out 'run39') 6>$null 3>$null | Out-Null; throw 'expected refusal' }
    catch { Assert ($_.Exception.Message -match 'Unexpected item id in Fabric response') "bad id in the role listing: $($_.Exception.Message)" }
    Assert (-not (Get-Sigs | Where-Object { $_ -like 'POST *' -or $_ -like '*semanticModels*' })) "stopped before any model listing or definition: $(Get-Sigs)"
    $shared.evil = ''
    'RUN 39 OK'

    # ---- Run 40: HTTP 429 without end: at most 20 minutes of waiting for one request, then a clear failure ----
    $shared.evil = 'throttle-always'; $global:sleeps.Clear()
    $r40 = @(& $target -Workspace $wsId -OutputPath (Join-Path $out 'run40') -Verbose 4>$null 3>$null 6>$null)
    Assert ($r40.Count -eq 1 -and $r40[0].Status -eq 'Failed' -and $r40[0].Error -match 'kept throttling') "throttling gives up: $($r40 | Out-String)"
    Assert (($global:sleeps -join ',') -eq '600,600') "waited 2 x 600 s, no more: $($global:sleeps -join ',')"
    $shared.evil = ''
    'RUN 40 OK'

    # ---- Run 41: an HTTP answer turned into an exception (a refused redirect) is not a network error to retry ----
    & {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($target, [ref]$null, [ref]$null)
        . ([scriptblock]::Create($ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-TransportError' }, $true).Extent.Text))
        $redirect = [Microsoft.PowerShell.Commands.HttpResponseException]::new('302', [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]::Found))
        Assert (-not (Test-TransportError $redirect)) 'HttpResponseException is not retried'
        Assert ((Test-TransportError ([Net.Http.HttpRequestException]::new('connection reset'))) -and -not (Test-TransportError ([Security.SecurityException]::new('x')))) 'network errors still retried, refusals never'
    }
    'RUN 41 OK'

    # ---- Run 42: the "Next" line can be pasted even when the output path holds a quote ----
    $d42 = Join-Path $out "it's here"
    $text42 = & $target -Workspace $wsId -Model '[fin]*' -OutputPath $d42 6>&1 | Out-String
    Assert ($text42.Contains("-ExportPath '$($d42.Replace("'", "''"))'")) "quote doubled in the Next line: $text42"
    'RUN 42 OK'
    'ALL OFFLINE TESTS PASSED'
}
finally {
    $shared.stop = $true
    $server | Wait-Job -Timeout 5 | Out-Null; $server | Remove-Job -Force
    if ($shared.error) { "listener error: $($shared.error)" }
    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Recurse -Force }
}

# Differential test of the two SQL analysers of Parse-SemanticModels.ps1 (Microsoft ScriptDom vs built-in scanner).
# Each query becomes the native query of one synthetic table; the script is run twice (with / without ScriptDom);
# both results are compared with each other AND with a hand-written expectation.
$ErrorActionPreference = 'Stop'
$target = Join-Path (Split-Path $PSScriptRoot -Parent) 'Parse-SemanticModels.ps1'
$root = Join-Path ([IO.Path]::GetTempPath()) "sqltest_$(Get-Random)"
$tables = Join-Path $root 'export\WS\Battery.SemanticModel\definition\tables'
New-Item -ItemType Directory -Path $tables -Force | Out-Null

# name = @( sql , expected objects in document order )      -- all of these are valid T-SQL
$tsql = [ordered]@{
    UserExample = @("/* Query Start */`nSELECT distinct vis.[crm_account_id] as [Partner ID]`nFROM dbo.crm_visit  vis`n`nUNION`n`nSELECT distinct con.crm_account_id as [Partner ID]`nFROM [dbo].[crm_contract] con`n/* Query End */", 'dbo.crm_visit|dbo.crm_contract')
    Nasty       = @(";WITH cte AS (SELECT * FROM [dbo].[Orders] o WHERE note = 'FROM fake.table' /* JOIN commented.out */),`n c2 (a) AS (SELECT 1)`nSELECT c.*, DATEPART(year, o.d), TRIM('x' FROM c.n) FROM cte c`nJOIN sales.Customers AS cu WITH (NOLOCK) ON 1=1 -- JOIN also.commented x`nLEFT JOIN [Other Db]..[Country] k ON k.id = cu.id`nCROSS APPLY dbo.fn_Split(c.x, ',') s, ref.CommaJoined cj`nWHERE EXISTS (SELECT 1 FROM c2 JOIN stg.[Weird.Name] w ON 1=1);`nEXEC rpt.usp_Load @p = 1;", 'dbo.Orders|sales.Customers|Other Db:Country|dbo.fn_Split|ref.CommaJoined|stg.[Weird.Name]|rpt.usp_Load')
    Derived     = @("SELECT * FROM (SELECT * FROM dbo.A) a, dbo.B b WHERE EXISTS (SELECT 1 FROM dbo.C c WHERE c.id = a.id)", 'dbo.A|dbo.B|dbo.C')
    NestedCmt   = @("/* a /* nested FROM dbo.No1 */ still comment JOIN dbo.No2 */ SELECT * FROM dbo.D", 'dbo.D')
    Quoted      = @('SELECT * FROM [my db].[my schema].[my.table] t INNER JOIN "dbo"."E" e ON 1=1', 'my db:[my schema].[my.table]|dbo.E')
    SetOps      = @("SELECT a FROM dbo.E1 x WITH (NOLOCK) UNION ALL SELECT a FROM dbo.E2 INTERSECT SELECT a FROM dbo.E3 OPTION (RECOMPILE)", 'dbo.E1|dbo.E2|dbo.E3')
    Pivot       = @("SELECT * FROM (SELECT a, b FROM dbo.F) s PIVOT (SUM(b) FOR a IN ([x],[y])) p", 'dbo.F')
    OpenJson    = @("SELECT g.id, j.id FROM dbo.G AS g CROSS APPLY OPENJSON(g.j) WITH (id int '$.id') j", 'dbo.G')
    Strings     = @("SELECT 'x -- FROM dbo.Fake' AS s, N'it''s FROM dbo.Fake3' AS t FROM dbo.H WHERE note LIKE '%JOIN dbo.Fake2%'", 'dbo.H')
    TempAndVar  = @("SELECT x INTO #tmp FROM dbo.I0; SELECT * FROM #tmp t JOIN @tv v ON 1=1 JOIN dbo.I i ON 1=1 JOIN ##global g ON 1=1", 'dbo.I0|dbo.I')
    DistinctFrom = @("SELECT * FROM dbo.J j WHERE j.a IS DISTINCT FROM j.b AND j.c IS NOT DISTINCT FROM j.d", 'dbo.J')
    Batch       = @("DECLARE @d date = GETDATE(); SELECT * FROM dbo.K WHERE d = @d;", 'dbo.K')
    Tvf         = @("SELECT * FROM dbo.fnTable(1, 'x') f JOIN sys.objects o ON 1=1", 'dbo.fnTable|sys.objects')
    FourPart    = @("SELECT * FROM srv9.db9.dbo.L", 'srv9:db9:dbo.L')
    Window      = @("SELECT SUM(x) OVER (PARTITION BY a ORDER BY b ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS s FROM dbo.M", 'dbo.M')
    Joins       = @("SELECT TOP 10 * FROM dbo.N n LEFT OUTER JOIN dbo.O o ON n.id=o.id FULL JOIN dbo.P p ON 1=1 WHERE n.id IN (SELECT id FROM dbo.Q)", 'dbo.N|dbo.O|dbo.P|dbo.Q')
    NoTable     = @("SELECT NULL AS ColumnA WHERE 1=2", '')
    CteOnly     = @("WITH r AS (SELECT 1 AS n UNION ALL SELECT n + 1 FROM r WHERE n < 5) SELECT * FROM r", '')
    SameTwice   = @("SELECT * FROM dbo.R a JOIN [dbo].[R] b ON a.id = b.parent", 'dbo.R')
    ForJson     = @("SELECT a, b FROM dbo.U u JOIN dbo.V v ON u.id = v.id AND u.k IN (1, 2), dbo.W w GROUP BY a, b FOR JSON PATH, ROOT('x')", 'dbo.U|dbo.V|dbo.W')
    ContainsTable = @("SELECT * FROM CONTAINSTABLE(dbo.Documents, *, 'term') AS f", 'dbo.Documents')
    ChangeTable = @("SELECT * FROM CHANGETABLE(CHANGES dbo.T, 0) AS c", 'dbo.T')
    UpdateAlias = @("UPDATE t SET v = 1 FROM dbo.Real AS t", 'dbo.Real')
    CteBatch    = @("WITH x AS (SELECT * FROM dbo.A) SELECT * FROM x; SELECT * FROM x;", 'dbo.A|x')
    ParenJoin   = @("SELECT * FROM (dbo.A AS a JOIN dbo.B AS b ON a.id = b.id) JOIN (dbo.C c JOIN dbo.D d ON c.id = d.id) ON a.id = c.id", 'dbo.A|dbo.B|dbo.C|dbo.D')
    SystemTime  = @("SELECT * FROM dbo.A FOR SYSTEM_TIME ALL AS a, dbo.B AS b", 'dbo.A|dbo.B')
    Merge       = @("MERGE dbo.T AS t USING dbo.S AS s ON t.id = s.id WHEN MATCHED THEN UPDATE SET v = s.v;", 'dbo.T|dbo.S')
    OldHint     = @("SELECT * FROM dbo.H1 (NOLOCK) JOIN dbo.H2 h2 (NOLOCK) ON 1=1", 'dbo.H1|dbo.H2')
    Substring   = @("SELECT SUBSTRING(name, 1, 3) AS s, CAST(x AS varchar(10)) FROM dbo.S WHERE EXISTS (SELECT 1 FROM dbo.T)", 'dbo.S|dbo.T')
    OdbcEscape  = @("SELECT * FROM dbo.E WHERE d > {d '2024-01-01'} AND t < {fn NOW()}", 'dbo.E')
    Dual        = @("SELECT 1 AS one FROM dual", '')
}
# not T-SQL: only the built-in scanner can read these (ScriptDom must fail cleanly and hand over)
$other = [ordered]@{
    Databricks  = @('SELECT * FROM `main`.`gold`.`fact_tickets` t JOIN main.gold.dim_date d ON 1=1 LIMIT 10', 'main:gold.fact_tickets|main:gold.dim_date')
    Postgres    = @("SELECT EXTRACT(YEAR FROM o.d) AS y, o.amount::numeric FROM public.orders o JOIN LATERAL (SELECT 1) x ON true", 'public.orders')
    Snowflake   = @("SELECT f.value FROM db1.sch.t, LATERAL FLATTEN(input => t.arr) f, TABLE(SPLIT_TO_TABLE(t.s, ',')) s LIMIT 5", 'db1:sch.t')
    OracleDual  = @("SELECT SYSDATE FROM dual WHERE ROWNUM <= 1 CONNECT BY LEVEL <= 2", '')
}
$embedded = [ordered]@{
    OpenQuery   = @("SELECT * FROM OPENQUERY(RemoteServer, 'SELECT * FROM dbo.A')", 'dbo.A')
    ExecString  = @("EXEC ('SELECT * FROM dbo.A a JOIN dbo.B b ON 1=1')", 'dbo.A|dbo.B')
    ExecVariable = @("DECLARE @s nvarchar(max) = N'SELECT 1'; EXEC (@s)", '')
}
$i = 0
foreach ($set in $tsql, $other, $embedded) {
    foreach ($name in $set.Keys) {
        $i++
        $m = $set[$name][0].Replace('"', '""')
        # Like Fabric does, every line of the expression (continuation lines of a multi-line string included) carries the base indentation.
        $expression = "let`n    sqlQuery = `"$m`",`n    Source = Sql.Database(`"srv`", `"db`", [Query=sqlQuery])`nin`n    Source"
        $indented = (($expression -split "`n") | ForEach-Object { "`t`t`t`t$_" }) -join "`n"
        [IO.File]::WriteAllText((Join-Path $tables "$name.tmdl"), "table $name`n`n`tpartition p = m`n`t`tmode: import`n`t`tsource =`n$indented`n")
    }
}
# not SQL at all (a DAX query sent through a SQL connector by mistake): checked separately below
[IO.File]::WriteAllText((Join-Path $tables 'DaxText.tmdl'), "table DaxText`n`n`tpartition p = m`n`t`tmode: import`n`t`tsource = Sql.Database(`"srv`", `"db`", [Query=`"EVALUATE 'Sales'`"])`n")
function Assert($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
try {
    Push-Location $root
    $withDom = @(& $target -ExportPath .\export -OutputPath .\out1 -Delimiter ';' -PassThru 6>$null)
    $scanner = @(& $target -ExportPath .\export -OutputPath .\out2 -Delimiter ';' -PassThru -NoTSqlParser 6>$null)
    # Which analyser read a query is written in the header of the resolved SQL file (it is no longer a column of the inventory).
    $analysedBy = { param($output, $table) [regex]::Match([IO.File]::ReadAllText((Join-Path $root "$output\WS\Battery\sql\$table.sql")), '(?m)^-- Analysed by : (.*?)\r?$').Groups[1].Value }
    Assert (@($tsql.Keys | Where-Object { (& $analysedBy 'out1' $_) -eq 'T-SQL parser' }).Count -gt 0) 'ScriptDom was not loaded on this machine: the differential test is meaningless'
    Assert (-not @($tsql.Keys | Where-Object { (& $analysedBy 'out2' $_) -ne 'built-in scanner' }).Count) '-NoTSqlParser sends every query to the built-in scanner'
    foreach ($rows in @($withDom, $scanner)) {
        foreach ($row in $rows) {
            Assert ($row.QueryFile -eq "WS\Battery\power-query\$($row.Table).pq" -and $row.SqlFile -eq "WS\Battery\sql\$($row.Table).sql") "query paths for $($row.Table)"
        }
    }
    foreach ($name in $tsql.Keys) {
        $expectedSql = $tsql[$name][0].Replace("`r`n", "`n").Replace("`n", "`r`n")
        foreach ($output in @('out1', 'out2')) {
            $written = [IO.File]::ReadAllText((Join-Path $root "$output\WS\Battery\sql\$name.sql"))
            Assert ($written.EndsWith($expectedSql + "`r`n")) "exact resolved SQL preserved for $name with $output"
        }
    }
    # Objects are rendered as [server:][database:]schema.object, relative to the connection (srv / db): a database or
    # server named in front of the object must have moved to the Database / Server columns.
    $objects = { param($rows, $table) (@($rows | Where-Object Table -EQ $table | Where-Object Object | ForEach-Object { $o = $_.Object; if ($_.Database -cne 'db') { $o = "$($_.Database):$o" }; if ($_.Server -cne 'srv') { $o = "$($_.Server):$o" }; $o }) -join '|') }
    $bad = 0
    foreach ($name in $tsql.Keys) {
        $expected = $tsql[$name][1]; $a = & $objects $withDom $name; $b = & $objects $scanner $name
        $by = & $analysedBy 'out1' $name
        $review = @($withDom | Where-Object Table -EQ $name | Where-Object NeedsReview).Count
        $ok = ($a -ceq $expected) -and ($b -ceq $expected) -and ($by -eq 'T-SQL parser') -and $review -eq 0
        if (-not $ok) { $bad++ }
        "{0,-5} {1,-13} expected=<{2}>{3}{4}{5}" -f $(if ($ok) { 'OK' } else { 'FAIL' }), $name, $expected, $(if ($a -cne $expected) { "`n        ScriptDom=<$a>" }), $(if ($b -cne $expected) { "`n        scanner  =<$b>" }), $(if ($by -ne 'T-SQL parser' -or $review) { "`n        by=$by review=$review" })
    }
    foreach ($name in $other.Keys) {
        $expected = $other[$name][1]; $a = & $objects $withDom $name; $b = & $objects $scanner $name
        $by = & $analysedBy 'out1' $name
        $ok = ($a -ceq $expected) -and ($b -ceq $expected) -and ($by -eq 'built-in scanner')
        if (-not $ok) { $bad++ }
        "{0,-5} {1,-13} expected=<{2}> (not T-SQL -> scanner){3}" -f $(if ($ok) { 'OK' } else { 'FAIL' }), $name, $expected, $(if (-not $ok) { "`n        with ScriptDom=<$a> scanner=<$b> by=$by" })
    }
    # SQL sent elsewhere as text: the T-SQL parser reads a CONSTANT inner text; the scanner cannot, and must say so.
    foreach ($name in @($embedded.Keys | Where-Object { $_ -ne 'ExecVariable' })) {
        $a = & $objects $withDom $name; $b = & $objects $scanner $name
        $reviewA = @($withDom | Where-Object Table -EQ $name | Where-Object NeedsReview).Count; $reviewB = @($scanner | Where-Object Table -EQ $name | Where-Object NeedsReview).Count
        $ok = ($a -ceq $embedded[$name][1]) -and $reviewA -eq 0 -and ($b -ceq '') -and $reviewB -gt 0
        if (-not $ok) { $bad++ }
        "{0,-5} {1,-13} expected=<{2}> with ScriptDom, and a row TO REVIEW with the scanner{3}" -f $(if ($ok) { 'OK' } else { 'FAIL' }), $name, $embedded[$name][1], $(if (-not $ok) { "`n        ScriptDom=<$a> review=$reviewA | scanner=<$b> review=$reviewB" })
    }
    $variable = @($withDom | Where-Object Table -EQ 'ExecVariable')
    Assert ($variable.Count -eq 1 -and $variable[0].NeedsReview -eq 'yes') "EXEC of a variable can hide anything: never certain: $($variable | Out-String)"
    # DAX (or any text that is not SQL) through a SQL connector: T-SQL reads "EVALUATE 'Sales'" as a call of a procedure
    # EVALUATE without EXEC. Never certain, with either analyser.
    foreach ($rows in @($withDom, $scanner)) {
        $dax = @($rows | Where-Object Table -EQ 'DaxText')
        Assert ($dax.Count -ge 1 -and -not @($dax | Where-Object { -not $_.NeedsReview }).Count) "text that is not SQL is never certain: $($dax | Select-Object Object, ObjectType, NeedsReview | Out-String)"
    }
    $noTable = @($withDom | Where-Object Table -EQ 'NoTable')
    Assert ($noTable.Count -eq 1 -and -not $noTable[0].Object -and -not $noTable[0].NeedsReview) "measure-table dummy query is a certainty: $($noTable | Out-String)"
    Assert (-not @($withDom | Where-Object { $_.Parameters -match '(?i)select|sqlQuery' }).Count) 'the SQL text must never land in the Parameters column'
    Assert ($bad -eq 0) "$bad query(ies) disagree"
    "ALL SQL PARSER TESTS PASSED ($($tsql.Count + $other.Count + $embedded.Count) queries, both analysers)"
}
finally { Pop-Location; if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction Ignore } }

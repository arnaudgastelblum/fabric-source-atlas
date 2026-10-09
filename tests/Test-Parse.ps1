# Offline test of Parse-SemanticModels.ps1 on synthetic TMDL fixtures ('→' stands for a tab).
$ErrorActionPreference = 'Stop'
$target = Join-Path (Split-Path $PSScriptRoot -Parent) 'Parse-SemanticModels.ps1'
$root = Join-Path ([IO.Path]::GetTempPath()) "srctest_$(Get-Random)"
$export = Join-Path $root 'export'; $out = Join-Path $root 'out'
$def = Join-Path $export '[WS] Test\Fixture Model.SemanticModel\definition'
New-Item -ItemType Directory -Path (Join-Path $def 'tables') -Force | Out-Null
function Put([string]$Relative, [string]$Text) { [IO.File]::WriteAllText((Join-Path $def $Relative), $Text.Replace('→', "`t").Replace("`r`n", "`n")) }
Set-Content -LiteralPath (Join-Path (Split-Path $def) '.platform') -Value '{"metadata":{"type":"SemanticModel","displayName":"Fixture: Model?"}}'

Put 'expressions.tmdl' @'
expression ParamSchema = "mart_x" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]
→lineageTag: 0

expression BaseUrl = "https://wabi.example.net" meta [IsParameterQuery=true, Type="Text"]

expression WorkspaceId = "1111" meta [IsParameterQuery=true, Type="Text"]
→lineageTag: 1

expression ServerName = "srv.database.windows.net" meta [IsParameterQuery=true, Type="Text"]

expression ConnStr = "dsn=Hive2;pwd=Hunter2;database=raw2" meta [IsParameterQuery=true, Type="Text"]

expression 'Shared Conn' =
→→let
→→    Source = Sql.Database(ServerName, "SharedDb")
→→in
→→    Source
→lineageTag: 2
→queryGroup: Conn

expression 'Shared Lookup' =
→→let
→→    Source = Sql.Database(ServerName, "Sales", [Query="select * from dbo.X"])
→→in
→→    Source
→lineageTag: 8

expression 'DirectQuery to AS - Sales' =
→→let
→→    Source = AnalysisServices.Database("powerbi://api.powerbi.com/v1.0/myorg/Shared WS", "Sales Model")
→→in
→→    Source
→lineageTag: 9

expression DatabaseQuery =
→→let
→→    database = Sql.Database("abc.datawarehouse.fabric.microsoft.com", "lakehouse_gold")
→→in
→→    database
→lineageTag: 3
'@

Put 'tables\NavParam.tmdl' @'
table NavParam
→lineageTag: x

→column A
→→dataType: string

→partition NavParam = m
→→mode: import
→→source =
→→→→let
→→→→    // Sql.Database("commented.out", "nope") must be ignored
→→→→    Source = Sql.Database("host1.datawarehouse.fabric.microsoft.com", "dwh"),
→→→→    t = Source{[Schema=""&  ParamSchema &"",Item="dim_customer"]}[Data]
→→→→in
→→→→    t

→annotation PBI_ResultType = Table
'@

Put 'tables\NativeSql.tmdl' @'
table 'Native''s SQL'

→partition p = m
→→mode: directQuery
→→source =
→→→→let
→→→→    Source = Sql.Database(ServerName, "Sales", [Query="WITH cte AS (SELECT * FROM [dbo].[Orders] o)#(lf)SELECT c.*, EXTRACT(YEAR FROM o.d) FROM cte c#(lf)JOIN sales.Customers AS cu ON 1=1 -- JOIN commented.out x#(lf)LEFT JOIN ""ref"".""Country"" k ON k.id = cu.id WHERE note = 'FROM fake.table'", CreateNavigationProperties=false])
→→→→in
→→→→    Source
'@

Put 'tables\NativeQuery2.tmdl' @'
table NativeQuery2

→partition p = m
→→mode: import
→→source = ```
→→→→let
→→→→    Source = Sql.Databases("srv2"),
→→→→    r = Value.NativeQuery(Source{[Name="DbFromNav"]}[Data], "select * from dbo.FactSales f join dbo.DimDate d on 1=1; exec dbo.usp_Load", null, [EnableFolding=true])
→→→→in
→→→→    r
→→→→```
'@

Put 'tables\ThreePart.tmdl' @'
table ThreePart

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Sql.Database(ServerName, "Sales", [Query="select r.*, l.x from OtherDb.dbo.Remote r join dbo.Local l on r.id = l.id join srv2.db2.dbo.Linked k on 1 = 1; exec OtherDb.dbo.usp_Run"])
→→→→in
→→→→    Source
'@

Put 'tables\ViaShared.tmdl' @'
table ViaShared

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = #"Shared Conn",
→→→→    t = Source{[Schema="dbo",Item="FromShared"]}[Data],
→→→→    lookup = SomeTable{[Code="X"]}[Label]
→→→→in
→→→→    t
'@

Put 'tables\Excel.tmdl' @'
table Excel

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Excel.Workbook(File.Contents("\\share\finance\Budget 2026.xlsx"), null, true),
→→→→    Sheet = Source{[Item="Budget",Kind="Sheet"]}[Data]
→→→→in
→→→→    Sheet
'@

Put 'tables\SharePoint.tmdl' @'
table SharePoint

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = SharePoint.Files("https://contoso.sharepoint.com/sites/fin", [ApiVersion = 15]),
→→→→    f = Source{[Name="targets.csv",#"Folder Path"="https://contoso.sharepoint.com/sites/fin/Shared Documents/"]}[Content],
→→→→    csv = Csv.Document(f, [Delimiter=";"])
→→→→in
→→→→    csv
'@

Put 'tables\Databricks.tmdl' @'
table Databricks

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Databricks.Catalogs("adb-1.azuredatabricks.net", "/sql/1.0/warehouses/x", [Catalog=null]),
→→→→    cat = Source{[Name="main",Kind="Database"]}[Data],
→→→→    sch = cat{[Name="gold",Kind="Schema"]}[Data],
→→→→    tbl = sch{[Name="fact_tickets",Kind="Table"]}[Data]
→→→→in
→→→→    tbl
'@

Put 'tables\Lakehouse.tmdl' @'
table Lakehouse

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Lakehouse.Contents(null),
→→→→    ws = Source{[workspaceId="w-1"]}[Data],
→→→→    lh = ws{[lakehouseId="l-1"]}[Data],
→→→→    t = lh{[Id="silver_games",ItemKind="Table"]}[Data]
→→→→in
→→→→    t
'@

Put 'tables\Odbc.tmdl' @'
table Odbc

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Odbc.Query("dsn=Hive;uid=svc;pwd=S3cret!;database=raw", "select * from raw.events")
→→→→in
→→→→    Source
'@

Put 'tables\Entered.tmdl' @'
table Entered

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45W", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [A = _t])
→→→→in
→→→→    Source
'@

Put 'tables\Calc.tmdl' @'
table Calc

→measure M1 = ```
→partition Fake = m
→→→```
→→formatString: 0

→partition Calc = calculated
→→mode: import
→→source = CALENDAR(DATE(2020,1,1), DATE(2030,12,31))
'@

Put 'tables\LocalDateTable_11111111-2222-3333-4444-555555555555.tmdl' @'
table LocalDateTable_11111111-2222-3333-4444-555555555555
→partition x = calculated
→→mode: import
→→source = Calendar(Date(2015,1,1), Date(2015,1,2))
'@

Put 'tables\DirectLake.tmdl' @'
table DirectLake

→partition DirectLake = entity
→→mode: directLake
→→source
→→→entityName: fact_sales
→→→schemaName: gold
→→→expressionSource: DatabaseQuery
'@

Put 'tables\Composite.tmdl' @"
table Composite

→partition Composite = entity
→→mode: directQuery
→→source
→→→entityName: Customer
→→→expressionSource: 'DirectQuery to AS - Sales'
"@

Put 'tables\DualMode.tmdl' @"
table DualMode

→partition DualMode = m
→→mode: dual
→→source =
→→→→let
→→→→    Source = Sql.Database("srv5", "dw"),
→→→→    t = Source{[Schema="dbo",Item="DimGeo"]}[Data]
→→→→in
→→→→    t
"@

Put 'tables\Concat.tmdl' @"
table Concat

→partition Concat = m
→→mode: import
→→source =
→→→→let
→→→→    sqlQuery = "
→→→→-- Query Start
→→→→SELECT a.*
→→→→FROM fin.VF_ADJUSTMENTS a
→→→→WHERE FK_DATE >= " & DateTime.ToText(#"RangeStart","yyyyMMdd") & "
→→→→  AND FK_DATE < " & DateTime.ToText(#"RangeEnd","yyyyMMdd") & "
→→→→-- Query End",
→→→→    Source = Sql.Database(ServerName, "Sales", [Query=sqlQuery])
→→→→in
→→→→    Source
"@

Put 'tables\DynTable.tmdl' @"
table DynTable

→partition DynTable = m
→→mode: import
→→source =
→→→→let
→→→→    statement = "SELECT * FROM " & TableParam & " t WHERE 1=1",
→→→→    Source = Sql.Database(ServerName, "Sales", [Query=statement])
→→→→in
→→→→    Source
"@

Put 'tables\Measures.tmdl' @"
table Measures

→partition Measures = m
→→mode: import
→→source = Sql.Database(ServerName, "Sales", [Query="SELECT NULL AS ColumnA WHERE 1=2"])
"@

Put 'tables\Dup.tmdl' @"
table Dup

→partition Dup = m
→→mode: import
→→source =
→→→→let
→→→→    a = #"Shared Lookup",
→→→→    b = Sql.Database(ServerName, "Sales", [Query="select * from [dbo].[X] x join dbo.Y y on 1=1"]),
→→→→    c = Table.NestedJoin(a, {"k"}, b, {"k"}, "n")
→→→→in
→→→→    c
"@

Put 'tables\NestedLet.tmdl' @"
table NestedLet

→partition NestedLet = m
→→mode: import
→→source =
→→→→let
→→→→    POS = let
→→→→        sqlQuery = "
→→→→SELECT p.* FROM dwh.DIM_POS p
→→→→LEFT JOIN dwh.DIM_REGION r ON r.id = p.region_id",
→→→→        Source = Sql.Database(ServerName, "Sales", [Query=sqlQuery])
→→→→    in
→→→→        Source,
→→→→    #"Added" = Table.AddColumn(POS, "x", each 1)
→→→→in
→→→→    #"Added"
"@

Put 'tables\RefOtherTable.tmdl' @"
table RefOtherTable

→partition RefOtherTable = m
→→mode: import
→→source =
→→→→let
→→→→    Source = NavParam,
→→→→    First = Table.FirstN(Source, 1)
→→→→in
→→→→    First
"@

Put 'tables\StepShadow.tmdl' @"
table StepShadow

→partition StepShadow = m
→→mode: import
→→source =
→→→→let
→→→→    Excel = Sql.Database("srv7", "dw"),
→→→→    t = Excel{[Schema="dbo",Item="Z"]}[Data]
→→→→in
→→→→    t
"@

Put 'tables\LastRefresh.tmdl' @"
table LastRefresh

→partition LastRefresh = m
→→mode: import
→→source =
→→→→let
→→→→    Source = #table({"LastRefresh"}, {{DateTime.LocalNow()}})
→→→→in
→→→→    Source
"@

Put 'tables\FieldAccess.tmdl' @"
table FieldAccess

→partition FieldAccess = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Sql.Database("srv8", "dw"),
→→→→    t = Source{[Schema="dbo",Item="F"]}[Data],
→→→→    x = Table.AddColumn(t, "k", each [Excel] & [NavParam Region] & Text.From(t[Databricks]{0})),
→→→→    y = Table.SelectColumns(x, {"Lakehouse", "k"}),
→→→→    z = Table.TransformColumnTypes(y, {{"k", type text}}, [Culture = Odbc])
→→→→in
→→→→    z
"@

Put 'tables\TwoConns.tmdl' @"
table TwoConns

→partition TwoConns = m
→→mode: import
→→source =
→→→→let
→→→→    a = Sql.Database("s1", "d1"),
→→→→    b = Sql.Database("s2", "d2"),
→→→→    x = a{[Schema="dbo",Item="A"]}[Data],
→→→→    y = b{[Schema="dbo",Item="B.Dotted"]}[Data],
→→→→    z = Value.NativeQuery(a, "select * from dbo.C")
→→→→in
→→→→    Table.Combine({x, y, z})
"@

Put 'tables\HiddenSql.tmdl' @"
table HiddenSql

→partition HiddenSql = m
→→mode: import
→→source = Sql.Database("s", "d", [Query="SELECT " & Text.Combine({"* FROM dbo.T"})])
"@

Put 'tables\TrimAndScope.tmdl' @"
table TrimAndScope

→partition TrimAndScope = m
→→mode: import
→→source =
→→→→let
→→→→    q = "SELECT * FROM dbo.Outer",
→→→→    inner = let q = "SELECT * FROM dbo.Inner", x = 0 in x,
→→→→    trimmed = Sql.Database("s", "d", [Query="SELECT * FROM dbo.[" & Text.Trim(" Trimmed ") & "]"]),
→→→→    r = Sql.Database("s", "d", [Query=q])
→→→→in
→→→→    r
"@

Put 'tables\number.tmdl' @"
table number

→partition number = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Sql.Database("s", "d"),
→→→→    t = Source{[Schema="dbo",Item="NumberTable"]}[Data]
→→→→in
→→→→    t
"@

Put 'tables\NotAReference.tmdl' @"
table NotAReference

→partition NotAReference = m
→→mode: import
→→source =
→→→→let
→→→→    f = (number) => number + 1,
→→→→    t = #table(type table [a = number], {{1}}),
→→→→    u = Table.TransformColumnTypes(t, {{"a", type number}}),
→→→→    v = Table.AddColumn(u, "b", each f([a]), type nullable number)
→→→→in
→→→→    v
"@

Put 'tables\CaseSensitive.tmdl' @"
table CaseSensitive

→partition CaseSensitive = m
→→mode: import
→→source =
→→→→let
→→→→    Source = DateTime.LocalNow(),
→→→→    t = #table(1, {{Source}}),
→→→→    excel = 1,
→→→→    u = Table.AddColumn(t, "n", each navparam + excel)
→→→→in
→→→→    u
"@

Put 'tables\AdobeCube.tmdl' @"
table AdobeCube

→partition AdobeCube = m
→→mode: import
→→source =
→→→→let
→→→→    Source = AdobeAnalytics.Cubes([HierarchicalNavigation=true]),
→→→→    #"my company" = Source{[Name="my company"]}[Data],
→→→→    suite = #"my company"{[Id="acme.prod"]}[Data],
→→→→    t = Cube.Transform(suite, {{Cube.AddMeasureColumn, "x", "cm1"}})
→→→→in
→→→→    t
"@

Put 'tables\UsageMetrics.tmdl' @"
table UsageMetrics

→partition UsageMetrics = m
→→mode: import
→→source =
→→→→let
→→→→    Source = UsageMetricsDataConnector.GetMetricsData(BaseUrl & "/metadata/v201906/metrics/workspace/" & WorkspaceId & "/reportviews"),
→→→→    t = Table.FromRecords(Source)
→→→→in
→→→→    t
"@

Put 'tables\PbiRestApi.tmdl' @"
table PbiRestApi

→partition PbiRestApi = m
→→mode: import
→→source =
→→→→let
→→→→    Source = PowerBIRESTAPI.Navigation(),
→→→→    ws = Source{[Key="AppWorkspace"]}[Data],
→→→→    links = ws{[Key="DatasetToDataflowLinks"]}[Data]
→→→→in
→→→→    links
"@
Put 'tables\Incremental.tmdl' @'
table Incremental

→refreshPolicy
→→policyType: basic
→→sourceExpression =
→→→→let
→→→→    Source = Sql.Database("srv3", "dw"),
→→→→    t = Source{[Schema="dbo",Item="FactBig"]}[Data],
→→→→    f = Table.SelectRows(t, each [d] >= RangeStart and [d] < RangeEnd)
→→→→in
→→→→    f

→partition 2024 = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Sql.Database("srv3", "dw"),
→→→→    t = Source{[Schema="dbo",Item="FactBig"]}[Data]
→→→→in
→→→→    t
'@

Put 'tables\Dynamic.tmdl' @'
table Dynamic

→partition p = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Sql.Database("srv4", "dw"),
→→→→    t = Source{[Schema="dbo",Item=UnknownVariable]}[Data],
→→→→    w = Web.Contents("https://api.example.com", [RelativePath="v1/items", Query=[a="1"]]),
→→→→    z = Acme.Feed("https://acme.example/feed")
→→→→in
→→→→    t
'@

# ---- fixtures of the review fixes ----
# [Query] of Analysis Services is DAX / MDX: one row to review, never read as SQL. The same text through a SQL connector:
# never a certain "stored procedure EVALUATE" (T-SQL reads a bare word as a procedure call without EXEC).
Put 'tables\AsDax.tmdl' @'
table AsDax

→partition AsDax = m
→→mode: import
→→source = AnalysisServices.Database("powerbi://api.powerbi.com/v1.0/myorg/Shared WS", "Sales Model", [Query="EVALUATE 'Sales'"])
'@
Put 'tables\DaxThroughSql.tmdl' @'
table DaxThroughSql

→partition DaxThroughSql = m
→→mode: import
→→source = Sql.Database("srvd", "dbd", [Query="EVALUATE 'Sales'"])
'@
# Dataflows: {[Id="Workspaces"]} is a navigation root, not an object.
Put 'tables\Dataflow.tmdl' @'
table Dataflow

→partition Dataflow = m
→→mode: import
→→source =
→→→→let
→→→→    Source = PowerPlatform.Dataflows(null),
→→→→    Workspaces = Source{[Id="Workspaces"]}[Data],
→→→→    ws = Workspaces{[workspaceId="w-9"]}[Data],
→→→→    df = ws{[dataflowId="d-9"]}[Data],
→→→→    e = df{[entity="Customers",version=""]}[Data]
→→→→in
→→→→    e
'@
# The same key reached by a path to review and by a certain one: the merged row stays to review.
Put 'tables\MergeFlag.tmdl' @'
table MergeFlag

→partition p1 = m
→→mode: import
→→source = Sql.Database("srvm", "dbm")

→partition p2 = m
→→mode: import
→→source = Sql.Database("srvm", "dbm", [Query="SELECT 1 AS one"])
'@
# Legacy (provider) query partitions: same doubts as any native SQL; server / database from dataSources.tmdl.
Put 'tables\LegacyQuery.tmdl' @'
table LegacyQuery

→partition LegacyQuery = query
→→mode: import
→→dataSource: Legacy Sql
→→query = SELECT * FROM dbo.LegacyA; EXEC (@s)
'@
Put 'tables\LegacyEmpty.tmdl' @'
table LegacyEmpty

→partition LegacyEmpty = query
→→mode: import
→→dataSource: Unknown Source
→→query = EXEC (@dynamic)
'@
Put 'dataSources.tmdl' @'
dataSource 'Legacy Sql'
→connectionString: Provider=SQLNCLI11;Data Source=legacysrv;Initial Catalog=LegacyDb;User ID=reader;Password=LegacySecret9
→impersonationMode: impersonateServiceAccount
'@
# Functions of the standard library are listed one by one where a namespace also reads data.
Put 'tables\SodaFeed.tmdl' @'
table SodaFeed

→partition SodaFeed = m
→→mode: import
→→source = Json.Document(Soda.Feed("https://data.example.org/resource/abcd.json"))
'@
Put 'tables\ExprEval.tmdl' @'
table ExprEval

→partition ExprEval = m
→→mode: import
→→source = Expression.Evaluate("Sql.Database(""hidden"", ""db"")", #shared)
'@
Put 'tables\CdmFolder.tmdl' @'
table CdmFolder

→partition CdmFolder = m
→→mode: import
→→source = Cdm.Contents(AzureStorage.DataLake("https://acct.dfs.core.windows.net/cdm"))
'@
# ODBC: driver={...} braces are not unresolved M; the password of a connection-string PARAMETER is masked everywhere.
Put 'tables\OdbcDriver.tmdl' @'
table OdbcDriver

→partition OdbcDriver = m
→→mode: import
→→source = Odbc.Query("driver={ODBC Driver 18 for SQL Server};server=srvx;database=dbx", "select * from dbo.OdbcT where d > {d '2024-01-01'}")
'@
Put 'tables\SecretParam.tmdl' @'
table SecretParam

→partition SecretParam = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Odbc.DataSource(ConnStr),
→→→→    t = Source{[Name="raw2",Kind="Database"]}[Data]{[Name="events2",Kind="Table"]}[Data]
→→→→in
→→→→    t
'@
# Schema-only navigation (Oracle / Teradata) and a connector whose first level is the object (Salesforce).
Put 'tables\OracleSchema.tmdl' @'
table OracleSchema

→partition OracleSchema = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Oracle.Database("orahost/svc", [HierarchicalNavigation=true]),
→→→→    HR = Source{[Schema="HR"]}[Data],
→→→→    EMP = HR{[Name="EMPLOYEES"]}[Data]
→→→→in
→→→→    EMP
'@
Put 'tables\SalesforceObj.tmdl' @'
table SalesforceObj

→partition SalesforceObj = m
→→mode: import
→→source =
→→→→let
→→→→    Source = Salesforce.Data(),
→→→→    Account = Source{[Name="Account"]}[Data]
→→→→in
→→→→    Account
'@
# A large query (about 30 KB): the structural scans must stay fast and give the same result.
$bigSteps = foreach ($k in 1..400) { "→→→→    #""Added $k"" = Table.AddColumn(#""Added $($k - 1)"", ""c$k"", each [a] + $k, type number)," }
Put 'tables\BigQuery.tmdl' ("table BigQuery`n`n→partition BigQuery = m`n→→mode: import`n→→source =`n→→→→let`n→→→→    Source = Sql.Database(""srvbig"", ""dbbig""),`n→→→→    #""Added 0"" = Source{[Schema=""dbo"",Item=""BigFact""]}[Data],`n" + ($bigSteps -join "`n") + "`n→→→→    Result = #""Added 400""`n→→→→in`n→→→→    Result`n")

# second workspace / model (for -Workspace and the CSV merge) and a workspace folder NAMED like the OneLake sync root
foreach ($d in (Join-Path $export '[WS] Other\Second.SemanticModel\definition\tables'), (Join-Path $root 'export2\OneLake - Microsoft\Trap.SemanticModel\definition\tables')) {
    New-Item -ItemType Directory -Path $d -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $d 'OtherTable.tmdl'), "table OtherTable`n`n`tpartition p = m`n`t`tmode: import`n`t`tsource =`n`t`t`t`tlet`n`t`t`t`t    Source = Sql.Database(`"srv9`", `"db9`"),`n`t`t`t`t    t = Source{[Schema=`"dbo`",Item=`"Other`"]}[Data]`n`t`t`t`tin`n`t`t`t`t    t`n")
}

[IO.File]::WriteAllText((Join-Path $def 'tables\CON.tmdl'), "table CON`n`n`tpartition p = calculated`n`t`tmode: import`n`t`tsource = ROW(`"a`", 1)`n")

function Assert($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } }
try {
    Push-Location $root
    $rows = @(& $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -PassThru 6>$null)
    $by = { param($t) @($rows | Where-Object Table -EQ $t) }
    # ---- web page: the rows and their query files, gzip + base64 inside one element, nothing else from the data ----
    $html = [IO.File]::ReadAllText((Join-Path $out 'model_sources.html'))
    $embedded = [regex]::Matches($html, '<script id="model-sources-data" type="application/octet-stream" data-encoding="gzip-base64">([^<]*)</script>')
    Assert ($embedded.Count -eq 1 -and $embedded[0].Groups[1].Value -cmatch '\A[A-Za-z0-9+/]+={0,2}\z') 'web page: one data element, base64 only'
    Assert ($html -notmatch 'MODEL_SOURCES_DATA') 'web page: marker replaced'
    function Read-PageData([string]$Html) {
        $m = [regex]::Match($Html, 'data-encoding="gzip-base64">([^<]*)</script>')
        $zipped = [IO.Compression.GZipStream]::new([IO.MemoryStream]::new([Convert]::FromBase64String($m.Groups[1].Value)), [IO.Compression.CompressionMode]::Decompress)
        $reader = [IO.StreamReader]::new($zipped, [Text.Encoding]::UTF8); try { return $reader.ReadToEnd() | ConvertFrom-Json -AsHashtable } finally { $reader.Dispose() }
    }
    $page = Read-PageData $html
    $fields = 'Workspace', 'SemanticModel', 'Table', 'SourceType', 'Server', 'Database', 'Object', 'ObjectType', 'Mode', 'Via', 'Parameters', 'NeedsReview', 'QueryFile', 'SqlFile'
    $pageLines = @($page.rows | ForEach-Object { $r = $_; ($fields | ForEach-Object { $r[$_] }) -join "`t" } | Sort-Object -CaseSensitive)
    $csvLines = @($rows | ForEach-Object { $r = $_; ($fields | ForEach-Object { [string]$r.$_ }) -join "`t" } | Sort-Object -CaseSensitive)
    Assert ($page.generator -ceq 'Parse-SemanticModels.ps1' -and ($pageLines -join "`n") -ceq ($csvLines -join "`n")) "web page: every row and field identical to the inventory ($($pageLines.Count) / $($csvLines.Count))"
    $referenced = @($page.rows | ForEach-Object { $_.QueryFile; $_.SqlFile } | Where-Object { $_ } | Sort-Object -Unique -CaseSensitive)
    Assert ($referenced.Count -and @($referenced | Where-Object { -not $page.files.ContainsKey($_) -or $page.files[$_] -cne [IO.File]::ReadAllText((Join-Path $out $_)) }).Count -eq 0) 'web page: every query file embedded, text identical'
    $outside = $html.Remove($html.IndexOf('data-encoding="gzip-base64">'), $embedded[0].Groups[1].Value.Length)
    Assert ($outside -notmatch 'srv\.database\.windows\.net|NavParam') 'web page: no inventory text outside the encoded element'
    # A junction planted in the output folder must not let the page embed a file from elsewhere (filtered run: the
    # rows of the other models, and the paths of their query files, come back from the CSV).
    $victim = @($page.rows | Where-Object { $_.QueryFile -like '*\power-query\*' } | Select-Object -First 1)[0]
    $other = @($page.rows | Where-Object { $_.SemanticModel -cne $victim.SemanticModel } | Select-Object -First 1)[0]
    $pqDir = Split-Path (Join-Path $out $victim.QueryFile) -Parent
    $evilDir = Join-Path $root 'evil-target'; New-Item -ItemType Directory -Path $evilDir -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $evilDir (Split-Path $victim.QueryFile -Leaf)) -Value 'EVIL-OUTSIDE-CONTENT'
    Rename-Item -LiteralPath $pqDir -NewName 'power-query-aside'
    New-Item -ItemType Junction -Path $pqDir -Target $evilDir | Out-Null
    try {
        & $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -Model $other.SemanticModel.Replace('[', '?').Replace(']', '?') 6>$null | Out-Null
        $linkedHtml = [IO.File]::ReadAllText((Join-Path $out 'model_sources.html'))
        Assert ($linkedHtml -notmatch 'EVIL' -and -not (Read-PageData $linkedHtml).files.ContainsKey($victim.QueryFile) -and ([regex]::Matches((Read-PageData $linkedHtml).files.Values -join "`n", 'EVIL')).Count -eq 0) 'web page: a query file reached through a junction is not embedded'
    } finally {
        [IO.Directory]::Delete($pqDir)   # the junction itself, not its target
        Rename-Item -LiteralPath (Join-Path (Split-Path $pqDir -Parent) 'power-query-aside') -NewName 'power-query'
    }
    Assert ((Get-Content -LiteralPath (Join-Path $evilDir (Split-Path $victim.QueryFile -Leaf)) -Raw) -match 'EVIL-OUTSIDE-CONTENT') 'junction target untouched'
    $rows | Select-Object Table, SourceType, Server, Database, Object, ObjectType, Via, Parameters, NeedsReview | Format-Table -AutoSize | Out-String -Width 320
    Assert ((@($rows[0].PSObject.Properties.Name) -join ',') -eq 'Workspace,SemanticModel,Table,SourceType,Server,Database,Object,ObjectType,Mode,Via,Parameters,NeedsReview,QueryFile,SqlFile') "columns: $(@($rows[0].PSObject.Properties.Name) -join ',')"
    foreach ($word in '/?', '-h', '--help', 'help') { $helpText = & $target $word 6>&1 | Out-String; Assert ($helpText -match '-QueryFilesOnly' -and $helpText -match '-NewOnly' -and $helpText -match 'Examples') "'$word' shows the short help instead of failing on a folder named like it" }
    Assert ((& $target -h 6>&1 | Out-String) -match '-NewOnly') '-h switch shows the help'

    Assert (-not ($rows | Where-Object Table -Like 'LocalDateTable*')) 'auto date table skipped'
    $first = @($rows | Where-Object Table -EQ 'NavParam')[0]
    Assert ($first.SemanticModel -eq 'Fixture: Model?' -and $first.Workspace -eq '[WS] Test') 'model display name from .platform, workspace from folder'
    $r = & $by 'NavParam'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'mart_x.dim_customer' -and $r[0].Parameters -eq 'ParamSchema=mart_x' -and $r[0].SourceType -eq 'Sql.Database (Fabric SQL endpoint)' -and $r[0].Database -eq 'dwh' -and -not $r[0].NeedsReview) "NavParam: $($r | Out-String)"
    $r = & $by "Native's SQL"; Assert ((@($r.Object) -join ',') -eq 'dbo.Orders,sales.Customers,ref.Country') "NativeSql objects: $(@($r.Object) -join ',')"
    Assert ($r[0].Server -eq 'srv.database.windows.net' -and $r[0].SourceType -eq 'Sql.Database (Azure SQL)' -and $r[0].Mode -eq 'directQuery') 'NativeSql server parameter resolved'
    $r = & $by 'NativeQuery2'; Assert ((@($r.Object) -join ',') -eq 'dbo.FactSales,dbo.DimDate,dbo.usp_Load' -and $r[0].Database -eq 'DbFromNav' -and $r[2].ObjectType -eq 'Stored procedure') "NativeQuery2: $($r | Out-String)"
    $r = & $by 'ViaShared'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'dbo.FromShared' -and $r[0].Database -eq 'SharedDb' -and $r[0].Server -eq 'srv.database.windows.net' -and $r[0].Via) "ViaShared: $($r | Out-String)"
    $r = & $by 'Excel'; Assert ($r.Count -eq 1 -and $r[0].Object -eq '\\share\finance\Budget 2026.xlsx | Sheet: Budget' -and $r[0].ObjectType -eq 'File') "Excel: $($r | Out-String)"
    $r = & $by 'SharePoint'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'https://contoso.sharepoint.com/sites/fin/Shared Documents/targets.csv') "SharePoint: $($r | Out-String)"
    $r = & $by 'Databricks'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'gold.fact_tickets' -and $r[0].Database -eq 'main' -and $r[0].ObjectType -eq 'Table') "Databricks: $($r | Out-String)"
    $r = & $by 'Lakehouse'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'silver_games' -and $r[0].Database -eq 'workspaceId=w-1, lakehouseId=l-1') "Lakehouse: $($r | Out-String)"
    $r = & $by 'Odbc'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'raw.events' -and $r[0].Server -eq 'dsn=Hive;database=raw') "Odbc: $($r | Out-String)"
    $r = & $by 'Entered'; Assert ($r.Count -eq 1 -and $r[0].SourceType -like 'Entered data*' -and -not $r[0].NeedsReview) "Entered: $($r | Out-String)"
    $r = & $by 'Calc'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'Calculated table (DAX)') "Calc (fenced measure must not create a fake partition): $($r | Out-String)"
    $r = & $by 'DirectLake'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'gold.fact_sales' -and $r[0].Database -eq 'lakehouse_gold' -and $r[0].Mode -eq 'directLake') "DirectLake: $($r | Out-String)"
    $r = & $by 'Composite'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'Entity of another model / source (AnalysisServices.Database)' -and $r[0].Server -eq 'powerbi://api.powerbi.com/v1.0/myorg/Shared WS' -and $r[0].Database -eq 'Sales Model' -and $r[0].Object -eq 'Customer' -and $r[0].Mode -eq 'directQuery') "Composite (DirectQuery to another model): $($r | Out-String)"
    $r = & $by 'DualMode'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'dbo.DimGeo' -and $r[0].Mode -eq 'dual') "Dual storage mode: $($r | Out-String)"
    $r = & $by 'DirectLake'; Assert ($r[0].SourceType -like 'Direct Lake entity*') "Direct Lake label kept: $($r[0].SourceType)"
    $r = & $by 'Concat'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'fin.VF_ADJUSTMENTS' -and -not $r[0].NeedsReview -and $r[0].Parameters -eq 'ServerName=srv.database.windows.net') "SQL assembled with RangeStart / RangeEnd: $($r | Out-String)"
    $r = & $by 'DynTable'; Assert ($r.Count -eq 1 -and $r[0].Object -eq '{?TableParam}' -and $r[0].NeedsReview -eq 'yes') "table name coming from a variable: $($r | Out-String)"
    $r = & $by 'Measures'; Assert ($r.Count -eq 1 -and -not $r[0].Object -and -not $r[0].NeedsReview) "measure table (valid SQL without any table is a certainty): $($r | Out-String)"
    $r = @(& $by 'ThreePart' | Sort-Object Object)
    Assert ((@($r | ForEach-Object { "$($_.Server)|$($_.Database)|$($_.Object)|$($_.ObjectType)" }) -join ',') -eq 'srv2|db2|dbo.Linked|Table/View,srv.database.windows.net|Sales|dbo.Local|Table/View,srv.database.windows.net|OtherDb|dbo.Remote|Table/View,srv.database.windows.net|OtherDb|dbo.usp_Run|Stored procedure' -and -not @($r | Where-Object NeedsReview).Count) "database / linked-server prefix moved to the Database / Server columns: $($r | Select-Object Server, Database, Object, ObjectType, NeedsReview | Out-String)"
    $r = & $by 'Dup'; Assert ($r.Count -eq 2 -and (@($r.Object) -join ',') -eq 'dbo.X,dbo.Y' -and $r[0].Via -eq 'Shared Lookup') "same object reached directly and through a shared expression -> one row: $($r | Out-String)"
    Assert (-not @($rows | Where-Object { $_.Parameters -match '(?i)\bselect\b|sqlQuery|statement' }).Count) 'a SQL text never lands in the Parameters column'
    $r = & $by 'NestedLet'; Assert ($r.Count -eq 2 -and (@($r.Object) -join ',') -eq 'dwh.DIM_POS,dwh.DIM_REGION' -and -not $r[0].NeedsReview) "SQL held in a step of a nested let: $($r | Out-String)"
    $r = & $by 'RefOtherTable'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'mart_x.dim_customer' -and $r[0].Via -eq 'NavParam') "a table that references ANOTHER table inherits its sources: $($r | Out-String)"
    $r = & $by 'StepShadow'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'dbo.Z' -and $r[0].Server -eq 'srv7') "a let step named like another table ('Excel') is not a reference to it: $($r | Out-String)"
    $r = & $by 'LastRefresh'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'No external source (generated in M)' -and -not $r[0].NeedsReview) "generated in M: $($r | Out-String)"
    $r = & $by 'FieldAccess'; Assert ((@($r.Object | Sort-Object) -join ',') -eq 'dbo.F,raw.events' -and ($r | Where-Object Object -EQ 'raw.events').Via -eq 'Odbc') "column names such as [Excel] are not references to the table 'Excel'; a record VALUE still is: $($r | Select-Object Object, Via | Out-String)"
    $r = & $by 'TwoConns'
    Assert (($r | Where-Object Object -EQ 'dbo.A').Server -eq 's1' -and ($r | Where-Object Object -EQ 'dbo.[B.Dotted]').Server -eq 's2' -and ($r | Where-Object Object -EQ 'dbo.C').Server -eq 's1' -and $r.Count -eq 3 -and -not @($r | Where-Object NeedsReview).Count) "each navigation / native query goes to the connection it targets, and a dotted item keeps its brackets: $($r | Select-Object Server, Database, Object, NeedsReview | Out-String)"
    $r = & $by 'HiddenSql'; Assert ($r.Count -eq 1 -and $r[0].NeedsReview -eq 'yes') "opaque M text may hide a FROM: never a certainty: $($r | Out-String)"
    $r = & $by 'TrimAndScope'
    Assert (@($r | Where-Object Object -EQ 'dbo.Trimmed').Count -eq 1 -and -not @($r | Where-Object Object -EQ 'dbo.[ Trimmed ]').Count) "Text.Trim is applied: $(@($r.Object) -join ',')"
    Assert (-not @($r | Where-Object { $_.Object -eq 'dbo.Inner' -and -not $_.NeedsReview }).Count -and @($r | Where-Object { $_.Object -eq '{dynamic SQL}' -and $_.NeedsReview }).Count -eq 1) "a step name defined twice (nested let) is never resolved to a certain, wrong table: $($r | Select-Object Object, NeedsReview | Out-String)"
    $r = & $by 'NotAReference'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'No external source (generated in M)' -and -not $r[0].Via) "a function parameter or a type name called 'number' is not the table 'number': $($r | Out-String)"
    $r = & $by 'CaseSensitive'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'No external source (generated in M)' -and -not $r[0].Via) "M is case-sensitive: 'navparam' is not the table NavParam: $($r | Out-String)"
    $r = & $by 'AdobeCube'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'acme.prod' -and $r[0].Database -eq 'my company' -and $r[0].SourceType -eq 'AdobeAnalytics.Cubes' -and -not $r[0].NeedsReview) "Adobe report suite by Id: $($r | Out-String)"
    $r = & $by 'UsageMetrics'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'https://wabi.example.net/metadata/v201906/metrics/workspace/1111/reportviews' -and -not $r[0].NeedsReview) "usage metrics URL resolved from parameters: $($r | Out-String)"
    $r = & $by 'PbiRestApi'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'DatasetToDataflowLinks' -and $r[0].Database -eq 'AppWorkspace' -and -not $r[0].NeedsReview) "Power BI REST API navigation by Key: $($r | Out-String)"
    $r = & $by 'Incremental'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'dbo.FactBig') "Incremental de-duplicated: $($r | Out-String)"
    $r = & $by 'Dynamic'
    Assert (($r | Where-Object Object -EQ 'dbo.{?UnknownVariable}').NeedsReview -eq 'yes') 'unresolved identifier flagged'
    Assert (($r | Where-Object Object -EQ 'https://api.example.com/v1/items').ObjectType -eq 'URL') 'Web.Contents RelativePath'
    Assert (($r | Where-Object SourceType -EQ 'Acme.Feed').NeedsReview -eq 'yes') 'unknown connector reported and flagged'
    # ---- review fixes ----
    $r = & $by 'AsDax'; Assert ($r.Count -eq 1 -and $r[0].ObjectType -eq 'DAX/MDX query' -and -not $r[0].Object -and $r[0].NeedsReview -eq 'yes' -and -not $r[0].SqlFile) "[Query] of Analysis Services is DAX: one row to review, no SQL analysis: $($r | Out-String)"
    $r = & $by 'DaxThroughSql'; Assert ($r.Count -ge 1 -and -not @($r | Where-Object { -not $_.NeedsReview }).Count) "DAX text through a SQL connector is never certain (no 'stored procedure EVALUATE' for sure): $($r | Select-Object Object, ObjectType, NeedsReview | Out-String)"
    $r = & $by 'Dataflow'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'Customers' -and $r[0].ObjectType -eq 'Dataflow entity' -and $r[0].Database -eq 'workspaceId=w-9, dataflowId=d-9' -and -not $r[0].NeedsReview) "dataflow: no 'Workspaces' item row: $($r | Select-Object Object, ObjectType, Database, NeedsReview | Out-String)"
    $r = & $by 'MergeFlag'; Assert ($r.Count -eq 1 -and $r[0].NeedsReview -eq 'yes') "a merge never clears a flag: $($r | Select-Object Object, NeedsReview, Via | Out-String)"
    $r = & $by 'LegacyQuery'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'dbo.LegacyA' -and $r[0].NeedsReview -eq 'yes' -and $r[0].Server -eq 'legacysrv' -and $r[0].Database -eq 'LegacyDb' -and $r[0].SourceType -eq 'Legacy data source (Legacy Sql)' -and $r[0].SqlFile) "legacy query partition: doubt kept (EXEC of a variable), server / database from dataSources.tmdl: $($r | Out-String)"
    $r = & $by 'LegacyEmpty'; Assert ($r.Count -eq 1 -and -not $r[0].Object -and $r[0].NeedsReview -eq 'yes') "legacy query partition without any object found: a row to review: $($r | Out-String)"
    $r = & $by 'SodaFeed'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'Soda.Feed' -and $r[0].Object -eq 'https://data.example.org/resource/abcd.json' -and $r[0].ObjectType -eq 'URL' -and -not $r[0].NeedsReview) "Soda.Feed is a connector: $($r | Out-String)"
    $r = & $by 'ExprEval'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'Expression.Evaluate' -and $r[0].NeedsReview -eq 'yes') "Expression.Evaluate is to review: $($r | Out-String)"
    $r = & $by 'CdmFolder'; Assert ($r.Count -eq 1 -and $r[0].SourceType -eq 'AzureStorage.DataLake' -and $r[0].Object -eq 'https://acct.dfs.core.windows.net/cdm' -and -not $r[0].NeedsReview) "Cdm.Contents wraps the folder connector without a bare row of its own: $($r | Out-String)"
    $r = & $by 'OdbcDriver'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'dbo.OdbcT' -and $r[0].Server -eq 'driver={ODBC Driver 18 for SQL Server};server=srvx;database=dbx' -and -not $r[0].NeedsReview) "driver={...} and {d '...'} are not unresolved M: $($r | Out-String)"
    Assert ([IO.File]::ReadAllText((Join-Path $out $r[0].SqlFile)) -notmatch 'WARNING') 'no unresolved-fragment warning for ODBC braces'
    $r = & $by 'SecretParam'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'events2' -and $r[0].Server -eq 'dsn=Hive2;database=raw2' -and $r[0].Parameters -eq 'ConnStr=dsn=Hive2;pwd=***;database=raw2') "password of a parameter masked in the Parameters column: $($r | Out-String)"
    $r = & $by 'OracleSchema'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'HR.EMPLOYEES' -and -not $r[0].Database -and -not $r[0].NeedsReview) "navigation by Schema then Name: $($r | Out-String)"
    $r = & $by 'SalesforceObj'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'Account' -and -not $r[0].Database -and -not $r[0].NeedsReview) "Salesforce object, not a database: $($r | Out-String)"
    $r = & $by 'BigQuery'; Assert ($r.Count -eq 1 -and $r[0].Object -eq 'dbo.BigFact' -and $r[0].Server -eq 'srvbig' -and -not $r[0].NeedsReview) "large query: $($r | Out-String)"
    # No secret in anything written: CSV, workbook, query files, page (connection strings in M, parameters, dataSources.tmdl).
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $everything = @(Get-ChildItem -LiteralPath $out -Recurse -File | Where-Object Extension -NE '.xlsx' | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) + @((Read-PageData $html).files.Values) + @($page.rows | ForEach-Object { $_.Values })
    $zip = [IO.Compression.ZipFile]::OpenRead((Join-Path $out 'model_sources.xlsx')); try { $everything += @(foreach ($entry in $zip.Entries) { $reader = [IO.StreamReader]::new($entry.Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() } }) } finally { $zip.Dispose() }
    Assert (-not @($everything | Where-Object { $_ -match 'S3cret|Hunter2|LegacySecret9' }).Count) 'no password in any output'
    Assert (@($everything | Where-Object { $_ -match 'pwd=\*\*\*' }).Count) 'masked passwords are still recognisable as such'

    $all = Get-Content -LiteralPath (Join-Path $out 'model_sources.csv') -Raw
    Assert ($all -notmatch 'S3cret|Hunter2|uid=' -and $all -notmatch 'pwd=(?!\*\*\*)') 'no credential in the CSV'
    Assert ($all.StartsWith([char]0xFEFF) -or [IO.File]::ReadAllBytes((Join-Path $out 'model_sources.csv'))[0] -eq 0xEF) 'UTF-8 BOM for Excel'
    Assert ((Get-Content -LiteralPath (Join-Path $out 'model_sources.csv') -TotalCount 1) -match '^"Workspace";"SemanticModel";"Table"') 'delimiter honoured'
    $modelOut = Join-Path $out '[WS] Test\Fixture_ Model_'
    $pq = Get-Content -LiteralPath (Join-Path $modelOut 'power-query\NavParam.pq') -Raw
    Assert ($pq -match '(?m)^let\r?$' -and $pq -match '(?m)^    Source = Sql\.Database' -and $pq -match '(?m)^in\r?$') "query de-indented:`n$pq"
    Assert (Test-Path -LiteralPath (Join-Path $modelOut "power-query\Native's SQL.pq")) 'quoted TMDL name unescaped'
    Assert ((Get-Content -LiteralPath (Join-Path $modelOut 'power-query\NativeQuery2.pq') -Raw) -notmatch '```') 'fence markers removed'
    Assert (Test-Path -LiteralPath (Join-Path $modelOut 'dax\Calc.dax')) 'calculated table -> .dax'
    Assert ((Get-Content -LiteralPath (Join-Path $modelOut 'power-query\Incremental.pq') -Raw) -match 'refresh policy source expression[\s\S]*partition ''2024''') 'incremental: policy + partition in one file'
    Assert (Test-Path -LiteralPath (Join-Path $modelOut '_shared\Shared Conn.pq')) 'shared expressions written'
    Assert (-not (Test-Path -LiteralPath (Join-Path $modelOut 'power-query\DirectLake.pq'))) 'no query file for a Direct Lake entity'
    foreach ($name in @("Native's SQL", 'NativeQuery2', 'Concat', 'Dup', 'NestedLet', 'Odbc', 'Measures')) {
        Assert (Test-Path -LiteralPath (Join-Path $modelOut "sql\$name.sql")) "resolved SQL for $name"
    }
    foreach ($name in @('NavParam', 'Calc', 'Entered', 'DirectLake')) {
        Assert (-not (Test-Path -LiteralPath (Join-Path $modelOut "sql\$name.sql"))) "no native SQL file for $name"
    }
    $sql = Get-Content -LiteralPath (Join-Path $modelOut 'sql\Concat.sql') -Raw
    Assert ($sql -match 'FROM fin.VF_ADJUSTMENTS' -and $sql -match '\{(?!\?)[^{}]*RangeStart[^{}]*\}' -and $sql -match '\{(?!\?)[^{}]*RangeEnd[^{}]*\}' -and $sql -match '(?m)^-- WARNING :') 'resolved concatenation with scalar placeholders and warning'
    $sql = Get-Content -LiteralPath (Join-Path $modelOut 'sql\NativeQuery2.sql') -Raw
    Assert ($sql -match 'dbo.FactSales' -and [regex]::Matches($sql, '(?m)^-- Semantic model :').Count -eq 1 -and $sql -match '-- Origin : .*Value.NativeQuery') 'one header per native query, not per source object'
    $sql = Get-Content -LiteralPath (Join-Path $modelOut 'sql\Dup.sql') -Raw
    Assert ([regex]::Matches($sql, [regex]::Escape('[dbo].[X] x join dbo.Y')).Count -eq 1 -and [regex]::Matches($sql, [regex]::Escape('select * from dbo.X')).Count -eq 1) 'own and shared SQL both retained exactly once'
    Assert ($sql -match '-- Origin : shared expression Shared Lookup') 'shared SQL origin'
    $modelSql = Join-Path $modelOut 'Fixture_ Model_.sql'
    Assert (Test-Path -LiteralPath $modelSql) 'model SQL file uses model folder name'
    $sql = Get-Content -LiteralPath $modelSql -Raw
    Assert ($sql.StartsWith('-- Semantic model :')) 'model SQL header'
    $navBlock = [regex]::Match($sql, '(?s)-- Table : NavParam\r?\n.*?(?=-- Table :|\z)').Value
    Assert ($navBlock -match '-- \(no native SQL: navigation / calculated table / generated in M / entity partition …\)') 'navigation block has no native SQL'
    $concatBlock = [regex]::Match($sql, '(?s)-- Table : Concat\r?\n.*?(?=-- Table :|\z)').Value
    Assert ($concatBlock -match '(?m)^--     .*sqlQuery' -and $concatBlock -match '(?m)^FROM fin.VF_ADJUSTMENTS') 'model block has commented M and resolved SQL'
    Assert ($sql -notmatch '/\*') 'model uses line comments for M'
    $r = & $by 'Concat'; Assert ($r[0].SqlFile -and (Test-Path -LiteralPath (Join-Path $out $r[0].SqlFile))) 'SqlFile points to resolved SQL'
    $r = & $by 'NavParam'; Assert (-not $r[0].SqlFile -and $r[0].QueryFile -match '\\power-query\\NavParam\.pq$') 'navigation QueryFile and empty SqlFile'
    Assert ((@($rows[0].PSObject.Properties.Name)[-2..-1] -join ',') -eq 'QueryFile,SqlFile') 'SqlFile follows QueryFile'
    $allowed = @('Fixture_ Model_.sql', 'power-query', 'dax', 'sql', '_shared')
    Assert (-not @(Get-ChildItem -LiteralPath $modelOut | Where-Object Name -NotIn $allowed).Count) 'no stray files in model root'
    foreach ($file in @(Get-ChildItem -LiteralPath $modelOut -File -Recurse)) {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        Assert (-not ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) "no BOM: $($file.Name)"
        Assert ([IO.File]::ReadAllText($file.FullName) -notmatch '(?<!\r)\n|\r(?!\n)') "CRLF: $($file.Name)"
    }
    $csvFile = Join-Path $out 'model_sources.csv'; $xlsxFile = Join-Path $out 'model_sources.xlsx'
    $csvTime = [IO.File]::GetLastWriteTimeUtc($csvFile); $csvBefore = [IO.File]::ReadAllText($csvFile)
    $xlsxTime = [IO.File]::GetLastWriteTimeUtc($xlsxFile); $xlsxBefore = [Convert]::ToBase64String([IO.File]::ReadAllBytes($xlsxFile))
    [IO.File]::WriteAllText($modelSql, 'stale')
    & $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -Workspace '[WS] Test' -Model 'Fixture*' -QueryFilesOnly 6>$null
    Assert ([IO.File]::ReadAllText($modelSql).StartsWith('-- Semantic model :')) '-QueryFilesOnly regenerates the model query file'
    Assert ([IO.File]::GetLastWriteTimeUtc($csvFile) -eq $csvTime -and [IO.File]::ReadAllText($csvFile) -ceq $csvBefore) '-QueryFilesOnly leaves the CSV untouched'
    Assert ([IO.File]::GetLastWriteTimeUtc($xlsxFile) -eq $xlsxTime -and [Convert]::ToBase64String([IO.File]::ReadAllBytes($xlsxFile)) -ceq $xlsxBefore) '-QueryFilesOnly leaves the XLSX untouched'
    # Excel workbook: valid zip, every part is well-formed XML, cells are text and round-trip exactly
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $xlsx = Join-Path $out 'model_sources.xlsx'
    Assert (Test-Path -LiteralPath $xlsx) 'xlsx written'
    $zip = [IO.Compression.ZipFile]::OpenRead($xlsx)
    try {
        $names = @($zip.Entries | ForEach-Object FullName)
        foreach ($need in '[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/worksheets/sheet1.xml') { Assert ($names -contains $need) "xlsx part $need" }
        $xml = @{}
        foreach ($entry in $zip.Entries) { $reader = [IO.StreamReader]::new($entry.Open()); try { $xml[$entry.FullName] = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() } }
    } finally { $zip.Dispose() }
    $sheetRows = @($xml['xl/worksheets/sheet1.xml'].worksheet.sheetData.row)
    $csvNow = @(Import-Csv -LiteralPath (Join-Path $out 'model_sources.csv') -Delimiter ';')
    Assert ($sheetRows.Count -eq $csvNow.Count + 1) "xlsx rows = header + csv rows ($($sheetRows.Count) vs $($csvNow.Count))"
    $cellText = { param($row, $col) $cell = @($row.c | Where-Object { $_.r -match "^$col\d+$" }); if ($cell.Count) { [string]$cell[0].is.t.'#text' } else { '' } }
    Assert ((& $cellText $sheetRows[0] 'A') -eq 'Workspace' -and $sheetRows[0].c[0].s -eq '1') 'bold header row'
    $native = @($sheetRows | Where-Object { (& $cellText $_ 'C') -eq "Native's SQL" })
    Assert ($native.Count -eq 3 -and (& $cellText $native[0] 'G') -eq 'dbo.Orders') "apostrophe survives XML escaping, objects present"
    $review = @($sheetRows | Where-Object { (& $cellText $_ 'G') -eq 'dbo.{?UnknownVariable}' })
    Assert ($review.Count -eq 1 -and $review[0].c[0].s -eq '2') 'row to review is highlighted'
    foreach ($i in 1..3) {
        $worksheet = $xml["xl/worksheets/sheet$i.xml"].worksheet
        Assert (@($worksheet.sheetData.row | ForEach-Object { $_.c } | Where-Object { $_.t -ne 'inlineStr' }).Count -eq 0) "sheet ${i}: every cell is text (nothing can run as a formula)"
        Assert ($worksheet.sheetViews.sheetView.pane.state -eq 'frozen' -and $worksheet.autoFilter.ref -eq $worksheet.dimension.ref) "sheet ${i}: frozen header and full auto-filter"
        Assert ($worksheet.sheetViews.sheetView.tabSelected -eq [string][int]($i -eq 1)) "sheet ${i}: only first tab selected"
    }
    Assert ($xml['xl/worksheets/sheet1.xml'].worksheet.autoFilter.ref -match '^A1:N\d+$') 'auto-filter on the whole table'
    $workbook = $xml['xl/workbook.xml'].workbook
    Assert ((@($workbook.sheets.sheet.name) -join ',') -ceq 'Sources,Models,Schemas') 'exactly three sheets in order'
    $filters = @($workbook.definedNames.definedName)
    Assert ($filters.Count -eq 3 -and (@($filters.localSheetId) -join ',') -eq '0,1,2' -and @($filters | Where-Object name -NE '_xlnm._FilterDatabase').Count -eq 0) 'one local filter name per sheet'
    foreach ($i in 1..3) {
        Assert (@($xml['[Content_Types].xml'].Types.Override | Where-Object PartName -EQ "/xl/worksheets/sheet$i.xml").Count -eq 1) "sheet $i content type"
        $rel = @($xml['xl/_rels/workbook.xml.rels'].Relationships.Relationship | Where-Object Target -EQ "worksheets/sheet$i.xml")
        Assert ($rel.Count -eq 1 -and $rel[0].Id -eq $workbook.sheets.sheet[$i - 1].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')) "sheet $i workbook relationship"
        Assert ($filters[$i - 1].InnerText -eq ("'" + $workbook.sheets.sheet[$i - 1].name + "'!" + '$A$1:$' + ($xml["xl/worksheets/sheet$i.xml"].worksheet.dimension.ref -replace 'A1:([A-Z]+)(\d+)', '$1$$$2'))) "sheet $i filter range"
    }
    $headerColumns = { param($rows) $map = @{}; foreach ($cell in $rows[0].c) { $map[$cell.is.t.InnerText] = $cell.r -replace '\d+$', '' }; $map }
    $sourceColumns = & $headerColumns $sheetRows
    $nav = @($sheetRows | Where-Object { (& $cellText $_ 'C') -eq 'NavParam' })[0]
    $queryRef = $sourceColumns.QueryFile + $nav.r
    $link = @($xml['xl/worksheets/sheet1.xml'].worksheet.hyperlinks.hyperlink | Where-Object ref -EQ $queryRef)
    Assert ($link.Count -eq 1) 'NavParam QueryFile hyperlink exists'
    $linkRel = @($xml['xl/worksheets/_rels/sheet1.xml.rels'].Relationships.Relationship | Where-Object Id -EQ $link[0].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'))[0]
    Assert ($linkRel.Target -ceq '%5bWS%5d%20Test\Fixture_%20Model_\power-query\NavParam.pq' -and $linkRel.TargetMode -eq 'External') "relative target encoded like Excel does (brackets and spaces only, backslashes kept): $($linkRel.Target)"
    $queryCell = @($nav.c | Where-Object r -EQ $queryRef)[0]
    Assert ((& $cellText $nav $sourceColumns.QueryFile) -ceq 'power-query\NavParam.pq' -and $queryCell.s -notin '0','1','2' -and $queryCell.t -eq 'inlineStr') 'short hyperlink display, dedicated style, still text'
    $xfs = @($xml['xl/styles.xml'].styleSheet.cellXfs.xf)
    $font = $xml['xl/styles.xml'].styleSheet.fonts.font[[int]$xfs[[int]$queryCell.s].fontId]
    Assert ($font.color.rgb -eq 'FF0563C1' -and $null -ne $font.u -and $font.sz.val -eq '11' -and $font.name.val -eq 'Calibri') 'blue underlined hyperlink font'
    Assert ($xml['xl/worksheets/sheet1.xml'].worksheet.autoFilter.NextSibling.LocalName -eq 'hyperlinks') 'hyperlinks follow autoFilter in schema order'
    foreach ($calc in @($sheetRows | Where-Object { (& $cellText $_ 'C') -eq 'Calc' })) {
        Assert (@($xml['xl/worksheets/sheet1.xml'].worksheet.hyperlinks.hyperlink | Where-Object ref -EQ ($sourceColumns.SqlFile + $calc.r)).Count -eq 0) 'Calc empty SqlFile has no hyperlink'
    }
    $modelRows = @($xml['xl/worksheets/sheet2.xml'].worksheet.sheetData.row); $modelColumns = & $headerColumns $modelRows
    $modelPairs = @($csvNow | ForEach-Object { ConvertTo-Json -InputObject @($_.Workspace, $_.SemanticModel) -Compress } | Select-Object -Unique)
    Assert ($modelRows.Count -eq $modelPairs.Count + 1) 'one Models row per CSV model'
    $fixture = @($modelRows | Where-Object { (& $cellText $_ $modelColumns.SemanticModel) -eq 'Fixture: Model?' })[0]
    $fixtureCsv = @($csvNow | Where-Object { $_.Workspace -eq (& $cellText $fixture $modelColumns.Workspace) -and $_.SemanticModel -eq 'Fixture: Model?' })
    Assert ((& $cellText $fixture $modelColumns.Tables) -eq [string]@($fixtureCsv.Table | Select-Object -Unique).Count) 'Models distinct table count'
    Assert ((& $cellText $fixture $modelColumns.SourceRows) -eq [string]$fixtureCsv.Count -and (& $cellText $fixture $modelColumns.ToReview) -eq [string]@($fixtureCsv | Where-Object NeedsReview -EQ 'yes').Count) 'Models source and review counts'
    $sources = (& $cellText $fixture $modelColumns.Sources) -split "`n"
    Assert ($sources -contains 'srv.database.windows.net / Sales' -and $sources -contains 'srv2 / db2' -and $sources -contains 'File.Contents' -and $sources -notcontains 'Calculated table (DAX)') 'Models sources include SQL and files, exclude calculated tables'
    Assert (((& $cellText $fixture $modelColumns.Schemas) -split "`n") -contains 'srv.database.windows.net / OtherDb / dbo') 'Models schema lines'
    foreach ($column in 'Sources', 'Schemas') {
        $cell = @($fixture.c | Where-Object r -EQ ($modelColumns[$column] + $fixture.r))[0]
        Assert ($xfs[[int]$cell.s].alignment.wrapText -eq '1' -and $xfs[[int]$cell.s].alignment.vertical -eq 'top') "Models $column wraps at top"
        Assert ($cell.InnerText -notmatch "`r") "Models $column uses LF only"
    }
    Assert (@($modelRows | Where-Object { $_.HasAttribute('ht') }).Count -eq 0) 'wrapped rows have no explicit height'
    $schemaRows = @($xml['xl/worksheets/sheet3.xml'].worksheet.sheetData.row); $schemaColumns = & $headerColumns $schemaRows
    $remote = @($schemaRows | Where-Object { (& $cellText $_ $schemaColumns.SourceType) -eq 'Sql.Database (Azure SQL)' -and (& $cellText $_ $schemaColumns.Server) -eq 'srv.database.windows.net' -and (& $cellText $_ $schemaColumns.Database) -eq 'OtherDb' -and (& $cellText $_ $schemaColumns.Schema) -eq 'dbo' })
    Assert ($remote.Count -eq 1 -and (& $cellText $remote[0] $schemaColumns.Objects) -eq '2' -and (& $cellText $remote[0] $schemaColumns.NeedsReview) -eq '') 'Schemas OtherDb dbo has two confirmed objects'
    $sales = @($schemaRows | Where-Object { (& $cellText $_ $schemaColumns.Database) -eq 'Sales' -and (& $cellText $_ $schemaColumns.Schema) -eq 'dbo' -and [int](& $cellText $_ $schemaColumns.Tables) -ge 2 })
    Assert ($sales.Count -ge 1) 'Schemas Sales dbo spans multiple tables'
    foreach ($row in @($schemaRows | Select-Object -Skip 1)) {
        if ((& $cellText $row $schemaColumns.NeedsReview) -eq 'yes') { Assert ($row.c[0].s -eq '2') 'Schemas review rows highlighted' }
    }
    for ($i = 2; $i -lt $schemaRows.Count; $i++) {
        $comparison = 0
        foreach ($column in 'Workspace', 'SemanticModel', 'Server', 'Database', 'Schema', 'SourceType') {
            $comparison = [StringComparer]::OrdinalIgnoreCase.Compare((& $cellText $schemaRows[$i - 1] $schemaColumns[$column]), (& $cellText $schemaRows[$i] $schemaColumns[$column]))
            if ($comparison) { break }
        }
        Assert ($comparison -le 0) 'Schemas sorted ordinally ignoring case'
    }
    # Exercise writer edge cases directly without re-running model extraction.
    & {
        Set-StrictMode -Version Latest
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($target, [ref]$null, [ref]$null)
        $writerFunction = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Export-SimpleXlsx' }, $false)
        . ([scriptblock]::Create($writerFunction.Extent.Text))
        $edgePath = Join-Path $root 'writer-edge.xlsx'
        Export-SimpleXlsx $edgePath @(@{ Name = 'Empty'; Rows = @(); Columns = @('Only') })
        Assert (-not (Test-Path -LiteralPath $edgePath)) 'empty first sheet does not write a workbook'
        Export-SimpleXlsx $edgePath @(
            @{ Name = 'Links'; Rows = @([pscustomobject]@{ Link = 'a&b/space #/sql/file%.sql'; Text = "12345678`r`n12345678"; Review = 'yes' }); LinkColumns = @('Link'); WrapColumns = @('Text'); HighlightColumn = 'Review' }
            @{ Name = 'Skipped'; Rows = @() }
            @{ Name = "Header's"; Rows = @(); Columns = @('Only') }
            @{ Name = 'Short'; Rows = @([pscustomobject]@{ Link = 'sql/x.sql' }); LinkColumns = @('Link') }
        )
        $edgeXml = @{}; $edgeZip = [IO.Compression.ZipFile]::OpenRead($edgePath)
        try { foreach ($entry in $edgeZip.Entries) { $reader = [IO.StreamReader]::new($entry.Open()); try { $edgeXml[$entry.FullName] = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() } } } finally { $edgeZip.Dispose() }
        Assert ((@($edgeXml['xl/workbook.xml'].workbook.sheets.sheet.name) -join ',') -eq "Links,Header's,Short") 'empty sheet skipped unless Columns supplied'
        Assert (@($edgeXml['xl/worksheets/sheet2.xml'].worksheet.sheetData.row).Count -eq 1 -and -not $edgeXml.ContainsKey('xl/worksheets/_rels/sheet2.xml.rels')) 'header-only sheet has no link relationships'
        Assert ($edgeXml['xl/worksheets/_rels/sheet1.xml.rels'].Relationships.Relationship.Target -ceq 'a&b\space%20%23\sql\file%25.sql') "link target: spaces, hashes and percent signs encoded, ampersand XML-escaped only, backslashes: $($edgeXml['xl/worksheets/_rels/sheet1.xml.rels'].Relationships.Relationship.Target)"
        $edgeSheet = $edgeXml['xl/worksheets/sheet1.xml'].worksheet
        Assert ($edgeSheet.cols.col[1].width -eq '11' -and $edgeSheet.sheetData.row[1].c[1].is.t.InnerText -ceq "12345678`n12345678") 'wrapped width uses longest line and strips CR'
        $edgeXfs = @($edgeXml['xl/styles.xml'].styleSheet.cellXfs.xf)
        foreach ($cell in $edgeSheet.sheetData.row[1].c) { Assert ($edgeXfs[[int]$cell.s].fillId -eq '3') 'review fill preserved on linked and wrapped cells' }
        Assert ($edgeXml['xl/worksheets/sheet3.xml'].worksheet.sheetData.row[1].c.is.t.InnerText -ceq 'sql/x.sql') 'short paths displayed unchanged in a single-column sheet'
        # Excel refuses a sheet with more than 65,530 hyperlinks: 65,000 at most, the other cells keep their text, one warning.
        function Write-UiWarning([string]$Message) { Write-Warning $Message }
        $capPath = Join-Path $root 'writer-cap.xlsx'
        $capRows = foreach ($i in 1..65005) { [pscustomobject]@{ Link = "W\M\sql\T$i.sql" } }
        $capWarnings = @(Export-SimpleXlsx $capPath @(@{ Name = 'Links'; Rows = $capRows; LinkColumns = @('Link') }) 3>&1)
        $capZip = [IO.Compression.ZipFile]::OpenRead($capPath)
        try {
            $capText = @{}; foreach ($entry in $capZip.Entries) { $reader = [IO.StreamReader]::new($entry.Open()); try { $capText[$entry.FullName] = $reader.ReadToEnd() } finally { $reader.Dispose() } }
        } finally { $capZip.Dispose() }
        Assert ([regex]::Matches($capText['xl/worksheets/sheet1.xml'], '<hyperlink ').Count -eq 65000 -and [regex]::Matches($capText['xl/worksheets/_rels/sheet1.xml.rels'], '<Relationship ').Count -eq 65000) 'at most 65,000 hyperlinks per sheet'
        Assert ($capText['xl/worksheets/sheet1.xml'].Contains('>sql\T65005.sql</t>') -and $capWarnings.Count -eq 1 -and "$($capWarnings[0])" -match '\b5 query file paths') "cells beyond the cap keep their text, one warning: $capWarnings"
    }
    # Masking of secrets, function by function: what must be hidden, and SQL that must NOT be touched.
    & {
        Set-StrictMode -Version Latest
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($target, [ref]$null, [ref]$null)
        foreach ($statement in $ast.EndBlock.Statements) { if ($statement.Extent.Text -match '^\$script:Secret') { . ([scriptblock]::Create($statement.Extent.Text)) } }
        foreach ($name in 'Hide-Secret', 'Test-NameMatch') { . ([scriptblock]::Create($ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true).Extent.Text)) }
        $masked = [ordered]@{
            'OleDb.DataSource("Provider=SQLOLEDB;Password=""abc;def"";Data Source=s")'          = 'OleDb.DataSource("Provider=SQLOLEDB;Password=""***"";Data Source=s")'
            'OleDb.DataSource("Provider=SQLOLEDB;Password=""a""""b;c"";Data Source=s")'         = 'OleDb.DataSource("Provider=SQLOLEDB;Password=""***"";Data Source=s")'
            'Provider=SQLOLEDB;Password="abc;def";Data Source=s'                                = 'Provider=SQLOLEDB;Password="***";Data Source=s'
            'dsn=Hive;uid=svc;pwd=S3cret!;database=raw'                                         = 'dsn=Hive;uid=svc;pwd=***;database=raw'
            'Odbc.Query("driver={SQL};pwd={a;b};server=s", "select 1")'                         = 'Odbc.Query("driver={SQL};pwd={***};server=s", "select 1")'
            'Web.Contents(u, [Headers=[Authorization="Bearer eyJhbGciOi.abc.def", #"x-api-key"="k-123456"]])' = 'Web.Contents(u, [Headers=[Authorization="***", #"x-api-key"="***"]])'
            'Web.Contents("https://maps.example.com/api?key=AIzaSyABC&q=1")'                    = 'Web.Contents("https://maps.example.com/api?key=***&q=1")'
            'Content=Text.ToBinary("client_secret=abc123&grant_type=client_credentials")'       = 'Content=Text.ToBinary("client_secret=***&grant_type=client_credentials")'
            'https://x/?code=0123456789abcdef&y=1'                                              = 'https://x/?code=***&y=1'
            'token: eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxIn0.sig'                                   = 'token: ***'
        }
        foreach ($text in $masked.Keys) { $got = Hide-Secret $text; Assert ($got -ceq $masked[$text]) "secret masked: $text -> $got" }
        foreach ($text in @(
                'SELECT u.id FROM dbo.Users u WHERE u.password = @p#(lf)UNION ALL SELECT id FROM dbo.Admins'
                "SELECT password = HASHBYTES('SHA2_256', p) FROM dbo.T"
                "SELECT password=HASHBYTES('x', p) FROM dbo.T"
                'SELECT * FROM dbo.Users WHERE pwd = @p AND token_id = 3'
                '[Query="select password, token from dbo.U where password = ""x"""]'
                'https://x/?code=BE&y=1')) {
            Assert ((Hide-Secret $text) -ceq $text) "SQL / text without a secret left as is: $text -> $(Hide-Secret $text)"
        }
        Assert ((Test-NameMatch 'Sales `EU`' 'Sales `EU`') -and (Test-NameMatch '[WS] A' '[WS]*') -and -not (Test-NameMatch 'Sales EU' 'Sales `EU`')) 'name patterns: backtick and brackets literal'
    }
    # a workbook that is open in Excel (locked) must not break the run: warning, CSV still refreshed
    $lock = [IO.File]::Open($xlsx, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    try { $w = @(& $target -ExportPath .\export -OutputPath .\out -Delimiter ';' 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }) } finally { $lock.Dispose() }
    Assert ($w.Count -eq 1 -and "$($w[0])" -match 'open in Excel') "locked workbook -> warning: $w"
    Assert (-not @(Get-ChildItem -LiteralPath $out -Filter '*.tmp').Count) 'no temporary file left behind'
    # re-run: stale file removed, refusals
    Set-Content -LiteralPath (Join-Path $modelOut 'power-query\Stale.pq') -Value 'x'
    & $target -ExportPath .\export -OutputPath .\out -Delimiter ';' 6>$null
    Assert (-not (Test-Path -LiteralPath (Join-Path $modelOut 'power-query\Stale.pq'))) 'stale query file removed on re-run'
    foreach ($bad in '.\export\sub', '.\OneLake - Microsoft\x') { try { & $target -ExportPath .\export -OutputPath $bad 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'must not be inside|OneLake') $_.Exception.Message } }
    # -Workspace filter (brackets literal, case-insensitive) + CSV merge: the other model's rows are kept, no duplicates
    Assert (@($rows | Where-Object Table -EQ 'OtherTable').Count -eq 1) 'second workspace processed by the full run'
    $full = @(Import-Csv -LiteralPath (Join-Path $out 'model_sources.csv') -Delimiter ';')
    $filtered = @(& $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -Workspace '[ws] test' -PassThru 6>$null)
    Assert ($filtered.Count -and -not ($filtered | Where-Object Workspace -NE '[WS] Test')) "only the selected workspace is processed: $(@($filtered.Workspace | Select-Object -Unique) -join ',')"
    $after = @(Import-Csv -LiteralPath (Join-Path $out 'model_sources.csv') -Delimiter ';')
    Assert ($after.Count -eq $full.Count -and @($after | Where-Object Table -EQ 'OtherTable').Count -eq 1) "filtered run keeps the other model's rows without duplicates: $($full.Count) -> $($after.Count)"
    $one = @(& $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -Workspace '[WS] Other' -Model 'Sec*' -PassThru 6>$null)
    Assert ($one.Count -eq 1 -and $one[0].Object -eq 'dbo.Other') '-Workspace combined with -Model'
    try { & $target -ExportPath .\export -OutputPath .\out -Workspace 'WS*' 6>$null; throw 'expected failure' } catch { Assert ($_.Exception.Message -match 'No exported semantic model matches') $_.Exception.Message }
    # -Exclude leaves a workspace out and keeps its rows; -NewOnly analyses only the models not in the inventory yet
    $excluded = @(& $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -Exclude '[ws] oth*' -PassThru 6>$null)
    Assert ($excluded.Count -and -not ($excluded | Where-Object Workspace -NE '[WS] Test')) "-Exclude leaves the workspace out: $(@($excluded.Workspace | Select-Object -Unique) -join ',')"
    Assert (@(Import-Csv -LiteralPath (Join-Path $out 'model_sources.csv') -Delimiter ';').Count -eq $full.Count) 'rows of the excluded workspace kept'
    $csvBytes = [IO.File]::ReadAllBytes((Join-Path $out 'model_sources.csv'))
    $nothingNew = & $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -NewOnly 6>&1 | Out-String
    Assert ($nothingNew -match 'Nothing new: the 2 selected model') "-NewOnly with nothing new: $nothingNew"
    Assert ([Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $out 'model_sources.csv'))) -ceq [Convert]::ToBase64String($csvBytes)) 'inventory untouched when nothing is new'
    $third = Join-Path $export '[WS] Other\Third.SemanticModel\definition\tables'; New-Item -ItemType Directory -Path $third -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $third 'ThirdTable.tmdl'), "table ThirdTable`n`n`tpartition p = m`n`t`tmode: import`n`t`tsource =`n`t`t`t`tlet`n`t`t`t`t    Source = Sql.Database(`"srv9`", `"db9`"),`n`t`t`t`t    t = Source{[Schema=`"dbo`",Item=`"Third`"]}[Data]`n`t`t`t`tin`n`t`t`t`t    t`n")
    $newRows = @(& $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -NewOnly -PassThru 6>$null)
    Assert ($newRows.Count -eq 1 -and $newRows[0].SemanticModel -eq 'Third' -and $newRows[0].Object -eq 'dbo.Third') "-NewOnly analyses the new model only: $($newRows | Out-String)"
    Assert (@(Import-Csv -LiteralPath (Join-Path $out 'model_sources.csv') -Delimiter ';').Count -eq $full.Count + 1) 'the new model is added to the inventory, the others kept'
    Remove-Item -LiteralPath (Split-Path (Split-Path $third)) -Recurse -Force
    try { & $target -ExportPath .\export -OutputPath .\out -Workspace '[WS] Test' -AllWorkspaces 6>$null; throw 'expected failure' } catch { Assert ($_.Exception.Message -match 'not both') $_.Exception.Message }
    # delimiter mismatch: the existing CSV cannot be merged -> visible warning, CSV rebuilt with the selected models only
    $warn = @(& $target -ExportPath .\export -OutputPath .\out -Delimiter ',' -Workspace '[WS] Other' 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
    Assert ($warn.Count -eq 1 -and "$($warn[0])" -match 'another layout or delimiter') "warning on delimiter mismatch: $warn"
    Assert (@(Import-Csv -LiteralPath (Join-Path $out 'model_sources.csv') -Delimiter ',').Count -eq 1) 'rebuilt with the selected model only'
    # a workspace folder named like the sync root must never become <OutputPath>\OneLake - Microsoft\...
    try { & $target -ExportPath .\export2 -OutputPath .\out2 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'OneLake File Explorer') $_.Exception.Message }
    Assert (-not (Test-Path -LiteralPath (Join-Path $root 'out2\OneLake - Microsoft'))) 'nothing written under a OneLake-named folder'
    # ---- protections added after the Codex audit ----
    Assert (Test-Path -LiteralPath (Join-Path $out '.semantic-model-tools') -PathType Leaf) 'marker written in a folder created by the script'
    Assert (Test-Path -LiteralPath (Join-Path $modelOut 'dax\_CON.dax')) 'table named like a Windows device gets a safe file name'
    # -ExportPath inside -OutputPath: the recursive clean-up could reach the exported definitions
    try { & $target -ExportPath .\export -OutputPath . 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match '-ExportPath must not be inside -OutputPath') $_.Exception.Message }
    Assert (Test-Path -LiteralPath (Join-Path $def 'tables\NavParam.tmdl')) 'exported definitions untouched'
    # junction in the output ancestry: an ordinary-looking path that resolves somewhere else
    $real = Join-Path $root 'real'; New-Item -ItemType Directory -Path (Join-Path $real '[WS] Test\Fixture_ Model_') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $real '[WS] Test\Fixture_ Model_\precious.txt') -Value 'do not delete'
    New-Item -ItemType Junction -Path (Join-Path $root 'jx') -Target $real | Out-Null
    try { & $target -ExportPath .\export -OutputPath .\jx 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'Unsafe directory') $_.Exception.Message }
    Assert (Test-Path -LiteralPath (Join-Path $real '[WS] Test\Fixture_ Model_\precious.txt')) 'nothing deleted through the junction'
    # a non-empty folder that these scripts did not create
    $docs = Join-Path $root 'docs'; New-Item -ItemType Directory -Path (Join-Path $docs '[WS] Test\Fixture_ Model_') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $docs '[WS] Test\Fixture_ Model_\thesis.docx') -Value 'irreplaceable'
    try { & $target -ExportPath .\export -OutputPath .\docs 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'was not created by these scripts') $_.Exception.Message }
    Assert (Test-Path -LiteralPath (Join-Path $docs '[WS] Test\Fixture_ Model_\thesis.docx')) 'same-named foreign folder NOT wiped'
    # only plain drive-letter paths, on both sides
    foreach ($pair in @{ E = ('\\?\' + $export); O = '.\out6' }, @{ E = '.\export'; O = ('\\?\' + (Join-Path $root 'out6')) }, @{ E = '.\export'; O = '\\localhost\c$\srctest-never-created' }) {
        try { & $target -ExportPath $pair.E -OutputPath $pair.O 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'plain drive-letter path') "$($pair.E) / $($pair.O) -> $($_.Exception.Message)" }
    }
    Assert (-not (Test-Path -LiteralPath (Join-Path $root 'out6'))) 'nothing created for a refused path spelling'
    # input-side junction: the exported model reached through a link could be the very folder that gets deleted
    New-Item -ItemType Junction -Path (Join-Path $root 'exlink') -Target $export | Out-Null
    try { & $target -ExportPath .\exlink -OutputPath .\out7 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'Unsafe directory') $_.Exception.Message }
    Assert (Test-Path -LiteralPath (Join-Path $def 'tables\NavParam.tmdl')) 'exported definitions untouched after the refusals'
    # junction BELOW the model folder: its 'definition' sub-folder points into the (marked) output folder
    $src8 = Join-Path $root 'out8\W\M\tables'; New-Item -ItemType Directory -Path $src8 -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $src8 'T.tmdl'), "table T`n")
    [IO.File]::WriteAllText((Join-Path $root 'out8\.semantic-model-tools'), 'marker')
    New-Item -ItemType Directory -Path (Join-Path $root 'export8\W\M.SemanticModel') -Force | Out-Null
    New-Item -ItemType Junction -Path (Join-Path $root 'export8\W\M.SemanticModel\definition') -Target (Join-Path $root 'out8\W\M') | Out-Null
    try { & $target -ExportPath .\export8 -OutputPath .\out8 6>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'Unsafe directory') $_.Exception.Message }
    Assert (Test-Path -LiteralPath (Join-Path $src8 'T.tmdl')) 'source reached through a junction below the model folder is NOT deleted'
    # clearer message when nothing exported matches
    try { & $target -ExportPath .\export -OutputPath .\out -Workspace 'Not exported (dev)' 6>$null; throw 'expected failure' } catch { Assert ($_.Exception.Message -match "exported there so far: '\[WS\] Other', '\[WS\] Test'" -and $_.Exception.Message -match 'Extract-And-Parse-SemanticModels\.ps1') $_.Exception.Message }
    # ---- review fixes that need their own export or output folder ----
    function New-TestModel([string]$Export, [string]$Folder, [hashtable]$Tables, [string]$DisplayName) {
        # $Folder: <workspace>\<model folder name>, without .SemanticModel. Every line of a query is indented, as Fabric does.
        $tablesDir = Join-Path $Export "$Folder.SemanticModel\definition\tables"
        New-Item -ItemType Directory -Path $tablesDir -Force | Out-Null
        if ($DisplayName) { Set-Content -LiteralPath (Join-Path (Split-Path (Split-Path $tablesDir)) '.platform') -Value ('{"metadata":{"type":"SemanticModel","displayName":' + (ConvertTo-Json $DisplayName) + '}}') }
        foreach ($name in $Tables.Keys) {
            $indented = (($Tables[$name] -split "`n") | ForEach-Object { "`t`t`t`t$_" }) -join "`n"
            [IO.File]::WriteAllText((Join-Path $tablesDir "$name.tmdl"), "table $name`n`n`tpartition p = m`n`t`tmode: import`n`t`tsource =`n$indented`n")
        }
    }
    function Read-PageJson([string]$Path) {
        # The raw JSON of the page (ConvertFrom-Json would turn the date into a DateTime).
        $m = [regex]::Match([IO.File]::ReadAllText($Path), 'data-encoding="gzip-base64">([^<]*)</script>')
        $zipped = [IO.Compression.GZipStream]::new([IO.MemoryStream]::new([Convert]::FromBase64String($m.Groups[1].Value)), [IO.Compression.CompressionMode]::Decompress)
        $reader = [IO.StreamReader]::new($zipped, [Text.Encoding]::UTF8); try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    $navigation = { param($Table) "let`n    Source = Sql.Database(""srvc"", ""dbc""),`n    t = Source{[Schema=""dbo"",Item=""$Table""]}[Data]`nin`n    t" }
    # Culture: under Turkish settings (?i:TRIM) does not match "trim", which made "FROM name" a certain, fake table; under
    # Finnish settings the time separator is a dot, which broke the date of the web page. The caller's culture is restored.
    $e4 = Join-Path $root 'export4'
    New-TestModel $e4 'W\Culture' @{ T = "let`n    Source = PostgreSQL.Database(""pg"", ""db""),`n    r = Value.NativeQuery(Source, ""select trim(both 'x' from name) as n, extract(year from d) as y from public.t limit 5"")`nin`n    r" }
    # Each culture runs in a new process: PowerShell caches its regular expressions, so in this process the expressions
    # built under the previous runs' culture would be reused and hide the problem.
    $cultureRun = Join-Path $root 'culture-run.ps1'
    Set-Content -LiteralPath $cultureRun -Value @'
param([string]$Target, [string]$Culture, [string]$Export, [string]$Output)
[Threading.Thread]::CurrentThread.CurrentCulture = $Culture
$rows = @(& $Target -ExportPath $Export -OutputPath $Output -Delimiter ';' -PassThru 6>$null)
"$([Threading.Thread]::CurrentThread.CurrentCulture.Name)|" + (@($rows | ForEach-Object { "$($_.Object)[$($_.NeedsReview)]" }) -join ',')
'@
    $tr = pwsh -NoProfile -File $cultureRun -Target $target -Culture 'tr-TR' -Export (Join-Path $root 'export4') -Output (Join-Path $root 'out4') 2>&1 | Out-String
    Assert ($tr.Trim() -ceq 'tr-TR|public.t[]') "Turkish settings: TRIM(... FROM x) and EXTRACT(... FROM x) read no table, and the caller's culture is restored: $tr"
    $fi = pwsh -NoProfile -File $cultureRun -Target $target -Culture 'fi-FI' -Export (Join-Path $root 'export4') -Output (Join-Path $root 'out4') 2>&1 | Out-String
    Assert ($fi.Trim() -ceq 'fi-FI|public.t[]') "Finnish settings: same result: $fi"
    $json = Read-PageJson (Join-Path $root 'out4\model_sources.html')
    Assert ($json -cmatch '"generatedAtUtc":"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z"') "Finnish settings: ISO date in the page: $([regex]::Match($json, '"generatedAtUtc":"[^"]*"').Value)"
    Assert ([IO.File]::ReadAllText((Join-Path $root 'out4\W\Culture\Culture.sql')) -cmatch '(?m)^-- Generated by Parse-SemanticModels\.ps1 on \d{4}-\d\d-\d\dT\d\d:\d\d:\d\d') 'Finnish settings: ISO date in the model SQL file'
    # Models saved without TMDL (model.bim): counted and named in ONE warning; an export with nothing else stops with a clear message.
    $e5 = Join-Path $root 'export5'
    New-TestModel $e5 'W\Tmdl' @{ T = (& $navigation 'Tmdl') }
    New-Item -ItemType Directory -Path (Join-Path $e5 'W\BimOnly.SemanticModel') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $e5 'W\BimOnly.SemanticModel\model.bim') -Value '{"name":"x","model":{"tables":[]}}'
    $w = @(& $target -ExportPath .\export5 -OutputPath .\out5 -Delimiter ';' 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
    Assert ($w.Count -eq 1 -and "$($w[0])" -match '1 model folder skipped' -and "$($w[0])" -match 'BimOnly' -and "$($w[0])" -match 'model\.bim' -and "$($w[0])" -match 'TMDL') "model.bim: one warning: $w"
    Remove-Item -LiteralPath (Join-Path $e5 'W\Tmdl.SemanticModel') -Recurse -Force
    try { & $target -ExportPath .\export5 -OutputPath .\out5 6>$null 3>$null; throw 'expected failure' } catch { Assert ($_.Exception.Message -match 'TMDL' -and $_.Exception.Message -match '1 model folder') $_.Exception.Message }
    # Folder names are given over all the models of the export: a filtered run gives a model the folder of a full run.
    $e6 = Join-Path $root 'export6'
    New-TestModel $e6 'W\M1' @{ T1 = (& $navigation 'One') } 'Sales:EU'
    New-TestModel $e6 'W\M2' @{ T2 = (& $navigation 'Two') } 'Sales|EU'
    $full6 = @(& $target -ExportPath .\export6 -OutputPath .\out6c -Delimiter ';' -PassThru 6>$null)
    Assert ((@($full6 | Sort-Object Table | ForEach-Object QueryFile) -join ',') -eq 'W\Sales_EU\power-query\T1.pq,W\Sales_EU (2)\power-query\T2.pq') "two models giving the same folder name: $(@($full6.QueryFile) -join ',')"
    $one6 = @(& $target -ExportPath .\export6 -OutputPath .\out6c -Delimiter ';' -Model 'Sales|EU' -PassThru 6>$null)
    Assert ($one6.Count -eq 1 -and $one6[0].QueryFile -eq 'W\Sales_EU (2)\power-query\T2.pq') "filtered run: same folder as the full run: $($one6.QueryFile)"
    Assert ((Test-Path -LiteralPath (Join-Path $root 'out6c\W\Sales_EU\power-query\T1.pq')) -and -not (Test-Path -LiteralPath (Join-Path $root 'out6c\W\Sales_EU\power-query\T2.pq'))) 'the other model keeps its folder and its files'
    # OneDrive: an output folder inside the OneDrive folder is accepted with one warning (its cloud files are reparse points
    # without being links); a junction stays refused there too. NOT simulated here: a real cloud placeholder (it needs the
    # Windows cloud files API and a sync provider), so the acceptance of a reparse point without link target is not exercised.
    $fakeOneDrive = Join-Path $root 'OneDrive - Test'; New-Item -ItemType Directory -Path $fakeOneDrive -Force | Out-Null
    $savedOneDrive = @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)
    try {
        $env:OneDrive = $fakeOneDrive; $env:OneDriveCommercial = $null; $env:OneDriveConsumer = $null
        $w = @(& $target -ExportPath .\export4 -OutputPath (Join-Path $fakeOneDrive 'inventory') -Delimiter ';' 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
        Assert ($w.Count -eq 1 -and "$($w[0])" -match 'OneDrive' -and "$($w[0])" -match 'synced') "output in OneDrive: one warning: $w"
        New-Item -ItemType Directory -Path (Join-Path $root 'od-target') -Force | Out-Null
        New-Item -ItemType Junction -Path (Join-Path $fakeOneDrive 'linked') -Target (Join-Path $root 'od-target') | Out-Null
        try { & $target -ExportPath .\export4 -OutputPath (Join-Path $fakeOneDrive 'linked\inventory') 6>$null 3>$null; throw 'expected refusal' } catch { Assert ($_.Exception.Message -match 'Unsafe directory' -and $_.Exception.Message -match 'Junction') $_.Exception.Message }
        Assert (-not (Test-Path -LiteralPath (Join-Path $root 'od-target\inventory'))) 'nothing written through the junction'
    } finally { $env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer = $savedOneDrive }
    # An output folder inside a git repository (other than data\ of this tool's own repository): one warning.
    New-Item -ItemType Directory -Path (Join-Path $root 'repo\.git') -Force | Out-Null
    $w = @(& $target -ExportPath .\export4 -OutputPath .\repo\inventory -Delimiter ';' 3>&1 6>$null | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
    Assert ($w.Count -eq 1 -and "$($w[0])" -match 'git repository') "output in a git repository: one warning: $w"
    # Default folders are next to the script (not the current folder): only the help text is checked here, since running
    # the script without -ExportPath would read the real export of this machine.
    Assert ((& $target -h 6>&1 | Out-String) -match 'next to this script') 'help: default folders next to the script'
    # The page: route types whitelisted (#constructor, #__proto__ cannot reach a prototype member), no development hook,
    # Ctrl + wheel to zoom; and its script parses (when Node.js is available).
    $source = [IO.File]::ReadAllText($target)
    Assert ($source.Contains('ROUTE_TYPES.has(route.type)') -and -not $source.Contains('window.MODEL_SOURCES') -and $source.Contains('if(!event.ctrlKey&&!event.metaKey)return;')) 'page: route whitelist, no development hook, Ctrl + wheel zoom'
    if (Get-Command node -ErrorAction Ignore) {
        $pageScript = Join-Path $root 'page.js'
        [IO.File]::WriteAllText($pageScript, [regex]::Match($source, "(?s)<script>\n('use strict';.*?)</script>").Groups[1].Value)
        $check = node --check $pageScript 2>&1 | Out-String
        Assert ($LASTEXITCODE -eq 0) "page script parses: $check"
    }
    # ---- console display: plain text with NO_COLOR, a short message and exit code 1 when a person hits an error ----
    $env:NO_COLOR = '1'
    try { $plain = & $target -ExportPath .\export -OutputPath .\out -Delimiter ';' -QueryFilesOnly 6>&1 | Out-String } finally { $env:NO_COLOR = $null }
    Assert ($plain -notmatch "`e" -and $plain -match 'Done in' -and $plain -match 'Query files') "NO_COLOR: summary without escape codes: $plain"
    $failure = pwsh -NoProfile -File $target -ExportPath (Join-Path $root 'missing') -OutputPath .\out 2>&1 | Out-String
    Assert ($LASTEXITCODE -eq 1 -and $failure -match 'Export folder not found' -and $failure -notmatch 'At line|Exception|CategoryInfo') "error typed by a person: short message, exit code 1: $failure"
    'ALL SOURCE-EXTRACTION TESTS PASSED'
}
finally { Pop-Location; Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction Ignore }

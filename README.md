# Fabric Source Atlas

PowerShell scripts that document what your Power BI / Fabric semantic models read: which SQL servers, databases, schemas, tables, views, stored procedures, files and other models feed each table.

**The scripts only read from Fabric. They never write, change, refresh or delete anything there.** They download the *definition* of the models (tables, columns, measures, queries), never their data.

## What you get

`data\sources\model_sources.xlsx`, with three sheets:

| Sheet | One row per | Use it to |
|---|---|---|
| Sources | model, table, source object | see exactly what each table reads; click a link to open the query or the resolved SQL |
| Models | semantic model | see which servers, databases and schemas each model uses |
| Schemas | model, server, database, schema | answer "which models read this schema?", then see what else those models read |

Rows in orange (`NeedsReview = yes`) are the ones the tool could not resolve with certainty: a table name built from a variable, a connector it does not know. Every other row is certain.

`data\sources\model_sources.html` shows the same inventory as a web page. Open it by double-click: it works offline and contacts nothing. Search a schema, a table or a model, see which models are affected and what else they read, follow the lineage from the server down to the model, and read the query of any table without leaving the page. It holds the full code of every query (M and SQL, with server names): treat it like your source code and share it with care.

Next to the workbook: `model_sources.csv` (the Sources sheet as text) and one folder per model with the M query of every table (`power-query\`), the calculated tables (`dax\`), the native SQL with its parameters resolved (`sql\`) and one `<Model>.sql` file that gathers everything.

## Before you start

- **Windows** with **PowerShell 7.2 or later** (`pwsh`, not the blue Windows PowerShell 5.1): https://aka.ms/powershell. Open a PowerShell 7 window in the folder of the scripts.
- **Azure CLI**, logged in: https://aka.ms/installazurecliwindows, then `az login`. The scripts use that login to talk to Fabric; they never see your password.
- **The Contributor role** (or Member / Admin) on the workspaces to document. Fabric requires it to hand out a model definition, even though nothing is written. Viewer is not enough.
- The public Microsoft cloud (`api.fabric.microsoft.com`). Government and China clouds are not supported.
- Optional, for a more precise SQL analysis: Microsoft's T-SQL parser, which comes with SQL Server Management Studio or with the `SqlServer` module (`Install-Module SqlServer -Scope CurrentUser`). Without it, a built-in reader is used.

Downloaded the scripts as a ZIP? Windows marks them as coming from the internet. Unblock them once:

```powershell
Get-ChildItem -Recurse -Filter *.ps1 | Unblock-File
```

## The three commands

```powershell
.\Extract-SemanticModels.ps1 -ListWorkspaces                          # which workspaces can I see?
.\Extract-SemanticModels.ps1 -Workspace '[FIN] Finance' -ListModels   # which models are in that workspace?

.\Extract-And-Parse-SemanticModels.ps1 -Workspace '[FIN] Finance'     # everything, in one go
```

The first lines tell you who you are signed in as and in which tenant. A status line then shows where the run is (workspace, model n of N, waits when Fabric asks to slow down), and a summary at the end says what was exported, skipped or failed and what to do next.

`Extract-And-Parse-SemanticModels.ps1` simply runs the two steps below one after the other. You can also run them separately:

```powershell
.\Extract-SemanticModels.ps1 -Workspace '[FIN] Finance'   # 1. download the definitions to data\export   (reads Fabric)
.\Parse-SemanticModels.ps1                                # 2. build the workbook and the query files    (offline)
```

Step 2 is offline: run it as often as you like, it never contacts Fabric.

## Options

They work with every command and can be combined:

| Option | What it does |
|---|---|
| `-Workspace 'Finance'` | these workspaces only |
| `-AllWorkspaces` | every workspace you can export (not My workspace, not those where you are only Viewer) |
| `-Exclude 'Sandbox*'` | leave these workspaces out |
| `-Model 'Sales*'` | these models only |
| `-NewOnly` | only the models not exported / not in the workbook yet |
| `-TenantId '<guid>'` | sign in to that tenant, when your account belongs to several (Fabric commands only) |
| `-ExportPath` / `-OutputPath` | other folders than `data\export` / `data\sources` next to the scripts |

`*` and `?` are wildcards, square brackets are taken literally, case does not matter. Examples:

```powershell
.\Extract-And-Parse-SemanticModels.ps1 -AllWorkspaces                          # everything you can export
.\Extract-And-Parse-SemanticModels.ps1 -AllWorkspaces -NewOnly                 # later: add the new models only
.\Extract-And-Parse-SemanticModels.ps1 -AllWorkspaces -Exclude 'Sandbox*', 'Archive'
.\Extract-And-Parse-SemanticModels.ps1 'Finance' -Model 'Sales*'
```

## Getting help on a command

Run a script with `/?` (or `-h`, `--help`, `help`, or no argument for the Fabric scripts) to see what it does and one example per option, each with what it changes:

```powershell
.\Extract-And-Parse-SemanticModels.ps1 /?
```

## Questions you may have

**Is it safe to run against production?**
Yes. The export script can only send four kinds of requests: list the workspaces, list the models of a workspace, ask for a model definition, and poll that request until Fabric answers. An allow-list inside the script refuses anything else before it is sent. This was audited several times, by independent reviewers and tools, and an offline test suite replays the Fabric API locally to check it on every change.

**Does it download the data of the models?**
No. Only the definition: tables, columns, measures, relationships, queries. No rows, no reports, no `.pbix`.

**Does it create anything in Fabric?**
No item, no model, no file, no report. Fabric keeps a technical trace of each definition request, as for any call to its API.

**It cannot sign in.**
The message says what Azure CLI answered and the command to run. The usual fixes:
- not signed in, or signed in long ago: `az login`;
- "No subscriptions found": `az login --allow-no-subscriptions` (Fabric needs no Azure subscription);
- multi-factor or conditional access: `az login --scope https://api.fabric.microsoft.com/.default`;
- your account belongs to several tenants, or you are a guest: add `--tenant <tenant id>` to `az login`, and `-TenantId <tenant id>` to the script.

Some organisations block Azure CLI with conditional access; then ask your administrator.

**It refuses my output folder.**
The scripts write only into a plain local folder: not a link or junction, not a OneLake File Explorer folder (whatever is written there syncs back to Fabric). OneDrive folders are accepted, with a warning: your exports will then be synced to OneDrive. Choose another folder with `-ExportPath` / `-OutputPath`, for example `C:\SemanticModels\export`.

**It warns that the output is inside a git repository.**
The exports and the inventory hold your company's code and server names. Keep them out of any repository you push (`data\` of this tool is ignored by git).

**The workbook was not updated.**
It is open in Excel. Close it and run `.\Parse-SemanticModels.ps1` again (offline, no need to export again). The CSV is always up to date.

**A workspace is missing with `-AllWorkspaces`.**
`-AllWorkspaces` means every workspace where you have a role, not the whole tenant (even for a Fabric administrator). My workspace is always left out, and so are the workspaces where you are only Viewer: Fabric does not hand out a model definition below Contributor. The summary counts them; run with `-Verbose` to see their names.

**A model failed.**
Fabric refused to hand out its definition: protected sensitivity label, push or streaming dataset, default model of an old lakehouse or warehouse, live connection. The other models are exported normally; the reason is in `_export_report.csv` next to the exports. If a whole workspace could not be listed, the run goes on with the others and says so.

**Some models are "skipped (no TMDL)".**
The analysis reads the TMDL format that Fabric exports. A folder saved by Power BI Desktop as `model.bim` (older project format) is skipped; save the project with the TMDL format enabled, or export the model with these scripts.

**Fabric is slow, or says "too many requests".**
The script waits as long as Fabric asks (up to 10 minutes at a time, 20 minutes in total for one request) and shows the countdown, retries network errors, and renews the sign-in token during long runs. Leave it running. If a model still fails, run the same command again with `-NewOnly`: it continues where it stopped.

**Are passwords written to the files?**
No. Passwords and keys are replaced by `***` in every file the analysis writes: connection-string passwords and keys (`pwd=`, `password=`, `AccountKey=`, ...), Bearer tokens, secret-named headers and fields (`Authorization`, `x-api-key`, ...) and URL parameters such as `sig=`, `key=` or `token=`. The SQL itself is left untouched. Still, check the files before sharing them.

**What does a row in orange mean?**
The tool is not sure: the table name comes from a variable, the SQL text could not be rebuilt, or the connector is unknown. Open the linked query file to check by hand.

**A database is written in front of an object in the SQL. Where is it?**
In the `Database` column. `Object` always reads `schema.object`; a database (or linked server) named in front of it moves to the `Database` (or `Server`) column.

## How reliable is the inventory

The source objects are found by reading the M code of every table: navigation steps, native SQL, files, URLs, dataflows and Direct Lake entities, with parameters, shared expressions and references to other tables resolved. Native SQL is read with Microsoft's T-SQL parser when it is installed, otherwise with a built-in reader; the two are tested against each other on a battery of tricky queries. A DAX or MDX query sent to another model is listed and flagged, never read as SQL. An unknown connector is listed and flagged, never ignored. The inventory was also cross-checked once, on a real workspace of 55 models and 459 tables, against two independent readings of every query by AI assistants: 466 of the 468 table readings agreed, and the two others described the same file at different levels of detail.

Limits: the inventory is a snapshot, run the scripts again to refresh it. A table name computed at run time is flagged, not guessed. The tool only sees what is written in the model.

## Where things are

```
Extract-SemanticModels.ps1             step 1: Fabric to data\export         (read-only)
Parse-SemanticModels.ps1               step 2: data\export to data\sources   (offline)
Extract-And-Parse-SemanticModels.ps1   step 1 then step 2
data\                                  your exports and results, ignored by git
tests\                                 offline tests, no connection needed
```

## License

MIT, see [LICENSE](LICENSE). Provided as is, without warranty: check the results before relying on them.

## Microsoft documentation

- Get semantic model definition API: https://learn.microsoft.com/en-us/rest/api/fabric/semanticmodel/items/get-semantic-model-definition
- What a semantic model definition contains: https://learn.microsoft.com/en-us/rest/api/fabric/articles/item-management/definitions/semantic-model-definition
- Long-running operations: https://learn.microsoft.com/en-us/rest/api/fabric/articles/long-running-operation

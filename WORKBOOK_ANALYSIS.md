# HIC UPDATE.xlsm: workbook and VBA analysis

## Scope and method

This review examined the `.xlsm` package, workbook XML, worksheet content, drawing/button assignments, embedded `vbaProject.bin`, VBA module names, and extracted source code. The original workbook was not opened for editing and was not modified.

## Business context and reported impact

The workbook supports a recurring Monday process used to prepare and distribute HIC reports to stakeholders and Lines of Business (LOBs). The project owner reported that the manual process took approximately 10 minutes per cycle and that the automated process now takes approximately 3–4 minutes. This indicates an estimated saving of 6–7 minutes per cycle, or roughly 60–70% less elapsed time. The timing is user-reported and was not independently benchmarked during this static workbook review.

## Workbook structure

The supplied file is a macro-enabled Excel workbook with two worksheets:

| Worksheet | Observed purpose |
| --- | --- |
| `Sheet1` | Empty worksheet with a legacy `RUN HIC AUTOMATION` shape. The shape is assigned to `RunHicAll`. |
| `HIC Controller` | Main controller panel. It contains the configuration fields, status fields, and three worksheet shapes used as buttons. |

The controller sheet contains a merged title area, editable network-root and year-folder cells, run date, status, selected-workbook, controller-log fields, and three actions:

1. `1. Validate Setup`
2. `2. Update Without Email`
3. `3. Update and Save Draft`

No Excel tables, PivotTables, QueryTables, workbook connections, charts, or local report data were found in the supplied controller workbook package. The report data and report calculations belong to a separate monthly workbook selected at runtime.

No UserForm streams or MSForms controls were found in the embedded VBA project. The UI is worksheet-based and uses Excel shape objects.

## Controller logic: Module4

The largest controller module is embedded under the VBA stream name `Module4` and has the source attribute `VB_Name = "Module4"`. Its comments describe the intended logical role as `HicController`.

### Setup

`SetupHicController` creates or refreshes the `HIC Controller` sheet, writes the setup fields, formats the panel, creates three rounded-rectangle shapes, and assigns them to validation, test, and draft procedures. It also writes explanatory notes that updates run in hidden Excel and that Outlook drafts are not sent automatically.

### Date and file selection

`IndiaDate` uses the Windows UTC system time and adds 330 minutes. `FindReport` searches the configured root or year folder for full-month and abbreviated report filenames, including space and hyphen variants. `EnsureMonthlyReport` returns an existing current-month report or copies the previous-month report when a current report is missing.

The copy process uses a temporary name, compares source and copy sizes, and moves the completed temporary file into place without overwriting a report that appeared during the copy.

### Validation mode

`RunHicValidate` calls `RunController "Validate"`. Validation checks the panel, network root, year-folder value, report accessibility, and current/previous report availability. It deliberately avoids creating a missing report, opening an Excel worker, or using Outlook.

### Hidden Excel update

`RunHicTest` calls `RunController "Test"`. The controller records existing `EXCEL.EXE` process IDs, creates a new Excel instance, verifies that its process ID is not pre-existing, hides it, and refuses to close any unrelated Excel process.

The report worker is expected to contain:

- A `Query` worksheet.
- A `Pivot` worksheet.
- Unmerged date inputs in `Query!I4:J4`.
- A shape named `Button 1` assigned to the report's existing `ThisWorkbook.MainSub` action.
- A `Query` table containing a `Sample ID` column.

The controller writes the date range, invokes the existing report action, waits for calculation/refresh work, runs `Workbook.RefreshAll`, confirms a non-empty Query table and Sample IDs, saves the report, and restores the worker Excel settings.

### Outlook draft

`RunHicDraft` calls `RunController "Draft"`, then `SaveOutlookDraft`. The draft routine uses late-bound classic Outlook COM automation, writes an HTML body with a report-folder link, adds configured recipients and sender, attaches the saved report, saves the message, and writes a durable `DRAFTED` marker.

The controller does not call `.Send`. The message remains in Outlook Drafts for human review.

### Operational controls

The controller writes logs under a local Windows application-data folder and uses a shared `launcher.lock` file under `.hic-state` for non-validation runs. It also records the selected report, start/end dates, worker PID, progress stages, save status, and draft status.

## Report-side automation modules

Three standard modules are present:

- `HicAutomation`
- `HicAutomation1`
- `HicAutomation2`

Their extracted source is identical apart from the module-name attribute. Each contains the same configuration constants, `HicRun`, refresh-waiting logic, error scanning, and `HicInjectSQL` procedure.

The report-side path is stricter than the controller path. It expects a 40-column `Query` table, 31 SQL source columns, nine calculated columns, and a PivotTable named `PivotTable1`. It uses ADODB and the Microsoft OLE DB SQL Server provider, fully materializes the recordset, resizes the table, reapplies formula templates, refreshes the PivotTable, checks counts and error values, and saves the report.

It also contains `Test`, `Review`, and `Scheduled` modes. The scheduled branch can call Outlook `.Send` when the `AUTO_SEND` constant and mode permit it. This is different from the supplied controller's visible workflow, whose draft routine saves but does not send.

## Important discrepancies and risks

### Button/module-name mismatch

The controller drawing assignments reference `HicController.RunHicValidate`, `HicController.RunHicTest`, and `HicController.RunHicDraft`. The extracted code module is named `Module4`, not `HicController`. Unless the VBA project is renamed elsewhere in the trusted Excel environment, these assignments may fail when clicked. This should be verified in a controlled copy before deployment.

### Legacy shape

`Sheet1` is otherwise blank but contains a `RUN HIC AUTOMATION` shape assigned to `RunHicAll`. It appears to be a leftover or alternate entry point and should be either documented and tested or removed in a reviewed release branch.

### Duplicate modules

The three `HicAutomation` modules duplicate public procedures and configuration. This increases ambiguity and makes future maintenance risky. The intended report-side module should be identified before the project is published.

### Static-analysis warning

A static VBA inspection reported a possible VBA-stomping condition and generic suspicious-operation indicators such as environment-variable reads, file writes, `CreateObject`, and process access. Those operations are consistent with the documented logging, file-copy, Outlook, ADODB, and worker-process behavior. The warning is not proof of malicious code, but the project should be compiled and reviewed in a trusted Excel/VBE environment before release.

## Data and confidentiality findings

The source contains enterprise-specific configuration that should not be published unchanged:

- Internal UNC network path.
- SQL Server instance and database identifiers.
- Real email recipients, CC recipients, and sender address.
- Company email signature, legal name, registered address, privacy-policy links, and social links.
- Local user/log path and saved run-state values.
- Workbook names that reveal the internal process.
- Office document creator, last editor, company, and revision metadata.
- Potentially sensitive report data if a live monthly workbook is added later.

No plaintext password was found. The SQL connection string uses Windows Integrated Security and the source comments explicitly prohibit storing a password in the workbook or an environment variable. That does not make the remaining identifiers public-safe; they should still be replaced with placeholders or moved to controlled deployment configuration.

## Implemented versus suggested

### Implemented in the supplied code

- Controller worksheet and shape-based buttons.
- Network report discovery and optional year-folder handling.
- Previous-month report copying.
- Hidden worker Excel process.
- Report date update and existing report-macro invocation.
- Refresh wait and `RefreshAll`.
- Basic non-empty Query/Sample ID validation.
- Save confirmation.
- Shared locking and operational logging.
- Outlook draft creation without automatic send in the controller path.
- Report-side ADODB/PivotTable automation in the embedded `HicAutomation` source.

### Not evidenced in the supplied controller

- UserForms.
- Self-contained report data or SQL query text.
- A workbook-local scheduler.
- Public-ready configuration management.
- Automated test suite.
- Guaranteed business-level data-quality validation.
- New Outlook support.

### Suggested future work

- Resolve and test all macro bindings.
- Remove duplicate and legacy modules.
- Redact environment configuration and metadata.
- Add repeatable month-boundary and failure-path tests.
- Add structured diagnostics and reconciliation guidance.
- Digitally sign approved production builds.
- Keep exported VBA source in Git and distribute the binary workbook privately.

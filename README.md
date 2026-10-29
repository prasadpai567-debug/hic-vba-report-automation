# HIC Controller: Excel VBA Report Automation

An Excel VBA controller for a recurring HIC report update process. The controller locates the correct monthly report on a Windows network share, prepares the reporting period, runs the report workbook's existing data-retrieval macro in a separate hidden Excel instance, refreshes the report, saves the updated workbook, and creates an Outlook draft for review.

This repository description is based on the supplied `HIC UPDATE.xlsm` workbook and its embedded VBA project. Internal paths, database identifiers, recipients, company addresses, and report data should be replaced with placeholders before publishing.

## Project overview

The application supports a recurring Monday reporting process in which updated HIC reports are distributed to stakeholders and Lines of Business (LOBs). It addresses a repetitive operational reporting task that otherwise requires a user to:

1. Find the correct monthly HIC report in a shared folder.
2. Create the next monthly report from the previous report when required.
3. Enter the reporting-period dates.
4. Run the report workbook's existing data-retrieval macro.
5. Refresh calculations and PivotTable output.
6. Save the updated report.
7. Prepare a standard email with the updated report attached.

### Reported process improvement

Before automation, the Monday report preparation and distribution process took approximately 10 minutes. With the macro, the same workflow can be completed in approximately 3–4 minutes, saving about 6–7 minutes per reporting cycle. This is a project-owner-reported operational measure and should be revalidated with a repeatable timing study if the repository is used for formal performance claims.

The supplied workbook acts as an orchestration layer. It does not contain the operational report data itself. The report workbook is an external monthly `.xlsm` file with expected `Query` and `Pivot` sheets, a `Query` table, a PivotTable, and an existing button macro that performs the report's data retrieval.

## Implemented features

- Worksheet-based controller panel with setup, status, selected-workbook, and log fields.
- Configurable network report root and optional year subfolder.
- India Standard Time run-date calculation using the Windows system clock plus a five-and-a-half-hour offset.
- Monthly report discovery using full-month and abbreviated filename conventions.
- Previous-month report copy when the current-month report is missing.
- Validation-only mode that checks accessibility without creating or updating a report.
- Separate hidden Excel worker instance for report processing.
- Protection against closing or taking ownership of pre-existing Excel processes.
- File locking in a shared `.hic-state` folder to reduce concurrent-run conflicts.
- Reporting-period update in the external report workbook.
- Invocation of the report workbook's existing retrieval macro through its assigned shape action.
- Waiting for calculation, QueryTable, and connection activity to finish.
- `RefreshAll` execution on the report workbook.
- Basic result validation, including a non-empty Query table and Sample ID count.
- Save confirmation and file-size checks.
- Local controller logging and shared operational markers.
- Classic Outlook draft creation with configured To, CC, sender, subject, HTML body, folder link, and report attachment.
- Draft status written to an intent marker so an operator can reconcile an ambiguous email step.
- No automatic send in the controller's visible workflow.

## Current workflow

### 1. Configure the controller

The `HIC Controller` sheet stores the main settings:

| Field | Purpose |
| --- | --- |
| Report network root | UNC path containing the monthly report files. |
| Use year subfolder? | Controls whether the report is searched below a year folder. |
| Run date | Calculated in India time when a run starts. |
| Status | Current progress or error message. |
| Selected workbook | Report file selected or created for the run. |
| Controller log | Local path to the run log. |

Yellow cells are intended for user configuration. Save the controller after setup.

### 2. Validate setup

`RunHicValidate` checks that the configured network folder is available and that a current-month report or previous-month template can be found. If the current report is missing, validation records the planned copy but does not create the report, update Excel, or open Outlook.

### 3. Update without email

`RunHicTest` performs the report update without creating an email:

1. Acquire the shared launcher lock.
2. Locate the current report or copy the previous-month report to the expected current-month filename.
3. Open a new Excel instance and verify that it is a newly created worker process.
4. Open the report read/write as an `.xlsm` file.
5. Write the calculated start and end dates to `Query!I4:J4`.
6. Invoke the existing report workbook action assigned to `Button 1`.
7. Wait for calculations and asynchronous refresh work to finish.
8. Run `Workbook.RefreshAll`.
9. Confirm that the `Query` table is non-empty and contains Sample IDs.
10. Save the report and confirm that Excel reports the workbook as saved.
11. Close the owned worker instance and release the lock.

The date logic in the code uses the 15th day of the previous month as the start date and the last day of the run month as the end date. For example, a run in October uses a period from September 15 through October 31.

### 4. Update and save an Outlook draft

`RunHicDraft` performs the same report update and then:

1. Create a classic Outlook message.
2. Add the configured recipients and sender.
3. Build an HTML body containing a link to the report folder.
4. Attach the saved report workbook.
5. Save the message to Outlook Drafts.
6. Write a durable `DRAFTED` marker.

The controller does not send this message automatically. The operator reviews and sends it manually.

## Technical architecture

```text
HIC UPDATE.xlsm
    |
    |  Finds or creates the monthly report
    |  Opens a separate hidden Excel worker
    v
Monthly HIC report.xlsm
    |
    |  Query!I4:J4 date inputs
    |  Existing Button 1 / report macro
    |  Query table and Pivot sheet
    v
Updated and saved report
    |
    v
Outlook Drafts + run log + .hic-state markers
```

### Languages and technologies

- VBA for Microsoft Excel.
- Excel workbook object model, shapes, tables, PivotTables, calculations, and `RefreshAll`.
- Classic Outlook COM automation for draft creation.
- Windows FileSystemObject for folders, file discovery, copying, locks, and markers.
- Windows Management Instrumentation to enumerate existing Excel processes.
- `kernel32` and `user32` Windows API calls for India-time calculation and Excel process identification.
- UNC/SMB network file access.
- ADODB and the Microsoft OLE DB SQL Server provider in the embedded report-automation modules.
- Windows Integrated Security (`SSPI`) for the report-side database connection. No database password is stored in the supplied code.

### VBA components found

| Component | Observed role |
| --- | --- |
| `Module4` | Controller implementation: setup, validation, worker Excel orchestration, report update, Outlook draft, logging, and monthly file creation. The code comments refer to this logical role as `HicController`. |
| `HicAutomation` | Report-side automation module containing `HicRun`, refresh checks, PivotTable validation, and `HicInjectSQL`. |
| `HicAutomation1` | Duplicate of `HicAutomation` with a different module name. |
| `HicAutomation2` | Duplicate of `HicAutomation` with a different module name. |
| `Module1`, `Module2`, `Module3` | Empty standard modules. |
| `ThisWorkbook`, `Sheet1`, `Sheet2` | Class modules with no user procedure body in the extracted source. |

The three `HicAutomation` modules are effectively identical apart from their module-name attribute. They appear to be report-side modules or retained copies rather than separate features of the controller. Before a public release, remove duplicate source copies in a separate reviewed branch and confirm which module belongs in the monthly report workbook.

### Report-side SQL and validation logic

The embedded `HicAutomation` source includes an alternative report-side path that:

- Opens an ADODB connection using a configured SQL Server provider.
- Executes a supplied SQL statement.
- Requires 31 source columns.
- Copies results into a 40-column `Query` table.
- Preserves and reapplies nine calculated columns.
- Rejects empty or invalid result sets.
- Refreshes `PivotTable1` from the dynamic table.
- Checks row counts, Sample IDs, Pivot cache record count, and Excel error values.
- Saves the report after successful validation.
- Supports `Test`, `Review`, and `Scheduled` modes in that module.

Those report-side modes are present in the embedded source, but the supplied controller's visible buttons call the controller's validation, test, and draft workflow. In particular, the visible controller saves an Outlook draft and does not automatically send mail.

## Requirements

- Windows with desktop Microsoft Excel that can run VBA macros.
- A trusted, macro-enabled `.xlsm` copy of the controller.
- Access to the configured UNC network folder.
- Read/write permission for the report folder and its `.hic-state` subfolder.
- Access to the monthly report workbook and its expected workbook structure.
- Classic Outlook installed and configured if the draft feature is used.
- Permission to create the configured sender and recipients in Outlook.
- For the report-side SQL path: the approved Microsoft OLE DB SQL Server provider, network/database access, and least-privilege Windows authentication.

The controller is Windows-specific because it uses Windows APIs, WMI, UNC paths, Excel COM automation, and classic Outlook automation.

## Installation and setup

1. Obtain the sanitized controller workbook from a trusted release or private repository. Do not download and enable macros from an untrusted public source.
2. Open the workbook in desktop Excel and enable content only after verifying the source.
3. If the `HIC Controller` sheet is absent, run `SetupHicController` from `Alt+F8`. The procedure creates the controller sheet, fields, and three buttons.
4. Set the report network root and year-folder option in the yellow cells.
5. Save the workbook as a macro-enabled `.xlsm` file.
6. Run `Validate Setup` first. Confirm the status, selected workbook, and log path.
7. Run `Update Without Email` as a controlled test.
8. Review the saved report and status. Only then use `Update and Save Draft`.
9. Review the generated draft in Outlook before sending it manually.

Before deployment, verify the button `OnAction` assignments and the module name. The supplied workbook contains controller shapes that refer to `HicController.*`, while the extracted standard-module name is `Module4`. This should be resolved and tested in a controlled copy before distribution. The source workbook was not modified during this review.

## Usage examples

### Operator workflow

```text
Open controller
  -> Validate Setup
  -> Update Without Email
  -> Review saved report
  -> Update and Save Draft
  -> Review and send Outlook draft manually
```

### Public controller procedures

```vb
RunHicValidate   ' Check network/report setup only
RunHicTest       ' Update and save the report without email
RunHicDraft      ' Update, save, and create an Outlook draft
RunHicAll        ' Current controller implementation routes to the draft workflow
```

## Business benefits

- Reduces manual file searching and monthly filename handling.
- Standardizes the reporting period entered into the report workbook.
- Reuses the report's existing data-retrieval process instead of duplicating report logic in the controller.
- Reduces accidental edits to the user's active Excel session by using a separate worker instance.
- Reduces duplicate concurrent runs through shared file locks.
- Provides visible status, selected-file, and log information for operational follow-up.
- Adds a controlled email handoff through Outlook Drafts instead of silently sending a message from the visible controller.
- Performs basic non-empty-result and save confirmations before completing the workflow.

For the Monday stakeholder/LOB distribution process, the reported elapsed time improved from about 10 minutes to about 3–4 minutes per cycle. That represents an approximate 60–70% reduction in manual turnaround time, subject to network, Excel, database, and Outlook response times.

The code does not prove the business correctness of every returned record. The controller performs basic result checks; detailed data-quality assurance remains dependent on the external report workbook and its source systems.

## Security and data-hygiene guidance

The supplied workbook should not be published unchanged. Remove or replace:

- Internal UNC network paths and folder names.
- SQL Server, instance, database, provider, and connection details.
- Real To, CC, sender, and distribution-list addresses.
- Company legal footer, office address, internal email signature, and internal links unless publication is authorized.
- Local Windows usernames, computer names, log paths, and run-history files.
- Workbook names that reveal internal reporting processes.
- Any report data, sample IDs, Pivot output, or operational records.
- Excel document metadata such as creator, last editor, company, and revision information.
- SharePoint or other internal absolute-path metadata.
- `.hic-state` lock files, intent markers, and logs.

The extracted code contains no plaintext database password. It explicitly uses Windows Integrated Security and includes a comment instructing maintainers not to put a password in the workbook or an environment variable. Database access should still be configured through an approved enterprise deployment process.

## Recommended repository structure

```text
.
├── README.md
├── LICENSE
├── .gitignore
├── docs/
│   ├── WORKBOOK_ANALYSIS.md
│   ├── architecture.md
│   ├── setup-and-operations.md
│   └── screenshots/
├── src/
│   ├── controller/
│   │   └── HicController.bas
│   └── report/
│       └── HicAutomation.bas
├── templates/
│   └── HIC-UPDATE.template.xlsm
├── tests/
│   └── test-plan.md
└── examples/
    └── sanitized-run-log.txt
```

### Files suitable for GitHub

For a public repository, upload documentation, sanitized exported `.bas` files, a redacted architecture description, a test plan, and screenshots with all internal values removed. A sanitized template may be included if it contains no live connections, data, recipients, internal paths, or confidential metadata.

Keep the live controller workbook, monthly report workbooks, extracted report data, logs, `.hic-state` files, internal configuration, and production database details in a private repository or controlled release artifact. Do not commit real credentials or operational exports.

## Screenshots to include

Recommended screenshots for a portfolio repository:

1. **Controller setup**: the `HIC Controller` sheet with placeholder network path, year-folder option, status fields, and three buttons.
2. **Validation result**: successful validation with the selected path redacted.
3. **Test result**: successful report update showing the saved status and a redacted Query/Pivot result.
4. **Outlook draft**: draft subject, attachment, and folder link with recipients and internal paths redacted.
5. **Operational log**: a small, sanitized excerpt showing start, report selection, refresh, save, and draft status.
6. **Architecture diagram**: controller, hidden Excel worker, monthly report workbook, network share, and Outlook Drafts.

Do not show internal server names, email addresses, SQL names, report data, local usernames, or legal signatures.

## Known limitations and release checks

- The supplied controller workbook contains an apparently blank `Sheet1` with a legacy `RUN HIC AUTOMATION` shape.
- The visible controller shapes reference `HicController.*`, but the extracted standard module is named `Module4`. Verify and correct the binding in a controlled release copy.
- Three duplicate `HicAutomation` modules are embedded. Confirm the intended report-side module and remove duplicate copies before publishing.
- No VBA UserForms were found. The user interface is built from worksheet cells and shapes.
- No self-contained scheduler was found in the controller. The `SCHEDULE_TIME_ZONE` and scheduled mode belong to the embedded report-side automation path, not the visible controller buttons.
- The controller depends on an external monthly report workbook and enterprise services. It cannot be meaningfully run from a clean GitHub clone without those dependencies.
- Classic Outlook COM automation is used. New Outlook compatibility should be tested separately.
- A static VBA inspection tool reported a possible VBA-stomping condition. This is a release-review signal, not proof of malicious behavior. Compile and compare the VBA source and P-code in a trusted Excel/VBE environment before public distribution.

## Suggested future improvements

These are recommendations, not features claimed to be implemented:

- Resolve and test all shape-to-macro bindings, including the `Module4`/`HicController` naming mismatch.
- Remove duplicate report-side modules and empty legacy modules.
- Move environment-specific values to a redacted configuration layer or deployment process.
- Add boundary tests for month changes, year changes, missing reports, duplicate reports, and failed network access.
- Add structured log levels and a clear operator reconciliation procedure for interrupted runs.
- Add a stronger data-quality report that distinguishes empty data, stale data, formula errors, and successful refresh.
- Use digital signing and a trusted deployment process for production macros.
- Provide a controlled scheduler outside the workbook, with least-privilege permissions and explicit approval for any automatic send.
- Export VBA modules to text for source control and keep binary workbook artifacts in a private release channel.

## Resume and portfolio description

Built an Excel VBA automation controller for a recurring Monday HIC operational reporting process distributed to stakeholders and Lines of Business. The solution locates and prepares monthly report workbooks on a network share, runs the existing report refresh process in an isolated Excel instance, validates and saves the output, writes operational logs, and creates a reviewable Outlook draft. It reduced reported turnaround time from approximately 10 minutes to 3–4 minutes per cycle. The project demonstrates Excel VBA, COM automation, Windows file handling, process coordination, reporting controls, and workflow-focused problem solving.

## License and data notice

Add a license only after confirming that the code and business process can be published. The original workbook contains enterprise-specific configuration and should be treated as confidential until it has been sanitized and approved for release.

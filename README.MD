# CA-CP: Convert XML files to Excel

A PowerShell tool that reads TMS CP **batch**, **alarm** and **sequence** XML
files from a folder and exports their content to Excel files, **one Excel file
per month**.

The `Type` property in the configuration file (`Batch`, `Alarm` or `Sequence`)
selects the worksheets, headers and row layout. Everything else is shared, so
all three types run the same code with a different configuration file.

## 🌟 How it works

Each run does four things:

1. **Read the months.** Every XML file in the folder is opened and the month of
   each record is determined.
2. **Group per month.** Files are grouped by month, one group per Excel file.
3. **Write the Excel files.** Each group is written to the Excel file of its
   month, appending to it when it already exists.
4. **Archive.** Every XML file that was written without a single error is moved
   to the `Archive` sub folder.

Afterwards a summary e-mail is sent, log files are written to a dated folder and
an entry is added to the Windows event log.

### One XML file can fill more than one Excel file

The month is decided **per record**, not per file:

| Type     | Month decided per | On                                    |
| -------- | ----------------- | ------------------------------------- |
| Batch    | delivery          | `deliveryHeader.load_start_date`      |
| Alarm    | alarm             | `alarm.raised`                        |
| Sequence | batch computer    | `batchComputerHeader.file_created_on` |

An XML file holding records of two months is therefore written to two Excel
files, each getting only the records of its own month. A batch file with one
delivery loaded on the last day of the month and another loaded the next day is
the common case.

A delivery is never split: its batches, discharging operations and batch items
follow their delivery. A row on the `plantBatchComputers` worksheet is written
only for a batch computer that actually has records in that month.

### Nothing is archived unless it fully succeeded

An XML file is moved to the `Archive` sub folder only when all of its records
were written without a single error. If anything fails, the file stays where it
is and is listed in the summary e-mail under **"XML files NOT archived, manual
action required"**, so the problem can be investigated and the file picked up
again on the next run.

### The same file name delivered twice

An XML file keeps its name when it is archived, so a file delivered a second
time collides with the copy already there. What happens then depends on the
content, not on the name:

- **Same content** — the file was simply delivered twice and nothing was lost,
  because its data is already in the Excel file. It is archived as
  `Duplicate_<yyyyMMdd_HHmmss>_<name>.xml`, so the archive never overwrites an
  older copy.
- **Different content** — a different file that happens to share a name. Its
  records were *not* added to the Excel file, because a file is recognized by
  its name and that name was already present. The file stays where it is and is
  reported for a manual action, because overwriting the archived copy would
  destroy the only proof of what changed.

### An XML file is never added twice

An XML file is added to an Excel file completely or not at all, so the file name
in column A of `plantBatchComputers` tells the script whether that file is
already in that workbook. The same file name can appear in two different monthly
Excel files, but only once in each.

### A file that does not fit is not written

Before writing, the script checks that every row still fits within the maximum
number of rows of a worksheet. If it does not, nothing is written for that file
and it is reported as an error rather than leaving half a file in the workbook.

## 🚀 Getting started

### Prerequisites

- **PowerShell 7.x or later.**
- **ImportExcel** — the *only* external PowerShell module dependency:
  ```powershell
  Install-Module -Name 'ImportExcel' -Scope 'AllUsers'
  ```
- **MailKit and MimeKit**, only when e-mail is used. These are .NET assemblies
  loaded with `Add-Type`, not PowerShell modules, so ImportExcel stays the only
  module dependency. Their paths go in the configuration file under
  `Settings.SendMail.AssemblyPath`:
  ```powershell
  Install-Package -Name 'MailKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  Install-Package -Name 'MimeKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  ```

Writing to the Windows event log, writing the log files and building and sending
the summary e-mail are done by the module's own helper functions in `Private`,
so no other module is needed.

### Usage

Review the samples in the **`Examples`** folder, create your own configuration
file, then run the entrypoint:

```powershell
pwsh.exe -Command "& '.\Scripts\Entrypoints\ConvertXmlToExcel.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit `$LASTEXITCODE"
```

Use one scheduled task per type, each with its own configuration file, XML
folder and Excel folder.

The exit code is `1` when a fatal error occurred, `0` otherwise. Warnings, such
as a single XML file that could not be read, do not fail the run: the other
files are still converted and the problem is reported in the e-mail and the log.

## 📁 Layout

```
CA-CP-convert-XML-to-Excel/
├── Examples/                       Configuration templates (Batch, Alarm, Sequence)
├── Modules/
│   └── ConvertXmlToExcel/
│       ├── ConvertXmlToExcel.psd1  Module manifest
│       ├── ConvertXmlToExcel.psm1  Loader (dot-sources Private + Public)
│       ├── Public/
│       │   └── Invoke-ConvertXmlToExcel.ps1   Orchestrator (grouping, archiving, logging, mail)
│       └── Private/                Grouped per subject:
│                                     ExcelWorkbook.ps1       open, fill, format the workbooks
│                                     XmlRow.ps1              turn XML nodes into rows
│                                     WorksheetDefinition.ps1 worksheets, headers, colors
│                                     ErrorHandling.ps1       collect errors without throwing
│                                     Utils.ps1               ENV: values, dated log folder
│                                     Logging.ps1             log files and Windows event log
│                                     Mail.ps1                MailKit send, save body to log
│                                     SummaryMail.ps1         subject and HTML body
├── Scripts/
│   ├── Entrypoints/
│   │   └── ConvertXmlToExcel.ps1   Thin bootstrap: imports the module, calls the orchestrator
│   └── Operations/
│       ├── GetXmlFileDates.ps1     Returns every month found in one XML file
│       └── ExportXmlFileToExcel.ps1  Writes one month to one Excel file
├── Tests/                          Pester 6 suite (see Tests/README.md)
└── README.MD
```

The scripts under `Scripts/Operations` are invoked by path and can run in a
parallel runspace, so they dot-source the module's private helpers directly
rather than relying on the module's exports.

## ⚙️ Configuration

Copy one of the `Examples/Example.*.json` files and adjust the values. Every
option is documented inside the file itself, in the `?` block at the bottom:
whether it is mandatory, the accepted values and the default. Most string fields
also accept an `ENV:` reference (for example `"ENV:SMTP_PASSWORD"`), so secrets
never have to be stored in the configuration file.

The file has two halves:

- **`Type`, `MaxConcurrentJobs`, `Path`, `ExcelFileName`** — what to convert and
  where the files live.
- **`Settings`** — how the script reports: `SendMail` (sender, recipients, when
  to send, SMTP server, MailKit/MimeKit assembly paths), `SaveLogFiles` (where
  the dated log folder goes and how long logs are kept) and `SaveInEventLog`
  (whether and where to write Windows event log entries).

A few values are worth calling out:

- **`ExcelFileName`** must contain the placeholder `{0}`, which is replaced by
  the month (for example `202408 August`), and must end with `.xlsx`. Changing
  it once Excel files exist starts new workbooks instead of appending to the
  existing ones.
- **`MaxConcurrentJobs.CreateExcelFile`** should stay at `1` unless every month
  writes to its own Excel file. Two jobs writing the same workbook at the same
  time corrupt it.
- **`SendMail.When`** accepts `Never`, `Always`, `OnlyOnError` or
  `OnlyOnErrorOrAction`.

## 📄 Logging

Each run creates a dated sub folder inside `Settings.SaveLogFiles.Where.Folder`
containing:

| File              | Contents                                                    |
| ----------------- | ----------------------------------------------------------- |
| `Overview.xlsx`   | One row per XML file per month, including its error, if any |
| `Errors.json`     | The errors of the run                                       |
| `Mail - ....html` | A copy of the summary e-mail                                |

`Overview.xlsx` has an `Error` column holding whatever went wrong for that row —
reading the date, writing to Excel or archiving — so a problem is always on the
line of the file it belongs to. `Archived` is `TRUE` or `FALSE`, and `ArchivedAs`
holds the name the file received, which differs from the file name only when it
was archived as a duplicate.

`Errors.json` is only written when there is something to report. Log folders
older than `SaveLogFiles.DeleteLogsAfterDays` are removed at the end of a run.

## 🧪 Tests

```powershell
Invoke-Pester -Path .\Tests
```

See **`Tests/README.md`** for the layout and conventions. `Fanout.Tests.ps1`
covers the behaviour end to end: one XML file with deliveries in two months
produces two Excel files, each holding its own delivery, and a second run adds
nothing.

## 📜 License

Provided as-is, for private and professional use, without any warranty. You are
responsible for the code you deploy and its consequences.

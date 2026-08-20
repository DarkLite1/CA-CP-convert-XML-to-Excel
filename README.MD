# CA-CP: Convert XML files to Excel

A PowerShell tool that reads TMS CP **batch**, **alarm** and **sequence** XML
files from a folder and exports their content to Excel files, **one Excel file
per month**. A single XML file that holds records of more than one month is
written to more than one Excel file, each getting only the records of its own
month.

This repository merges three former repositories into one:

- `CA-CP-convert-batch-XML-file-to-Excel`
- `CA-CP-convert-alarm-XML-file-to-Excel`
- `CA-CP-convert-sequence-XML-file-to-Excel`

The differences between the three are now data, not code: the `Type` property in
the configuration file (`Batch`, `Alarm` or `Sequence`) selects the worksheets,
headers and row layout. Everything else is shared.

## 🌟 What changed

### One XML file can fill more than one Excel file

The old scripts read a single date per XML file and wrote the whole file to the
Excel file of that month. That was wrong for a file holding records of more than
one month — for example a batch file with one delivery loaded today and another
loaded tomorrow, where tomorrow is the first day of the next month. Now the
month is decided per record:

| Type     | Month decided per | On                                    |
| -------- | ----------------- | ------------------------------------- |
| Batch    | delivery          | `deliveryHeader.load_start_date`      |
| Alarm    | alarm             | `alarm.raised`                        |
| Sequence | batch computer    | `batchComputerHeader.file_created_on` |

A delivery is never split: its batches, discharging operations and batch items
follow their delivery. A row on the `plantBatchComputers` worksheet is written
only for a batch computer that actually has records in that month.

### Archiving moved to the orchestrator, and only on success

Because one XML file can now feed more than one Excel file, it may only be
archived after every Excel file has been handled — so archiving moved out of the
export step and into `Invoke-ConvertXmlToExcel`. A file is archived only when
all of its records were written without a single error. If anything fails, the
file stays in place and is listed in the summary mail under **"XML files NOT
archived, manual action required"**, so it can be investigated and retried on
the next run. Because the run is daily, this forces a manual look at problems
instead of silently moving a half-processed file away.

### Duplicate check

An XML file is added to an Excel file completely or not at all, so the file name
in column A of `plantBatchComputers` is enough to know whether it is already
there. The same file name may appear in two different monthly Excel files, but
only once in each.

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
the summary e-mail are all done by the module's own helper functions, taken from
the [Permission-matrix](https://github.com/DarkLite1/Permission-matrix-set-permissions)
project. There is no dependency on `Toolbox.EventLog` or `Toolbox.HTML`.

### Usage

Review the samples in the **`Examples`** folder, create your own configuration
file, then run the entrypoint:

```powershell
pwsh.exe -Command "& '.\Scripts\Entrypoints\ConvertXmlToExcel.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit `$LASTEXITCODE"
```

One scheduled task per type, each with its own configuration file, XML folder and
Excel folder.

## 📁 Layout

```
CA-CP-convert-XML-file-to-Excel/
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
│                                     ErrorHandling.ps1       Add-ErrorHC
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

The operation scripts under `Scripts/Operations` are invoked by path, possibly
in a parallel runspace, so they dot-source the module's private helpers directly
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

Changes compared with the old scripts:

- **`Type`** added — one of `Batch`, `Alarm` or `Sequence`.
- **`MaxConcurrentJobs.GetXmlFileLoadingDate`** / `.GetAlarmRaisedDate` /
  `.GetXmlFileCreationDate` are replaced by a single
  **`MaxConcurrentJobs.GetXmlFileDates`**.
- **`SendMail`** moved under **`Settings`** and gained `From`, `Smtp` and
  `AssemblyPath`, because mail is no longer sent by an internal module.

`ExcelFileName` and the worksheet and table names are unchanged, so the tool
keeps appending to the Excel files created by the previous version.

## 🧪 Tests

```powershell
Invoke-Pester -Path .\Tests
```

See **`Tests/README.md`** for the layout and conventions. The integration test
proves the fanout end to end: one XML file with deliveries in two months
produces two Excel files, each holding its own delivery, and a second run adds
nothing.

## 📜 License

Provided as-is, for private and professional use, without any warranty. You are
responsible for the code you deploy and its consequences.

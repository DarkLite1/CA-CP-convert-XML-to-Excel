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
- **ImportExcel** for reading and writing `.xlsx` files:
  ```powershell
  Install-Module -Name 'ImportExcel' -Scope 'AllUsers'
  ```
- The internal **`Toolbox.EventLog`** and **`Toolbox.HTML`** modules, used by the
  orchestrator for event logging and mail.

### Usage

Review the samples in the **`Examples`** folder, create your own configuration
file, then run the entrypoint:

```powershell
pwsh.exe -Command "& '.\Scripts\Entrypoints\ConvertXmlToExcel.ps1' -ScriptName 'CP Batch' -ImportFile 'C:\MyConfig.json'; exit `$LASTEXITCODE"
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
│       │   └── Invoke-ConvertXmlToExcel.ps1   Orchestrator (grouping, archiving, mail)
│       └── Private/                One function per file (helpers, row builders, workbook engine)
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

Copy one of the `Examples/Example.*.json` files and adjust the values. Changes
compared with the old scripts:

- **`Type`** added — one of `Batch`, `Alarm` or `Sequence`.
- **`MaxConcurrentJobs.GetXmlFileLoadingDate`** / `.GetAlarmRaisedDate` /
  `.GetXmlFileCreationDate` are replaced by a single
  **`MaxConcurrentJobs.GetXmlFileDates`**.

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

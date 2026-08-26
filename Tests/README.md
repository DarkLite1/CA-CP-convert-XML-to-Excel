# Tests

Pester 6 test suite for the ConvertXmlToExcel module. PowerShell 7+ required.

## Folder layout

```
Tests/
├── Helpers/       Shared fixtures, dot-sourced by unit and integration tests
├── TestData/      Sample XML files
├── Unit/          One test file per source file, everything external is faked
│   ├── Private/
│   └── Public/
└── Integration/   Runs the operation scripts and the orchestrator for real
```

### Unit/

Mirrors the production source tree. A source file at
`Modules/ConvertXmlToExcel/Private/Foo.ps1` has its tests at
`Tests/Unit/Private/Foo.Tests.ps1`. Unit tests dot-source the private files they
exercise (and any direct dependencies), then call the function directly.
Anything external — a real Excel workbook, the filesystem outside `TestDrive:` —
is replaced by a small fake built in the test's `BeforeAll`.

### Integration/

Tests that run the real thing end to end: the operation scripts, or the whole
orchestrator against real folders and real `.xlsx` files. They are slower than
the unit tests and prove that the pieces fit together.

`Fanout.Tests.ps1` builds a batch XML file with deliveries in two different
months, runs `GetXmlFileDates.ps1` and `ExportXmlFileToExcel.ps1` exactly as the
orchestrator does, and reads the resulting `.xlsx` files back to prove:

- one XML file produces two monthly Excel files,
- each file holds only its own month's delivery,
- a second run adds nothing (the duplicate check works per file per workbook),
- a month with no records produces no rows.

`Batch.Tests.ps1`, `Alarm.Tests.ps1` and `Sequence.Tests.ps1` run the whole
orchestrator once per type. They are deliberately the same four scenarios, so a
difference between the types shows up as a difference in the results and not in
the way they are tested:

- one file with one record,
- one file with records in two months, which has to produce two Excel files,
- the same file delivered twice, archived as a duplicate,
- a file without a usable date, left behind and reported.

The month is taken from a different place per type: the delivery's
`load_start_date` for batch, the alarm's `raised` for alarm, and the batch
computer's `file_created_on` for sequence.

`Overview.Tests.ps1` runs `Invoke-ConvertXmlToExcel` and reads `Overview.xlsx`
back, covering what happens to a file after conversion:

- a file that converts and is archived,
- the same file delivered again with the same content, archived as a duplicate,
- the same name with different content, left behind for a manual action,
- a file without a usable date,
- a file with records in two months.

### Helpers/

`Fixtures.Xml.ps1` provides `New-BatchXmlHC` and `New-AlarmXmlHC`, which build
minimal XML documents in the exact shape the row builders consume, including the
odd cases (a file spanning two months, a record with no date). It sits at the
`Tests` root because both the unit and the integration tests use it.

## Running

```powershell
# Full suite
Invoke-Pester -Path .\Tests

# Fast feedback (unit only)
Invoke-Pester -Path .\Tests\Unit

# The slower end to end tests
Invoke-Pester -Path .\Tests\Integration

# One file, detailed output
Invoke-Pester -Path .\Tests\Unit\Private\Get-XmlFileMonthHC.Tests.ps1 -Output Detailed
```

## Conventions

- **Test files end with `.Tests.ps1`** (Pester's default discovery pattern).
- **No spaces in filenames.** The source is already space-free
  (`GetXmlFileDates.ps1`), and each test matches its source name.
- **Pester 6 assertions.** This suite uses the `Should-Be`, `Should-BeTrue`,
  `Should-BeCollection`, `Should-Throw`, `Should-MatchString`, `Should-HaveType`
  (and related) operators, with
  `#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }` at the
  top of every file.
- **One Context per behaviour, one It per concrete claim.** Setup that makes a
  failure readable at the point of failure beats sharing it away into helpers.

## What is not covered

`Invoke-ConvertXmlToExcel` has two kinds of coverage. `Tests/Unit/Public`
checks its contract: that it is exported, that ImportExcel is its only external
module dependency, and which parameters are mandatory.
`Tests/Integration/Overview.Tests.ps1` runs it for real against temporary
folders.

What is not covered by any test: sending mail with `Send-MailKitMessageHC` and
writing to the Windows event log with `Write-EventLogSafeHC`. Both need a real
SMTP server and a Windows event log, so they are exercised in the real
environment instead.

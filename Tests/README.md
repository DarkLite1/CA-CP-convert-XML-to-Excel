# Tests

Pester 6 test suite for the ConvertXmlToExcel module. PowerShell 7+ required.

## Folder layout

```
Tests/
├── Helpers/       Shared fixtures (dot-sourced by tests)
├── TestData/      Sample XML files used by the integration tests
├── Unit/          One test file per source file; mocks/fakes everything external
│   ├── Private/
│   └── Public/
└── Integration/   Multi-component scenarios; runs the operation scripts against real .xlsx
```

### Unit/

Mirrors the production source tree. A source file at
`Modules/ConvertXmlToExcel/Private/Foo.ps1` has its tests at
`Tests/Unit/Private/Foo.Tests.ps1`. Unit tests dot-source the private files they
exercise (and any direct dependencies), then call the function directly.
Anything external — a real Excel workbook, the filesystem outside `TestDrive:` —
is replaced by a small fake built in the test's `BeforeAll`.

### Integration/

`Fanout.Tests.ps1` is the end-to-end safety net. It builds a batch XML file with
deliveries in two different months, runs `GetXmlFileDates.ps1` and
`ExportXmlFileToExcel.ps1` exactly as the orchestrator does, and reads the
resulting `.xlsx` files back with `Import-Excel` to prove:

- one XML file produces two monthly Excel files,
- each file holds only its own month's delivery,
- a second run adds nothing (the duplicate check works per file per workbook),
- a month with no records produces no rows.

### Helpers/

`Fixtures.Xml.ps1` provides `New-BatchXmlHC` and `New-AlarmXmlHC`, which build
minimal XML documents in the exact shape the row builders consume, including the
odd cases (a file spanning two months, a record with no date).

## Running

```powershell
# Full suite
Invoke-Pester -Path .\Tests

# Fast feedback (unit only)
Invoke-Pester -Path .\Tests\Unit

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

`Invoke-ConvertXmlToExcel` (the orchestrator) is only checked for its public
contract — that it is exported, that ImportExcel is its only external module
dependency, and which parameters are mandatory. Its full run needs the Windows
event log, the MailKit/MimeKit assemblies and Windows file shares, so it is
exercised in the real environment rather than in unit tests. Everything it
orchestrates is covered directly: `Get-XmlFileMonthHC` and `Get-XmlRowHC` for
the grouping, `Fanout.Tests.ps1` for the export end to end, and
`Utils.Tests.ps1`, `ErrorHandling.Tests.ps1`, `Mail.Tests.ps1` and
`SummaryMail.Tests.ps1` for the bundled helpers.

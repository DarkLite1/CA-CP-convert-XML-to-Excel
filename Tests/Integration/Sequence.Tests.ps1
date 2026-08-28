#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Modules ImportExcel

<#
    The sequence type end to end.

    Sequence is the only type whose month is decided per batch computer, on
    'batchComputerHeader.file_created_on', instead of per record inside a batch
    computer. The other two types are covered by Fanout.Tests.ps1 and
    Overview.Tests.ps1, which both run batch files.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\.."

    Import-Module "$root\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psd1" -Force

    . "$root\Tests\Helpers\Fixtures.Xml.ps1"
    . "$root\Modules\ConvertXmlToExcel\Private\Get-MonthKeyHC.ps1"

    $script:scriptPath = @{
        GetXmlFileDates = "$root\Scripts\Operations\GetXmlFileDates.ps1"
        ExportToExcel   = "$root\Scripts\Operations\ExportXmlFileToExcel.ps1"
    }
    $script:modulePath = "$root\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psm1"

    function New-SequenceRunHC {
        $testRoot = Join-Path $TestDrive "seq_$(New-Guid)"

        $run = @{
            Xml     = Join-Path $testRoot 'xml'
            Excel   = Join-Path $testRoot 'excel'
            Log     = Join-Path $testRoot 'log'
            Config  = Join-Path $testRoot 'config.json'
            Archive = Join-Path $testRoot 'xml\Archive'
        }

        foreach ($folder in $run.Xml, $run.Excel, $run.Log) {
            $null = New-Item -Path $folder -ItemType 'Directory' -Force
        }

        @{
            Type              = 'Sequence'
            MaxConcurrentJobs = @{ GetXmlFileDates = 1; CreateExcelFile = 1 }
            Path              = @{ XmlFiles = $run.Xml; ExcelFiles = $run.Excel }
            ExcelFileName     = 'Sequences - {0}.xlsx'
            <#
                Zero, because these tests write their XML file and convert it in
                the same breath. A file is normally left alone for a few seconds
                after it was last written to, in case it is still arriving, and
                the tests would find nothing to convert. The waiting itself is
                covered in Overview.Tests.ps1.
            #>
            SkipFilesModifiedWithinSeconds = 0
            Settings          = @{
                ScriptName     = 'CP Sequence Test'
                SendMail       = @{
                    From         = 'a@b.com'
                    To           = @('x@y.com')
                    Bcc          = @()
                    When         = 'Never'
                    Smtp         = @{
                        ServerName = 'smtp'; Port = 25
                        ConnectionType = 'None'; UserName = ''; Password = ''
                    }
                    AssemblyPath = @{ MailKit = 'C:\m.dll'; MimeKit = 'C:\k.dll' }
                }
                SaveLogFiles   = @{ Where = @{ Folder = $run.Log }; DeleteLogsAfterDays = 0 }
                SaveInEventLog = @{ Save = $false; LogName = 'Scripts' }
            }
        } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $run.Config -Encoding UTF8

        $run
    }

    function Invoke-SequenceRunHC {
        param ($Run)

        $errors = [System.Collections.Generic.List[object]]::new()

        Invoke-ConvertXmlToExcel -ConfigurationJsonFile $Run.Config `
            -ScriptPath $scriptPath -ModulePath $modulePath `
            -SystemErrors ([ref]$errors) -WarningAction 'SilentlyContinue'

        $Run.Errors = $errors
    }

    function Get-RowMonthHC {
        <#
            .SYNOPSIS
                The month key of a date read back from an Excel file.

            .DESCRIPTION
                Import-Excel returns a date either as a [datetime] or as the raw
                OLE Automation serial number of the cell, depending on the
                number format that was applied and on the ImportExcel version.
                Both have to be handled, otherwise the test fails with
                'String 45464.1633680556 was not recognized as a valid DateTime'
                on one machine and passes on another.
        #>
        param ($Value)

        $date = if ($Value -is [datetime]) {
            $Value
        }
        else {
            [datetime]::FromOADate([double]$Value)
        }

        Get-MonthKeyHC -Date $date
    }

    function Get-SequenceRowHC2 {
        param ($Run, [String]$MonthKey)

        @(
            Import-Excel -Path (Join-Path $Run.Excel "Sequences - $MonthKey.xlsx") `
                -WorksheetName 'sequences' -StartRow 2
        )
    }
}

AfterAll {
    Remove-Module ConvertXmlToExcel -Force -ErrorAction Ignore
}

Describe 'A sequence file with one batch computer' {
    BeforeAll {
        $script:run = New-SequenceRunHC

        (New-SequenceXmlHC -FileCreatedOn '2024-06-21T03:55:15+02:00').Save(
            (Join-Path $run.Xml 'Sequence.xml'))

        Invoke-SequenceRunHC -Run $run
    }

    It 'converts without errors' {
        $run.Errors | Should-BeCollection -Count 0
    }

    It 'creates the Excel file of the month the batch computer was created in' {
        Test-Path (Join-Path $run.Excel 'Sequences - 202406 June.xlsx') | Should-BeTrue
    }

    It 'writes the sequence rows' {
        Get-SequenceRowHC2 -Run $run -MonthKey '202406 June' |
        Should-BeCollection -Count 2
    }

    It 'archives the XML file' {
        Test-Path (Join-Path $run.Archive 'Sequence.xml') | Should-BeTrue
    }
}

Describe 'A sequence file with batch computers in two months' {
    BeforeAll {
        $script:run = New-SequenceRunHC

        <#
            The month is decided per batch computer, so one file with two batch
            computers created in different months has to end up in two Excel
            files, each holding only its own batch computer.
        #>
        (New-SequenceXmlHC -FileCreatedOn '2024-06-21T03:55:15+02:00', '2024-07-02T08:00:00+02:00').Save(
            (Join-Path $run.Xml 'TwoMonths.xml'))

        Invoke-SequenceRunHC -Run $run
    }

    It 'creates one Excel file per month' {
        @(Get-ChildItem $run.Excel -Filter '*.xlsx') | Should-BeCollection -Count 2
    }

    It 'writes only the June batch computer to the June file' {
        $rows = Get-SequenceRowHC2 -Run $run -MonthKey '202406 June'

        $rows | Should-BeCollection -Count 2

        foreach ($row in $rows) {
            Get-RowMonthHC -Value $row.fileCreatedOn | Should-Be '202406 June'
        }
    }

    It 'writes only the July batch computer to the July file' {
        $rows = Get-SequenceRowHC2 -Run $run -MonthKey '202407 July'

        $rows | Should-BeCollection -Count 2

        foreach ($row in $rows) {
            Get-RowMonthHC -Value $row.fileCreatedOn | Should-Be '202407 July'
        }
    }

    It 'archives the file only once, after both months were written' {
        @(Get-ChildItem $run.Archive -Filter '*.xml') | Should-BeCollection -Count 1
    }

    It 'raises no error' {
        $run.Errors | Should-BeCollection -Count 0
    }
}

Describe 'A sequence file that is delivered twice' {
    BeforeAll {
        $script:run = New-SequenceRunHC

        (New-SequenceXmlHC -FileCreatedOn '2024-06-21T03:55:15+02:00').Save(
            (Join-Path $run.Xml 'Sequence.xml'))
        Invoke-SequenceRunHC -Run $run

        (New-SequenceXmlHC -FileCreatedOn '2024-06-21T03:55:15+02:00').Save(
            (Join-Path $run.Xml 'Sequence.xml'))
        Invoke-SequenceRunHC -Run $run
    }

    It 'does not write the rows a second time' {
        Get-SequenceRowHC2 -Run $run -MonthKey '202406 June' |
        Should-BeCollection -Count 2
    }

    It 'archives the second delivery as a duplicate' {
        @(Get-ChildItem $run.Archive -Filter 'Duplicate_*.xml') |
        Should-BeCollection -Count 1
    }
}

Describe 'A sequence file without a creation date' {
    BeforeAll {
        $script:run = New-SequenceRunHC

        (New-SequenceXmlHC -FileCreatedOn '').Save((Join-Path $run.Xml 'NoDate.xml'))

        Invoke-SequenceRunHC -Run $run
    }

    It 'writes no Excel file' {
        @(Get-ChildItem $run.Excel -Filter '*.xlsx') | Should-BeCollection -Count 0
    }

    It 'leaves the file in the folder for a manual action' {
        Test-Path (Join-Path $run.Xml 'NoDate.xml') | Should-BeTrue
    }

    It 'reports both the unreadable date and the file left behind' {
        <#
            Two errors, not one: the date could not be read, and because of that
            the file was not archived. A file staying behind always needs a
            manual action, so it counts as an error of its own.
        #>
        $run.Errors | Should-BeCollection -Count 2

        $run.Errors.Category | Should-ContainCollection 'Date'
        $run.Errors.Category | Should-ContainCollection 'Archive'
    }
}
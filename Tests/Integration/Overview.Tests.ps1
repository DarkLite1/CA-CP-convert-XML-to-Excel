#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Modules ImportExcel

<#
    Runs the orchestrator for real and reads Overview.xlsx back, because the
    overview is built from the results of three separate stages (reading the
    dates, writing to Excel, archiving) and only shows its true shape once
    those have actually run.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\.."

    Import-Module "$root\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psd1" -Force

    . "$root\Tests\Helpers\Fixtures.Xml.ps1"

    $script:scriptPath = @{
        GetXmlFileDates = "$root\Scripts\Operations\GetXmlFileDates.ps1"
        ExportToExcel   = "$root\Scripts\Operations\ExportXmlFileToExcel.ps1"
    }
    $script:modulePath = "$root\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psm1"

    function New-TestRunHC {
        <#
            .SYNOPSIS
                A folder set and a configuration file for one test.

            .PARAMETER MaxConcurrentJobs
                How many jobs the run may use for both stages. The default of 1
                runs everything in this runspace; anything higher runs it in
                parallel runspaces, which is a different code path in the
                orchestrator.
        #>
        param (
            [int]$MaxConcurrentJobs = 1,
            <#
                Zero by default, because these tests write their XML file and
                convert it in the same breath. A file is normally left alone for
                a few seconds after it was last written to, in case it is still
                arriving, and the tests would find nothing to convert.
            #>
            [int]$SkipFilesModifiedWithinSeconds = 0,
            <#
                Leaves the property out of the configuration file altogether,
                for the tests that are about the waiting itself and want the
                default of the script rather than a value of their own.
            #>
            [Switch]$WithoutSkipFilesModifiedWithinSeconds
        )

        $testRoot = Join-Path $TestDrive "run_$(New-Guid)"

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

        $configuration = @{
            Type              = 'Batch'
            MaxConcurrentJobs = @{
                GetXmlFileDates = $MaxConcurrentJobs
                CreateExcelFile = $MaxConcurrentJobs
            }
            Path              = @{ XmlFiles = $run.Xml; ExcelFiles = $run.Excel }
            ExcelFileName     = 'Batches - {0}.xlsx'
            Settings          = @{
                ScriptName     = 'CP Batch Test'
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
        }

        if (-not $WithoutSkipFilesModifiedWithinSeconds) {
            $configuration.SkipFilesModifiedWithinSeconds =
            $SkipFilesModifiedWithinSeconds
        }

        $configuration | ConvertTo-Json -Depth 10 |
        Set-Content -LiteralPath $run.Config -Encoding UTF8

        $run
    }

    function Invoke-TestRunHC {
        param ($Run)

        $errors = [System.Collections.Generic.List[object]]::new()

        Invoke-ConvertXmlToExcel -ConfigurationJsonFile $Run.Config `
            -ScriptPath $scriptPath -ModulePath $modulePath `
            -SystemErrors ([ref]$errors) -WarningAction 'SilentlyContinue'

        $Run.Errors = $errors
    }

    function Get-OverviewHC {
        param ($Run)

        $file = Get-ChildItem -LiteralPath $Run.Log -Recurse -File -Filter 'Overview.xlsx' |
        Sort-Object 'LastWriteTime' | Select-Object -Last 1

        @(Import-Excel -Path $file.FullName)
    }
}

AfterAll {
    Remove-Module ConvertXmlToExcel -Force -ErrorAction Ignore
}

Describe 'Overview.xlsx' {
    Context 'a file that converts and archives' {
        BeforeAll {
            $script:run = New-TestRunHC
            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save(
                (Join-Path $run.Xml 'Plant.xml'))

            Invoke-TestRunHC -Run $run
            $script:row = (Get-OverviewHC -Run $run)[0]
        }

        It 'holds the columns of the run' {
            $row.PSObject.Properties.Name | Should-ContainCollection 'File'
            $row.PSObject.Properties.Name | Should-ContainCollection 'Archived'
            $row.PSObject.Properties.Name | Should-ContainCollection 'ArchivedAs'
            $row.PSObject.Properties.Name | Should-ContainCollection 'Error'
        }

        It 'reports Archived as a real boolean, so Excel shows TRUE' {
            <#
                Compared with Should-BeTrue and not with the text 'TRUE':
                Import-Excel returns a real boolean, and '$false -eq ''FALSE'''
                is $false in PowerShell, because the non empty string converts
                to $true. A text comparison would pass for TRUE and quietly fail
                for FALSE.
            #>
            $row.Archived | Should-BeTrue
            $row.Archived | Should-HaveType ([bool])
        }

        It 'names the file it was archived under' {
            $row.ArchivedAs | Should-Be 'Plant.xml'
        }

        It 'leaves the error empty' {
            $row.Error | Should-BeFalsy
        }

        It 'reports the month and the Excel file' {
            $row.MonthKey | Should-Be '202408 August'
            $row.ExcelFile | Should-Be 'Batches - 202408 August.xlsx'
        }

        It 'raises no error' {
            $run.Errors | Should-BeCollection -Count 0
        }
    }

    Context 'the same file delivered again with the same content' {
        BeforeAll {
            $script:run = New-TestRunHC

            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save(
                (Join-Path $run.Xml 'Plant.xml'))
            Invoke-TestRunHC -Run $run

            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save(
                (Join-Path $run.Xml 'Plant.xml'))
            Invoke-TestRunHC -Run $run

            $script:row = (Get-OverviewHC -Run $run)[0]
        }

        It 'reports that the file was already in the Excel file' {
            $row.AlreadyInSheet | Should-BeTrue
            $row.AddedToSheet | Should-BeFalse
        }

        It 'still archives it' {
            $row.Archived | Should-BeTrue
        }

        It 'archives it under a name with the Duplicate prefix' {
            $row.ArchivedAs | Should-MatchString '^Duplicate_\d{8}_\d{6}_Plant\.xml$'
        }

        It 'raises no error, because nothing was lost' {
            $run.Errors | Should-BeCollection -Count 0
        }

        It 'empties the folder, so the file is not seen again next run' {
            @(Get-ChildItem -LiteralPath $run.Xml -Filter '*.xml' -File) |
            Should-BeCollection -Count 0
        }
    }

    Context 'the same file name delivered again with different content' {
        BeforeAll {
            $script:run = New-TestRunHC

            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save(
                (Join-Path $run.Xml 'Plant.xml'))
            Invoke-TestRunHC -Run $run

            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', '2024-08-17T10:00:00').Save(
                (Join-Path $run.Xml 'Plant.xml'))
            Invoke-TestRunHC -Run $run

            $script:row = (Get-OverviewHC -Run $run)[0]
        }

        It 'reports Archived as FALSE' {
            $row.Archived | Should-BeFalse
        }

        It 'leaves ArchivedAs empty' {
            $row.ArchivedAs | Should-BeFalsy
        }

        It 'puts the reason in the Error column of that file' {
            $row.Error | Should-MatchString 'content is different'
            $row.Error | Should-MatchString 'NOT added to the Excel file'
        }

        It 'counts it as an error, so the mail cannot report a clean run' {
            $run.Errors | Should-BeCollection -Count 1
        }

        It 'leaves the file in the folder for a manual action' {
            @(Get-ChildItem -LiteralPath $run.Xml -Filter '*.xml' -File) |
            Should-BeCollection -Count 1
        }
    }

    Context 'a file without a usable date' {
        BeforeAll {
            $script:run = New-TestRunHC
            (New-BatchXmlHC -LoadStartDate '').Save((Join-Path $run.Xml 'NoDate.xml'))

            Invoke-TestRunHC -Run $run
            $script:row = (Get-OverviewHC -Run $run)[0]
        }

        It 'still gets a row, even though it never reached an Excel file' {
            $row.File | Should-Be 'NoDate.xml'
        }

        It 'reports the reason in the Error column' {
            $row.Error | Should-MatchString 'No date could be read'
        }

        It 'is not archived' {
            $row.Archived | Should-BeFalse
        }
    }

    Context 'a delivery that holds no batches' {
        <#
            The month of a batch file is read from the delivery date while the
            rows come from the batches under that delivery, so this file gives a
            month that holds no rows at all. It was read without a problem, so
            it has to be archived like any other converted file.

            Before 'NothingToAdd' existed the orchestrator saw a month that was
            neither added nor already present, left the file in the folder and
            reported it as an error. Nothing about the file would ever change,
            so it was reported again on every following run.
        #>
        BeforeAll {
            $script:run = New-TestRunHC
            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00' -WithoutBatches).Save(
                (Join-Path $run.Xml 'NoBatches.xml'))

            Invoke-TestRunHC -Run $run
            $script:row = (Get-OverviewHC -Run $run)[0]
        }

        It 'reports that there was nothing to add' {
            $row.NothingToAdd | Should-BeTrue
            $row.AddedToSheet | Should-BeFalse
            $row.AlreadyInSheet | Should-BeFalse
            $row.RowsAdded | Should-Be 0
        }

        It 'archives the file' {
            $row.Archived | Should-BeTrue
            $row.ArchivedAs | Should-Be 'NoBatches.xml'
        }

        It 'leaves the error empty' {
            $row.Error | Should-BeFalsy
        }

        It 'raises no error, so it is not reported again on the next run' {
            $run.Errors | Should-BeCollection -Count 0
        }

        It 'empties the folder, so the file is not seen again next run' {
            @(Get-ChildItem -LiteralPath $run.Xml -Filter '*.xml' -File) |
            Should-BeCollection -Count 0
        }
    }

    Context 'a file that is still being written to' {
        <#
            A file that is still being copied into the folder is already there
            to be found while only part of it is on disk, so reading it fails on
            an XML document that has no end. Nothing is wrong with the file: the
            run looked too early. It has to be left alone, without a word, and
            converted whole on the next run.

            The configuration file of this run does not hold
            'SkipFilesModifiedWithinSeconds' at all, so this also proves the
            default of five seconds that a configuration file written before the
            property existed falls back on.
        #>
        BeforeAll {
            $script:run = New-TestRunHC -WithoutSkipFilesModifiedWithinSeconds

            $script:xmlPath = Join-Path $run.Xml 'Arriving.xml'
            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save($xmlPath)

            <#
                Set on the file itself rather than by waiting, so the test says
                what it means and takes no time at all.
            #>
            (Get-Item $xmlPath).LastWriteTime = Get-Date

            Invoke-TestRunHC -Run $run
        }

        It 'converts nothing' {
            @(Get-ChildItem -LiteralPath $run.Excel -Filter '*.xlsx') |
            Should-BeCollection -Count 0
        }

        It 'leaves the file in the folder for the next run' {
            Test-Path -LiteralPath $xmlPath | Should-BeTrue
        }

        It 'says nothing about it, because there is nothing wrong' {
            $run.Errors | Should-BeCollection -Count 0
        }
    }

    Context 'a file that was written long enough ago' {
        BeforeAll {
            $script:run = New-TestRunHC -WithoutSkipFilesModifiedWithinSeconds

            $xmlPath = Join-Path $run.Xml 'Arrived.xml'
            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save($xmlPath)

            (Get-Item $xmlPath).LastWriteTime = (Get-Date).AddMinutes(-1)

            Invoke-TestRunHC -Run $run
            $script:row = (Get-OverviewHC -Run $run)[0]
        }

        It 'converts it' {
            $row.AddedToSheet | Should-BeTrue
        }

        It 'archives it' {
            $row.Archived | Should-BeTrue
        }
    }

    Context 'the waiting time is set to zero' {
        <#
            For a folder the files are written to locally, where there is
            nothing to wait for. This is the value the rest of the tests in this
            file run on, so it is the one path proved on purpose rather than in
            passing: a value of zero in the configuration file is read and used,
            and is not mistaken for a missing property that falls back on the
            default of five.
        #>
        BeforeAll {
            $script:run = New-TestRunHC -SkipFilesModifiedWithinSeconds 0

            $xmlPath = Join-Path $run.Xml 'JustWritten.xml'
            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save($xmlPath)

            (Get-Item $xmlPath).LastWriteTime = Get-Date

            Invoke-TestRunHC -Run $run
            $script:row = (Get-OverviewHC -Run $run)[0]
        }

        It 'converts a file that was just written' {
            $row.AddedToSheet | Should-BeTrue
        }

        It 'raises no error' {
            $run.Errors | Should-BeCollection -Count 0
        }
    }

    Context 'a run in parallel' {
        <#
            The parallel and the sequential path are two different script
            blocks, because a parallel one runs in its own runspace and has to
            pull in what it needs with '$using:' while a sequential one must
            not. Only the sequential path was covered, so a mistake in the
            parallel one showed up in production and nowhere else.

            Each month writes its own workbook here. Two jobs writing the same
            workbook at the same time corrupt it, which is why the configuration
            warns against 'CreateExcelFile' above 1 for anything else.
        #>
        BeforeAll {
            $script:run = New-TestRunHC -MaxConcurrentJobs 2

            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00').Save(
                (Join-Path $run.Xml 'August.xml'))
            (New-BatchXmlHC -LoadStartDate '2024-09-02T08:15:00').Save(
                (Join-Path $run.Xml 'September.xml'))

            Invoke-TestRunHC -Run $run
            $script:rows = Get-OverviewHC -Run $run
        }

        It 'converts both files' {
            $rows | Should-BeCollection -Count 2
            $rows.AddedToSheet | Should-BeCollection @($true, $true)
        }

        It 'writes one Excel file per month' {
            @(Get-ChildItem -LiteralPath $run.Excel -Filter '*.xlsx') |
            Should-BeCollection -Count 2
        }

        It 'archives both files' {
            $rows.Archived | Should-BeCollection @($true, $true)
        }

        It 'raises no error' {
            $run.Errors | Should-BeCollection -Count 0
        }
    }

    Context 'a file with records in two months' {
        BeforeAll {
            $script:run = New-TestRunHC
            (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', '2024-09-02T08:15:00').Save(
                (Join-Path $run.Xml 'TwoMonths.xml'))

            Invoke-TestRunHC -Run $run
            $script:rows = Get-OverviewHC -Run $run
        }

        It 'gets one row per month' {
            $rows | Should-BeCollection -Count 2
        }

        It 'names both Excel files' {
            $rows.ExcelFile | Should-ContainCollection 'Batches - 202408 August.xlsx'
            $rows.ExcelFile | Should-ContainCollection 'Batches - 202409 September.xlsx'
        }

        It 'reports the file as archived on every row' {
            $rows.Archived | Should-BeCollection @($true, $true)
        }
    }
}
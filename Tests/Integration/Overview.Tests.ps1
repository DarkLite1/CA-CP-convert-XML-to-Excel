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
        #>
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

        @{
            Type              = 'Batch'
            MaxConcurrentJobs = @{ GetXmlFileDates = 1; CreateExcelFile = 1 }
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
        } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $run.Config -Encoding UTF8

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

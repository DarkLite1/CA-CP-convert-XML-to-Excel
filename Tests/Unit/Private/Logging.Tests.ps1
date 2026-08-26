#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Modules ImportExcel

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ErrorHandling.ps1"
    . "$moduleRoot\Private\Logging.ps1"
}

Describe 'Out-LogFileHC' {
    BeforeEach {
        $partialPath = Join-Path $TestDrive "log_$(New-Guid)"

        $data = @(
            [PSCustomObject]@{ File = 'A.xml'; RowsAdded = 3; Error = $null }
            [PSCustomObject]@{ File = 'B.xml'; RowsAdded = 0; Error = 'Boom' }
        )
    }

    It 'writes an .xlsx file and returns its path' {
        $result = Out-LogFileHC -DataToExport $data -PartialPath $partialPath -FileExtensions '.xlsx'

        $result | Should-BeTruthy
        Test-Path -LiteralPath $result | Should-BeTrue
    }

    It 'writes the rows so they can be read back' {
        $result = Out-LogFileHC -DataToExport $data -PartialPath $partialPath -FileExtensions '.xlsx'

        $rows = @(Import-Excel -Path $result)

        $rows | Should-BeCollection -Count 2
        $rows[1].Error | Should-Be 'Boom'
    }

    It 'writes a .json file when asked' {
        $result = Out-LogFileHC -DataToExport $data -PartialPath $partialPath -FileExtensions '.json'

        Test-Path -LiteralPath $result | Should-BeTrue

        $rows = @(Get-Content -LiteralPath $result -Raw | ConvertFrom-Json)

        $rows | Should-BeCollection -Count 2
    }

    It 'writes one file per requested extension' {
        $result = @(
            Out-LogFileHC -DataToExport $data -PartialPath $partialPath -FileExtensions '.xlsx', '.json'
        )

        $result | Should-BeCollection -Count 2
    }
}

Describe 'Remove-OldLogsHC' {
    BeforeEach {
        $logFolder = Join-Path $TestDrive "logs_$(New-Guid)"
        $errors = [System.Collections.Generic.List[object]]::new()

        $null = New-Item -Path $logFolder -ItemType 'Directory' -Force
    }

    It 'removes a log file older than the retention' {
        $old = Join-Path $logFolder 'old.xlsx'
        Set-Content -LiteralPath $old -Value 'x'
        (Get-Item $old).CreationTime = (Get-Date).AddDays(-40)

        Remove-OldLogsHC -LogFolder $logFolder -RetentionDays 30 -SystemErrors ([ref]$errors)

        Test-Path -LiteralPath $old | Should-BeFalse
    }

    It 'keeps a log file inside the retention' {
        $recent = Join-Path $logFolder 'recent.xlsx'
        Set-Content -LiteralPath $recent -Value 'x'

        Remove-OldLogsHC -LogFolder $logFolder -RetentionDays 30 -SystemErrors ([ref]$errors)

        Test-Path -LiteralPath $recent | Should-BeTrue
    }

    It 'removes the run folder left empty, keeping the script folder with recent runs' {
        <#
            Runs live in '<LogFolder>\<script>\<run>', so an old run folder has
            to disappear without taking the script folder with it.
        #>
        $scriptFolder = Join-Path $logFolder 'CP Batch'
        $oldRun = Join-Path $scriptFolder '2020_01_01_000000'
        $newRun = Join-Path $scriptFolder '2030_01_01_000000'

        $null = New-Item -Path $oldRun -ItemType 'Directory' -Force
        $null = New-Item -Path $newRun -ItemType 'Directory' -Force

        $oldFile = Join-Path $oldRun 'Overview.xlsx'
        Set-Content -LiteralPath $oldFile -Value 'x'
        (Get-Item $oldFile).CreationTime = (Get-Date).AddDays(-40)

        Set-Content -LiteralPath (Join-Path $newRun 'Overview.xlsx') -Value 'x'

        Remove-OldLogsHC -LogFolder $logFolder -RetentionDays 30 -SystemErrors ([ref]$errors)

        Test-Path -LiteralPath $oldRun | Should-BeFalse
        Test-Path -LiteralPath $newRun | Should-BeTrue
        Test-Path -LiteralPath $scriptFolder | Should-BeTrue
    }

    It 'does nothing when the log folder does not exist' {
        Remove-OldLogsHC -LogFolder (Join-Path $TestDrive 'nope') -RetentionDays 30 `
            -SystemErrors ([ref]$errors)

        $errors | Should-BeCollection -Count 0
    }
}

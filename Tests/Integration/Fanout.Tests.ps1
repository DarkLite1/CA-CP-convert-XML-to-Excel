#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Modules ImportExcel

<#
    Exercises the two operation scripts against real .xlsx files, the way the
    orchestrator runs them. Proves the fanout end to end: one XML file with
    deliveries in two months produces two Excel files, each holding its own
    delivery, and a second run adds nothing.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    $script:modulePath = "$moduleRoot\ConvertXmlToExcel.psm1"
    $script:getDates = "$root\Scripts\Operations\GetXmlFileDates.ps1"
    $script:export = "$root\Scripts\Operations\ExportXmlFileToExcel.ps1"

    . "$moduleRoot\Private\Get-MonthKeyHC.ps1"

    # A batch file with deliveries in August and September
    . "$root\Tests\Helpers\Fixtures.Xml.ps1"

    function New-EmptiedWorkbookHC {
        <#
            .SYNOPSIS
                An Excel file whose plant worksheet lost its header rows.

            .DESCRIPTION
                What a workbook looks like after someone emptied a worksheet by
                hand. Opening it has to fail, and the point of the tests using
                it is what the export script does with that failure.
        #>
        param ([String]$Path)

        $package = Open-ExcelPackage -Path $Path

        try {
            $package.Workbook.Worksheets.Delete('plantBatchComputers')
            $null = $package.Workbook.Worksheets.Add('plantBatchComputers')
        }
        finally { Close-ExcelPackage $package }
    }

    function Get-RowMonthHC {
        param($Value)
        $date = if ($Value -is [datetime]) {
            $Value
        }
        else {
            [datetime]::FromOADate([double]$Value)
        }
        Get-MonthKeyHC -Date $date
    }
}

Describe 'Fanout end to end' {
    BeforeEach {
        $xmlFolder = New-Item -Path (Join-Path $TestDrive "xml_$(New-Guid)") -ItemType Directory
        $excelFolder = New-Item -Path (Join-Path $TestDrive "xlsx_$(New-Guid)") -ItemType Directory

        $xml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', '2024-09-02T08:15:00'
        $xmlPath = Join-Path $xmlFolder.FullName 'BatchTwoMonths.xml'
        $xml.Save($xmlPath)
        $xmlFile = Get-Item $xmlPath
    }

    It 'finds both months in the file' {
        $dates = & $getDates -XmlFile $xmlFile -Type 'Batch' -ModulePath $modulePath

        $dates.Error | Should-BeFalsy
        $dates.MonthKeys | Should-BeCollection @('202408 August', '202409 September')
    }

    It 'writes one delivery to each monthly Excel file' {
        $dates = & $getDates -XmlFile $xmlFile -Type 'Batch' -ModulePath $modulePath

        foreach ($monthKey in $dates.MonthKeys) {
            $xlsx = Join-Path $excelFolder.FullName "Batches - $monthKey.xlsx"

            $result = & $export -XmlFiles $xmlFile -ExcelFilePath $xlsx `
                -Type 'Batch' -MonthKey $monthKey -ModulePath $modulePath

            $result.AddedToSheet | Should-BeTrue
            $result.Error | Should-BeFalsy
        }

        @(Get-ChildItem $excelFolder.FullName -Filter '*.xlsx') |
        Should-BeCollection -Count 2

        foreach ($monthKey in $dates.MonthKeys) {
            $xlsx = Join-Path $excelFolder.FullName "Batches - $monthKey.xlsx"
            $rows = @(Import-Excel -Path $xlsx -WorksheetName 'deliveriesBatches' -StartRow 2)

            $rows | Should-BeCollection -Count 1
            Get-RowMonthHC -Value $rows[0].loadStartDate | Should-Be $monthKey
        }
    }

    It 'adds nothing on a second run of the same month' {
        $monthKey = '202408 August'
        $xlsx = Join-Path $excelFolder.FullName "Batches - $monthKey.xlsx"

        $null = & $export -XmlFiles $xmlFile -ExcelFilePath $xlsx `
            -Type 'Batch' -MonthKey $monthKey -ModulePath $modulePath

        $second = & $export -XmlFiles $xmlFile -ExcelFilePath $xlsx `
            -Type 'Batch' -MonthKey $monthKey -ModulePath $modulePath

        $second.AlreadyInSheet | Should-BeTrue
        $second.AddedToSheet | Should-BeFalse

        $rows = @(Import-Excel -Path $xlsx -WorksheetName 'deliveriesBatches' -StartRow 2)
        $rows | Should-BeCollection -Count 1
    }

    It 'writes nothing for a month the file has no records for' {
        $xlsx = Join-Path $excelFolder.FullName 'Batches - 202401 January.xlsx'

        $result = & $export -XmlFiles $xmlFile -ExcelFilePath $xlsx `
            -Type 'Batch' -MonthKey '202401 January' -ModulePath $modulePath

        $result.AddedToSheet | Should-BeFalse
    }

    Context 'the Excel file cannot be opened' {
        BeforeEach {
            $script:xlsx = Join-Path $excelFolder.FullName 'Batches - 202408 August.xlsx'

            $null = & $export -XmlFiles $xmlFile -ExcelFilePath $xlsx `
                -Type 'Batch' -MonthKey '202408 August' -ModulePath $modulePath

            New-EmptiedWorkbookHC -Path $xlsx

            $script:result = & $export -XmlFiles $xmlFile -ExcelFilePath $xlsx `
                -Type 'Batch' -MonthKey '202408 August' -ModulePath $modulePath `
                -WarningAction 'SilentlyContinue'
        }

        It 'reports the failure per file instead of throwing' {
            <#
                A throw would not fail this month alone. The script runs inside
                the orchestrator's ForEach-Object, so it would end the whole
                pipeline: the results of the months that did succeed never reach
                the caller and not one XML file of the run is archived.
            #>
            $result | Should-BeCollection -Count 1
            $result.Error | Should-MatchString 'Failed opening Excel file'
        }

        It 'reports nothing as written' {
            $result.AddedToSheet | Should-BeFalse
            $result.AlreadyInSheet | Should-BeFalse
            $result.NothingToAdd | Should-BeFalse
        }

        It 'leaves the Excel file closed, so the next run can open it' {
            <#
                The package holds the file until it is disposed. Deleting the
                file is what proves it was released: a file still held open
                cannot be removed.
            #>
            Remove-Item -LiteralPath $xlsx -Force

            Test-Path -LiteralPath $xlsx | Should-BeFalse
        }
    }

    It 'reports NothingToAdd for a delivery that holds no batches' {
        <#
            The month comes from the delivery date while the rows come from the
            batches under it, so a delivery without batches gives a month with
            no rows. The file was read correctly, so this has to be reported as
            a third success state and not as a month that went missing: the
            orchestrator refuses to archive a file whose months were not all
            handled, and no later run would ever change the outcome.
        #>
        $emptyXml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00' -WithoutBatches
        $emptyPath = Join-Path $xmlFolder.FullName 'NoBatches.xml'
        $emptyXml.Save($emptyPath)

        $xlsx = Join-Path $excelFolder.FullName 'Batches - 202408 August.xlsx'

        $result = & $export -XmlFiles (Get-Item $emptyPath) -ExcelFilePath $xlsx `
            -Type 'Batch' -MonthKey '202408 August' -ModulePath $modulePath

        $result.Error | Should-BeFalsy
        $result.AddedToSheet | Should-BeFalse
        $result.AlreadyInSheet | Should-BeFalse
        $result.NothingToAdd | Should-BeTrue
        $result.RowsAdded | Should-Be 0
    }
}
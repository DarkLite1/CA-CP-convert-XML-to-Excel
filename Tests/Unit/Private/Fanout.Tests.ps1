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
}

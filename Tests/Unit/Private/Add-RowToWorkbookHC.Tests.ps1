#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Modules ImportExcel

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ExcelWorkbook.ps1"

    <#
        A real in-memory EPPlus worksheet, never saved to disk.

        A hand written fake does not work here: the function assigns with
        '$sheet.Worksheet.Cells[$address].Value = $value', and a PSCustomObject
        with a ScriptMethod called 'Item' is not a real indexer, so the index
        returns $null and setting '.Value' fails. Using the real object also
        means the test cannot drift away from how EPPlus actually behaves.
    #>
    function New-TestSheetHC {
        param([int]$RowNumber = 3)

        $package = [OfficeOpenXml.ExcelPackage]::new()

        @{
            RowNumber = $RowNumber
            Worksheet = $package.Workbook.Worksheets.Add('Test')
            Package   = $package
        }
    }
}

Describe 'Add-RowToWorkbookHC' {
    It 'returns 0 and writes nothing for an empty row set' {
        $plant = New-TestSheetHC
        $workbook = @{ Sheet = @{ plant = $plant } }

        try {
            Add-RowToWorkbookHC -Workbook $workbook -Rows @() | Should-Be 0

            $plant.RowNumber | Should-Be 3
        }
        finally { $plant.Package.Dispose() }
    }

    It 'writes each cell and advances the row number' {
        $plant = New-TestSheetHC -RowNumber 3
        $workbook = @{ Sheet = @{ plant = $plant } }

        $rows = @(
            @{ Key = 'plant'; Cells = @{ A = 'File.xml'; C = 'BE' } }
        )

        try {
            $added = Add-RowToWorkbookHC -Workbook $workbook -Rows $rows

            $added | Should-Be 1
            $plant.RowNumber | Should-Be 4
            $plant.Worksheet.Cells['A3'].Value | Should-Be 'File.xml'
            $plant.Worksheet.Cells['C3'].Value | Should-Be 'BE'
        }
        finally { $plant.Package.Dispose() }
    }

    It 'writes one row per entry and keeps them in order' {
        $plant = New-TestSheetHC -RowNumber 3
        $workbook = @{ Sheet = @{ plant = $plant } }

        $rows = @(
            @{ Key = 'plant'; Cells = @{ A = 'First.xml' } }
            @{ Key = 'plant'; Cells = @{ A = 'Second.xml' } }
        )

        try {
            Add-RowToWorkbookHC -Workbook $workbook -Rows $rows | Should-Be 2

            $plant.RowNumber | Should-Be 5
            $plant.Worksheet.Cells['A3'].Value | Should-Be 'First.xml'
            $plant.Worksheet.Cells['A4'].Value | Should-Be 'Second.xml'
        }
        finally { $plant.Package.Dispose() }
    }

    It 'throws without writing when the rows do not fit' {
        $plant = New-TestSheetHC -RowNumber 1048576
        $workbook = @{ Sheet = @{ plant = $plant } }

        $rows = @(@{ Key = 'plant'; Cells = @{ A = 'File.xml' } })

        try {
            { Add-RowToWorkbookHC -Workbook $workbook -Rows $rows } |
            Should-Throw '*max row limit reached*'

            $plant.RowNumber | Should-Be 1048576
        }
        finally { $plant.Package.Dispose() }
    }
}
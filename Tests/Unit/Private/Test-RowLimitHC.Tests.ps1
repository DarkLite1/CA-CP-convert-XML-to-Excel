#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ExcelWorkbook.ps1"

    function New-FakeWorkbook {
        param([hashtable]$RowNumbers)
        $sheet = @{}
        foreach ($key in $RowNumbers.Keys) {
            $sheet[$key] = @{ RowNumber = $RowNumbers[$key] }
        }
        @{ Sheet = $sheet }
    }
}

Describe 'Test-RowLimitHC' {
    It 'returns true when the rows still fit' {
        $workbook = New-FakeWorkbook -RowNumbers @{ plant = 3; delivery = 3 }
        $rows = @(
            @{ Key = 'plant'; Cells = @{} }
            @{ Key = 'delivery'; Cells = @{} }
            @{ Key = 'delivery'; Cells = @{} }
        )

        Test-RowLimitHC -Workbook $workbook -Rows $rows | Should-BeTrue
    }

    It 'returns false when one worksheet would pass the max row number' {
        $workbook = New-FakeWorkbook -RowNumbers @{ delivery = 1048576 }
        $rows = @(@{ Key = 'delivery'; Cells = @{} })

        Test-RowLimitHC -Workbook $workbook -Rows $rows | Should-BeFalse
    }

    It 'respects a lowered MaxRowNumber so the limit is easy to test' {
        $workbook = New-FakeWorkbook -RowNumbers @{ delivery = 3 }
        $rows = @(
            @{ Key = 'delivery'; Cells = @{} }
            @{ Key = 'delivery'; Cells = @{} }
            @{ Key = 'delivery'; Cells = @{} }
        )

        # room for two more (3,4) but not three (5 > 4)
        Test-RowLimitHC -Workbook $workbook -Rows $rows -MaxRowNumber 4 | Should-BeFalse
    }

    It 'returns true for an empty set of rows' {
        $workbook = New-FakeWorkbook -RowNumbers @{ plant = 3 }
        Test-RowLimitHC -Workbook $workbook -Rows @() | Should-BeTrue
    }
}

#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Get-ColumnCellHC.ps1"
}

Describe 'Get-ColumnCellHC' {
    It 'returns one address per required column' {
        $result = Get-ColumnCellHC -RowNumber 2 -RequiredColumns 3
        $result | Should-BeCollection -Count 3
    }

    It 'numbers the first columns A, B, C on the requested row' {
        $result = Get-ColumnCellHC -RowNumber 2 -RequiredColumns 3
        $result | Should-BeCollection @('A2', 'B2', 'C2')
    }

    It 'rolls over to AA after Z' {
        $result = Get-ColumnCellHC -RowNumber 1 -RequiredColumns 27
        $result[25] | Should-Be 'Z1'
        $result[26] | Should-Be 'AA1'
    }

    It 'keeps rolling over: column 53 is BA' {
        $result = Get-ColumnCellHC -RowNumber 1 -RequiredColumns 53
        $result[52] | Should-Be 'BA1'
    }
}

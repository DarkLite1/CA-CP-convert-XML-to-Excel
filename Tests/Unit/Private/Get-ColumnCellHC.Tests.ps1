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

    Context 'past the two letter addresses' {
        <#
            A column letter is not quite a number in base 26: there is no zero
            digit, so 'Z' is followed by 'AA'. Counting a letter for the first
            position and a letter for the second runs out at 'ZZ' and starts
            handing out two letter addresses again from there, which are
            addresses of columns that were already written.
        #>
        BeforeAll {
            $script:result = Get-ColumnCellHC -RowNumber 1 -RequiredColumns 705
        }

        It 'ends the two letter addresses at ZZ' {
            $result[701] | Should-Be 'ZZ1'
        }

        It 'goes on with three letters' {
            $result[702] | Should-Be 'AAA1'
            $result[703] | Should-Be 'AAB1'
        }

        It 'never returns the same address twice' {
            @($result | Sort-Object -Unique) | Should-BeCollection -Count 705
        }
    }
}
#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Test-RowLimitHC.ps1"
    . "$moduleRoot\Private\Add-RowToWorkbookHC.ps1"

    # Fake worksheet whose Cells[...] .Value setter records into a shared store
    function New-FakeSheet {
        param([int]$RowNumber = 3)
        $store = [ordered]@{}
        $cells = [PSCustomObject]@{ Store = $store }
        $cells | Add-Member -MemberType ScriptMethod -Name 'Item' -Value {
            param($Address)
            $target = $this.Store
            [PSCustomObject]@{ Address = $Address } |
            Add-Member -PassThru -MemberType ScriptProperty -Name Value `
                -Value { $target[$this.Address] } `
                -SecondValue { param($v) $target[$this.Address] = $v }
        }
        @{
            RowNumber = $RowNumber
            Worksheet = [PSCustomObject]@{ Cells = $cells }
            Store     = $store
        }
    }
}

Describe 'Add-RowToWorkbookHC' {
    It 'returns 0 and writes nothing for an empty row set' {
        $workbook = @{ Sheet = @{ plant = New-FakeSheet } }

        Add-RowToWorkbookHC -Workbook $workbook -Rows @() | Should-Be 0
    }

    It 'writes each cell and advances the row number' {
        $plant = New-FakeSheet -RowNumber 3
        $workbook = @{ Sheet = @{ plant = $plant } }

        $rows = @(
            @{ Key = 'plant'; Cells = @{ A = 'File.xml'; C = 'BE' } }
        )

        $added = Add-RowToWorkbookHC -Workbook $workbook -Rows $rows

        $added | Should-Be 1
        $plant.RowNumber | Should-Be 4
        $plant.Store['A3'] | Should-Be 'File.xml'
        $plant.Store['C3'] | Should-Be 'BE'
    }

    It 'throws without writing when the rows do not fit' {
        $plant = New-FakeSheet -RowNumber 1048576
        $workbook = @{ Sheet = @{ plant = $plant } }

        $rows = @(@{ Key = 'plant'; Cells = @{ A = 'File.xml' } })

        { Add-RowToWorkbookHC -Workbook $workbook -Rows $rows } |
        Should-Throw '*max row limit reached*'
    }
}

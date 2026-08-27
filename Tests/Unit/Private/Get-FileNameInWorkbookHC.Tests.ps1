#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ExcelWorkbook.ps1"

    <#
        A fake plant worksheet. Dimension.End.Row reports the last row and
        GetValue(row, column) returns the value of one cell, matching the shape
        EPPlus exposes. The file names sit in column A from row 3 down, below
        the two header rows.
    #>
    function New-FakePlantWorkbook {
        param([string[]]$FileNames)

        $lastRow = 2 + $FileNames.Count

        $worksheet = [PSCustomObject]@{
            Dimension = [PSCustomObject]@{
                End = [PSCustomObject]@{ Row = $lastRow }
            }
        }

        <#
            GetNewClosure, so the method keeps hold of the names it was built
            with. Without it the script block would look for '$FileNames' when
            it is called, long after this function returned, and find nothing.
        #>
        $worksheet | Add-Member -MemberType 'ScriptMethod' -Name 'GetValue' -Value {
            param ($Row, $Column)

            if ($Column -ne 1) { return $null }
            if ($Row -lt 3) { return $null }

            $FileNames[$Row - 3]
        }.GetNewClosure()

        @{
            PlantKey = 'plant'
            Sheet    = @{ plant = @{ Worksheet = $worksheet } }
        }
    }
}

Describe 'Get-FileNameInWorkbookHC' {
    It 'returns an empty hashtable for a sheet with only headers' {
        $workbook = New-FakePlantWorkbook -FileNames @()

        $result = Get-FileNameInWorkbookHC -Workbook $workbook

        $result.Keys.Count | Should-Be 0
    }

    It 'reads column A only' {
        <#
            The file name lives in column A. Reading any other column would
            report a name that is not there and skip one that is.
        #>
        $workbook = New-FakePlantWorkbook -FileNames @('A.xml')

        $workbook.Sheet['plant'].Worksheet |
        Add-Member -MemberType 'ScriptMethod' -Name 'GetValue' -Force -Value {
            param ($Row, $Column)

            if ($Column -eq 1) { 'A.xml' } else { 'B.xml' }
        }

        $result = Get-FileNameInWorkbookHC -Workbook $workbook

        $result.ContainsKey('A.xml') | Should-BeTrue
        $result.ContainsKey('B.xml') | Should-BeFalse
    }

    It 'returns an empty hashtable for a worksheet with no cells at all' {
        <#
            EPPlus reports no dimension for a worksheet that holds nothing, so
            there is no last row to read. Reading through it threw and took the
            whole run with it.
        #>
        $workbook = @{
            PlantKey = 'plant'
            Sheet    = @{
                plant = @{
                    Worksheet = [PSCustomObject]@{ Dimension = $null }
                }
            }
        }

        $result = Get-FileNameInWorkbookHC -Workbook $workbook

        $result.Keys.Count | Should-Be 0
    }

    It 'reports a file name that is present' {
        $workbook = New-FakePlantWorkbook -FileNames @('A.xml', 'B.xml')

        $result = Get-FileNameInWorkbookHC -Workbook $workbook

        $result['A.xml'] | Should-BeTrue
        $result['B.xml'] | Should-BeTrue
    }

    It 'does not report a file name that is absent' {
        $workbook = New-FakePlantWorkbook -FileNames @('A.xml')

        $result = Get-FileNameInWorkbookHC -Workbook $workbook

        $result.ContainsKey('C.xml') | Should-BeFalse
    }

    It 'collapses the repeated names of one file to a single key' {
        # one XML file produces many plant rows; the lookup should be by name
        $workbook = New-FakePlantWorkbook -FileNames @('A.xml', 'A.xml', 'A.xml')

        $result = Get-FileNameInWorkbookHC -Workbook $workbook

        $result.Keys.Count | Should-Be 1
        $result['A.xml'] | Should-BeTrue
    }
}
#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ExcelWorkbook.ps1"

    <#
        A fake plant worksheet. Dimension.End.Row reports the last row and
        Cells["A3:A{last}"].Value returns the file names in column A, matching
        the shape EPPlus exposes.
    #>
    function New-FakePlantWorkbook {
        param([string[]]$FileNames)

        $lastRow = 2 + $FileNames.Count

        # EPPlus exposes Cells[range].Value; a hashtable keyed by the exact
        # range string reproduces that indexing without a real workbook.
        $cells = @{
            "A3:A$lastRow" = [PSCustomObject]@{ Value = $FileNames }
        }

        $worksheet = [PSCustomObject]@{
            Cells     = $cells
            Dimension = [PSCustomObject]@{
                End = [PSCustomObject]@{ Row = $lastRow }
            }
        }

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
                    Worksheet = [PSCustomObject]@{
                        Cells     = @{}
                        Dimension = $null
                    }
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
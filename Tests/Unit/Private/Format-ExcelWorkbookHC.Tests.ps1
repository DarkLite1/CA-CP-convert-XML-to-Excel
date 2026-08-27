#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Modules ImportExcel

<#
    A real workbook is used throughout: the table and its filter only exist in
    the file once EPPlus has written them, and the point of these tests is what
    a SECOND run does to a table that is already there.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Get-ColumnCellHC.ps1"
    . "$moduleRoot\Private\WorksheetDefinition.ps1"
    . "$moduleRoot\Private\ExcelWorkbook.ps1"

    function Add-PlantRowHC {
        <#
            .SYNOPSIS
                Write one plant row to an open workbook.
        #>
        param ($Workbook, [String]$FileName)

        $null = Add-RowToWorkbookHC -Workbook $Workbook -Rows @(
            @{ Key = 'plant'; Cells = @{ A = $FileName } }
        )
    }

    function New-WorkbookWithTwoRunsHC {
        <#
            .SYNOPSIS
                An Excel file written by two runs, left open after the second.

            .DESCRIPTION
                The first run creates the file and its table, the second appends
                to the table that is already there. The caller closes the
                package it gets back.
        #>
        $path = Join-Path $TestDrive "table_$(New-Guid).xlsx"

        $first = Open-ExcelWorkbookHC -Path $path -Type 'Batch'
        Add-PlantRowHC -Workbook $first -FileName 'One.xml'
        Format-ExcelWorkbookHC -Workbook $first
        Close-ExcelPackage $first.Package

        $second = Open-ExcelWorkbookHC -Path $path -Type 'Batch'
        Add-PlantRowHC -Workbook $second -FileName 'Two.xml'
        Format-ExcelWorkbookHC -Workbook $second

        $second
    }
}

Describe 'Format-ExcelWorkbookHC' {
    Context 'a table that grows on a second run' {
        BeforeEach {
            $script:workbook = New-WorkbookWithTwoRunsHC
            $script:table = $workbook.Sheet['plant'].Worksheet.Tables['plantBatchComputer']
        }

        AfterEach {
            Close-ExcelPackage $workbook.Package -NoSave
        }

        It 'grows the table over the row added by the second run' {
            <#
                Two header rows and two data rows, so the table runs from row 2
                to row 4.
            #>
            $table.TableXml.table.ref | Should-MatchString '^A2:[A-Z]+4$'
        }

        It 'has a filter on the table' {
            <#
                The fix below only matters because this element exists. Should
                this ever fail, the filter is not written at all and keeping the
                two ranges in step is unnecessary.
            #>
            $table.TableXml.table.autoFilter | Should-BeTruthy
        }

        It 'grows the filter along with the table' {
            <#
                Excel expects both ranges to say the same thing. Growing only
                the table left the filter on the range of the previous run, so
                the rows added last fell outside it and Excel could refuse to
                open the workbook without repairing it first.
            #>
            $table.TableXml.table.autoFilter.ref |
            Should-Be $table.TableXml.table.ref
        }
    }

    Context 'a table created by the first run' {
        BeforeEach {
            $path = Join-Path $TestDrive "new_$(New-Guid).xlsx"

            $script:workbook = Open-ExcelWorkbookHC -Path $path -Type 'Batch'
            Add-PlantRowHC -Workbook $workbook -FileName 'One.xml'
            Format-ExcelWorkbookHC -Workbook $workbook

            $script:table = $workbook.Sheet['plant'].Worksheet.Tables['plantBatchComputer']
        }

        AfterEach {
            Close-ExcelPackage $workbook.Package -NoSave
        }

        It 'creates the table' {
            $table | Should-BeTruthy
        }

        It 'covers the header row and the row that was written' {
            $table.TableXml.table.ref | Should-MatchString '^A2:[A-Z]+3$'
        }
    }
}